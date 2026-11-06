"""Fail-closed XLSB remapping tool. Runs only on Windows with Microsoft Excel."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HISTORY = ROOT / "state" / "processing_history.jsonl"
REF = re.compile(r"(?i)(?<![A-Za-z0-9_])(?:'ASAP_SQL'|ASAP_SQL)!")
COLUMN = re.compile(r"(?i)(?<![A-Za-z0-9_])(?:'ASAP_SQL'|ASAP_SQL)!\$?([A-Z]{1,3}):\$?([A-Z]{1,3})(?![A-Za-z0-9_])")
VERSION = re.compile(r"(?i)[-_]v(\d+(?:\.\d+)*)(?=$|[-_.])")
FORMULAS = -4123
LAST_UP = -4162

class Blocked(RuntimeError):
    pass

def versioned_name(filename):
    p = Path(filename)
    if p.suffix.lower() != ".xlsb":
        raise Blocked("Expected XLSB input")
    matches = list(VERSION.finditer(p.stem))
    if len(matches) > 1:
        raise Blocked("Ambiguous version tokens")
    if not matches:
        return p.stem + "_v01" + p.suffix
    m = matches[0]
    parts = m.group(1).split(".")
    parts[-1] = str(int(parts[-1]) + 1).zfill(len(parts[-1]))
    return p.stem[:m.start(1)] + ".".join(parts) + p.stem[m.end(1):] + p.suffix

def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def known_hash(value):
    if not HISTORY.exists():
        return False
    for line in HISTORY.read_text(encoding="utf-8").splitlines():
        try:
            item = json.loads(line)
        except ValueError:
            continue
        if item.get("sha256") == value and item.get("status") == "PASS":
            return True
    return False

def record_history(record):
    HISTORY.parent.mkdir(parents=True, exist_ok=True)
    with HISTORY.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")

def get_sheet(book, name):
    try:
        return book.Worksheets(name)
    except Exception as exc:
        raise Blocked("Missing worksheet " + name) from exc

def get_table(book, cfg):
    query = get_sheet(book, cfg["target_worksheet"])
    try:
        table = query.ListObjects(cfg["target_table"])
    except Exception as exc:
        raise Blocked("Missing real Excel table Query!tblASAP_SQL_DB") from exc
    headers = [str(c.Name) for c in table.ListColumns]
    if len(headers) != len({h.casefold() for h in headers}):
        raise Blocked("Ambiguous SQL table headers")
    return table, headers

def formula_cells(sheet):
    try:
        areas = sheet.UsedRange.SpecialCells(FORMULAS).Areas
    except Exception:
        return
    for area in areas:
        for cell in area.Cells:
            yield cell

def dependencies(book, source):
    formulas, blocking = [], []
    for sheet in book.Worksheets:
        if sheet.Name == source:
            continue
        for cell in formula_cells(sheet):
            expression = str(cell.Formula)
            if REF.search(expression):
                formulas.append((sheet.Name, cell.Address, expression))
        try:
            for condition in sheet.UsedRange.FormatConditions:
                for attr in ("Formula1", "Formula2"):
                    try:
                        if REF.search(str(getattr(condition, attr))):
                            blocking.append(f"Conditional formatting: {sheet.Name}/{attr}")
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            validated = sheet.UsedRange.SpecialCells(-4174)
            for cell in validated.Cells:
                for attr in ("Formula1", "Formula2"):
                    try:
                        if REF.search(str(getattr(cell.Validation, attr))):
                            blocking.append(f"Data validation: {sheet.Name}!{cell.Address}")
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            for chart in sheet.ChartObjects():
                for series in chart.Chart.SeriesCollection():
                    if REF.search(str(series.Formula)):
                        blocking.append(f"Chart series: {sheet.Name}/{chart.Name}")
        except Exception:
            pass
    for name in book.Names:
        if REF.search(str(name.RefersTo)):
            blocking.append("Defined name: " + str(name.Name))
    try:
        for query in book.Queries:
            if "ASAP_SQL" in str(query.Formula).upper():
                blocking.append("Power Query: " + str(query.Name))
    except Exception:
        blocking.append("Cannot inspect Power Query definitions")
    try:
        for connection in book.Connections:
            if "ASAP_SQL" in str(connection.Description).upper():
                blocking.append("Workbook connection: " + str(connection.Name))
    except Exception:
        blocking.append("Cannot enumerate workbook connections")
    return formulas, blocking

def column_index(letters):
    number = 0
    for ch in letters:
        number = number * 26 + ord(ch) - 64
    if not 1 <= number <= 16384:
        raise Blocked("Invalid Excel column " + letters)
    return number

def source_columns(formulas):
    needed = set()
    for name, address, expr in formulas:
        if re.search(r"\b(?:INDIRECT|OFFSET)\s*\(", expr, re.I):
            raise Blocked(f"Dynamic reference requires review: {name}!{address}")
        def collect(match):
            first, last = match.group(1).upper(), match.group(2).upper()
            if first != last:
                raise Blocked("Unsupported multi-column reference " + match.group(0))
            needed.add(first)
            return "__COLUMN__"
        remaining = COLUMN.sub(collect, expr)
        if REF.search(remaining):
            raise Blocked(f"Unsupported finite/direct reference: {name}!{address}")
    return sorted(needed)

def mapping_for(old, headers, columns, overrides, chosen_row):
    available = {h.casefold(): h for h in headers}
    candidates = []
    rows = [int(chosen_row)] if chosen_row else range(1, 21)
    for row in rows:
        mapping = {}
        for col in columns:
            if col in overrides:
                requested = str(overrides[col]).casefold()
            else:
                value = old.Cells(row, column_index(col)).Value2
                requested = str(value).strip().casefold() if value is not None else ""
            if requested not in available:
                break
            mapping[col] = available[requested]
        if len(mapping) == len(columns):
            candidates.append((row, mapping))
    if len(candidates) != 1:
        raise Blocked("Cannot uniquely verify old column headers; set verified mapping and header_row_overrides")
    return candidates[0]

def values_of(region):
    if region is None:
        return []
    data = region.Value2
    if isinstance(data, tuple):
        return [r[0] if isinstance(r, tuple) else r for r in data]
    return [data]

def verify_rows(old, table, row, mapping, strict):
    if not strict:
        raise Blocked("Unsafe strict_source_row_match=false is not supported")
    if table.DataBodyRange is None:
        raise Blocked("SQL table has no data rows")
    total = int(table.ListRows.Count)
    for col, header in mapping.items():
        index = column_index(col)
        last = int(old.Cells(old.Rows.Count, index).End(LAST_UP).Row)
        if last != row + total:
            raise Blocked(f"Row-count mismatch for {col}/{header}; ASAP_SQL last row {last}, SQL rows {total}")
        original = old.Range(old.Cells(row + 1, index), old.Cells(row + total, index))
        a, b = values_of(original), values_of(table.ListColumns(header).DataBodyRange)
        if a != b:
            first = next((i for i, (x, y) in enumerate(zip(a, b), 1) if x != y), 0)
            raise Blocked(f"Source data differs for {col}/{header} at row offset {first}")

def bracket(header):
    escaped = header.replace("'", "''").replace("[", "'[").replace("]", "']").replace("#", "'#")
    return "[" + escaped + "]"

def remap_formula(expr, mapping, table):
    def swap(match):
        a, b = match.group(1).upper(), match.group(2).upper()
        if a != b or a not in mapping:
            raise Blocked("Unmapped reference " + match.group(0))
        return table + bracket(mapping[a])
    changed = COLUMN.sub(swap, expr)
    if REF.search(changed):
        raise Blocked("Residual ASAP_SQL reference")
    return changed

def label(book):
    value = book.SensitivityLabel.GetLabel()
    return str(value.LabelId or "").strip().lower(), str(value.SiteId or "").strip()

def set_internal(book, cfg):
    expected = os.getenv(cfg["label_id_env"], "").lower().strip()
    if not re.fullmatch(r"[0-9a-f-]{36}", expected):
        raise Blocked("Set approved Internal LabelId in " + cfg["label_id_env"])
    current, prior_site = label(book)
    if current and current != expected:
        raise Blocked("Different existing sensitivity label: refusal to downgrade/change it")
    if current == expected:
        return expected
    site_id = os.getenv(cfg["label_site_id_env"], "").strip() or prior_site
    if not re.fullmatch(r"[0-9a-fA-F-]{36}", site_id):
        raise Blocked("An approved tenant SiteId is required in " + cfg["label_site_id_env"])
    info = book.SensitivityLabel.CreateLabelInfo()
    info.LabelName = cfg["label_name"]
    info.LabelId = expected
    info.SiteId = site_id
    info.IsEnabled = True
    info.AssignmentMethod = 1  # Office MsoAssignmentMethod.PRIVILEGED
    book.SensitivityLabel.SetLabel(info, info)
    if label(book)[0] != expected:
        raise Blocked("Internal label assignment did not succeed")
    return expected

def snapshot(book, omitted):
    sheet_names, tables, formulas, values = [], {}, {}, {}
    for sheet in book.Worksheets:
        if sheet.Name == omitted:
            continue
        sheet_names.append(sheet.Name)
        tables[sheet.Name] = [(t.Name, tuple(c.Name for c in t.ListColumns), int(t.ListRows.Count)) for t in sheet.ListObjects]
        for cell in formula_cells(sheet):
            formulas[sheet.Name + "!" + cell.Address] = (str(cell.Formula), repr(cell.Value2))
        region = sheet.UsedRange
        values[sheet.Name] = (str(region.Address), repr(region.Value2))
    return sheet_names, tables, formulas, values

def compare(before, after, replacements):
    sheets1, tables1, formulas1, values1 = before
    sheets2, tables2, formulas2, values2 = after
    if sheets1 != sheets2 or tables1 != tables2:
        raise Blocked("Worksheet order or SQL table metadata changed")
    if formulas1.keys() != formulas2.keys():
        raise Blocked("Formula-cell locations changed")
    for where, (expr, result) in formulas1.items():
        changed_expr, new_result = formulas2[where]
        if changed_expr != replacements.get(where, expr):
            raise Blocked("Unexpected formula change at " + where)
        if result != new_result:
            raise Blocked("Calculated value changed at " + where)
    if values1 != values2:
        raise Blocked("Workbook cell values changed unexpectedly")

def run_file(path, deal, cfg):
    path = path.resolve()
    fingerprint = sha256(path)
    if known_hash(fingerprint):
        print("SKIP already processed:", path.name)
        return True
    output = ROOT / "output" / deal / versioned_name(path.name)
    report = ROOT / "reports" / deal / (output.stem + "_remap_report.md")
    output.parent.mkdir(parents=True, exist_ok=True)
    report.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        print("FAIL output collision:", output, file=sys.stderr)
        return False
    staging = output.with_name("." + output.stem + ".staging.xlsb")
    if staging.exists():
        print("FAIL staging file exists:", staging, file=sys.stderr)
        return False
    audit = {"time": datetime.now(timezone.utc).isoformat(), "deal": deal,
             "input": path.name, "sha256": fingerprint, "proposed_output": output.name,
             "formula_count": 0, "status": "FAIL"}
    excel, book = None, None
    try:
        import pythoncom
        from win32com.client import DispatchEx
        pythoncom.CoInitialize()
        shutil.copy2(path, staging)
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        book = excel.Workbooks.Open(str(staging), UpdateLinks=0, ReadOnly=False)
        try:
            old = book.Worksheets(cfg["source_worksheet"])
        except Exception:
            audit["status"] = "ALREADY_MIGRATED"
            print("SKIP source sheet already absent:", path.name)
            return True
        table, headers = get_table(book, cfg)
        formulas, blockers = dependencies(book, cfg["source_worksheet"])
        if blockers:
            raise Blocked("Workbook-object dependencies require manual review: " + "; ".join(blockers[:20]))
        columns = source_columns(formulas)
        overrides = cfg.get("mapping_overrides", {}).get(deal, {})
        explicit_row = cfg.get("header_row_overrides", {}).get(deal)
        header_row, mapping = mapping_for(old, headers, columns, overrides, explicit_row) if columns else (1, {})
        if columns:
            verify_rows(old, table, header_row, mapping, cfg.get("strict_source_row_match", True))
        excel.CalculateFullRebuild()
        baseline = snapshot(book, cfg["source_worksheet"])
        replacements = {}
        for name, address, expr in formulas:
            translated = remap_formula(expr, mapping, cfg["target_table"])
            if translated != expr:
                cell = book.Worksheets(name).Range(address)
                if bool(cell.HasArray):
                    raise Blocked("Array formula needs manual review at " + name + "!" + address)
                cell.Formula = translated
                replacements[name + "!" + address] = translated
        unresolved, extras = dependencies(book, cfg["source_worksheet"])
        if unresolved or extras:
            raise Blocked("Unresolved workbook references: " + str(len(unresolved)) + " formulas; " + str(extras[:10]))
        excel.CalculateFullRebuild()
        compare(baseline, snapshot(book, cfg["source_worksheet"]), replacements)
        book.Worksheets(cfg["source_worksheet"]).Delete()
        set_internal(book, cfg)
        book.Save()
        book.Close(SaveChanges=False)
        book = excel.Workbooks.Open(str(staging), UpdateLinks=0, ReadOnly=True)
        if label(book)[0] != os.getenv(cfg["label_id_env"], "").strip().lower():
            raise Blocked("Internal sensitivity label failed save/reopen verification")
        if any(s.Name == cfg["source_worksheet"] for s in book.Worksheets):
            raise Blocked("ASAP_SQL still present in saved file")
        get_table(book, cfg)
        remaining, extras = dependencies(book, cfg["source_worksheet"])
        if remaining or extras:
            raise Blocked("ASAP_SQL dependency still present after save")
        excel.CalculateFullRebuild()
        compare(baseline, snapshot(book, cfg["source_worksheet"]), replacements)
        book.Close(SaveChanges=False)
        book = None
        os.replace(staging, output)
        audit.update(status="PASS", output=output.name, formula_count=len(replacements), label="Internal")
        report.write_text("# SQL Remap — PASS\n\n" +
                          f"- Input: {path.name}\n- Output: {output.name}\n- Deal: {deal}\n" +
                          f"- Migrated formulas: {len(replacements)}\n- Header row: {header_row}\n" +
                          f"- Verified mapping: {json.dumps(mapping, ensure_ascii=False)}\n" +
                          "- Formula results and full cell values: matched\n" +
                          "- Query table: retained; ASAP_SQL: removed\n" +
                          "- Internal label: persisted after reopen\n" +
                          "- Limitation: binary workbook identity and all formatting objects are not exhaustively compared.\n",
                          encoding="utf-8")
        print("PASS:", path.name, "->", output)
        return True
    except Exception as exc:
        audit["error"] = str(exc)
        report.write_text("# SQL Remap — FAIL\n\n" + str(exc) + "\n", encoding="utf-8")
        print("FAIL:", path.name, str(exc), file=sys.stderr)
        return False
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        if staging.exists():
            try:
                staging.unlink()
            except OSError:
                pass
        record_history(audit)

def main():
    parser = argparse.ArgumentParser(description="Fail-closed XLSB remapping for ASAP proofs")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--deal")
    group.add_argument("--all", action="store_true")
    group.add_argument("--file", type=Path)
    opts = parser.parse_args()
    cfg = json.loads((ROOT / "config" / "remap_rules.json").read_text(encoding="utf-8"))
    if opts.file:
        tasks = [(opts.file, opts.file.parent.name)]
    else:
        base = ROOT / "inbox"
        path = base / opts.deal if opts.deal else base
        mask = "*.xlsb" if opts.deal else "*/*.xlsb"
        tasks = [(f, f.parent.name) for f in sorted(path.glob(mask))]
    if not tasks:
        print("No XLSB input files in selected scope.")
        return 0
    return 0 if all(run_file(path, deal, cfg) for path, deal in tasks) else 1

if __name__ == "__main__":
    sys.exit(main())
