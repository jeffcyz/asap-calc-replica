# Agent instructions: ASAP Calculation Replica

## Task discipline
- This repository starts as a **scaffold**, not a working calculation engine.
- Work in approved phases from `docs/implementation_plan.md`; update `docs/progress.md` at the end of each phase.
- Before writing formulas, analyze the **complete original** ASAP script and persist field lineage, sequential mutation and portfolio dependency order.
- Do not invent missing data, implicit join keys, period mappings, runtime variables or SQL values.
- Original scripts are read-only authoritative logic; never rewrite, format or overwrite them.
- Do not assume Deal-specific scripts are equivalent. Share infrastructure, not unproven business logic.
- Prefer one end-to-end verified deal (AALC2201) before adding more.
- Never claim formula validation until real Excel recalculation and benchmark comparison succeed.

## Excel rules
- Use INDEX/MATCH exact-match lookups, not XLOOKUP/VLOOKUP/XMATCH.
- Produce **native Excel formulas**; Python builds and validates workbooks but does not replace formula output with hardcoded calculations.
- On Windows, create a workbook using local Microsoft Excel `Workbooks.Add()` via COM to inherit New Workbook defaults.
- Never set/autofit column widths, decorate headers or apply custom fills, borders, table styles, alignments, row heights or zoom.
- Actual date values use local built-in Short Date; monetary amounts use Excel built-in Comma Style; all others General.
- The Comparison table consists of identity columns + `Excel`, `ASAP SQL`, `Difference`; signed Difference = Excel - ASAP SQL.
- SQL benchmark values must never feed the Excel calculation graph.
- Missing match is not zero; fail visibly and report unresolved dependencies.
- Exact-match keys require validated uniqueness. AssetName alone is not assumed unique.
- Preserve transaction ordering when it changes semantics, previous-period carryforward, numeric precision, day counts, period exceptions and portfolio waterfall sequence.

## Source paths
- Deal source: `deals/{deal_id}/asap_scripts/` (local files, ignored by Git).
- Shared engine: `src/asap_calc_replica/`.
- Mappings: `deals/{deal_id}/mappings/` (local copies, ignored by Git).
- Period data: `runs/{deal_id}/{YYYY-MM}/` (ignored by Git).

## Security and publication
- Repository is initially public. **Do not stage or push** scripts, actual client/loan data, deal-specific mappings with private terms, payment_inputs, SQL benchmark, workbooks, credentials or confidential debug logs.
- Keep examples synthetic. Check `git status` and `git diff --cached` before any push. Confirm the user's organizational data policy and repository visibility before considering exceptions.
