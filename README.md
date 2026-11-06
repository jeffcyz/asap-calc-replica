# ASAP-SQL-Remap

Reusable fail-closed XLSB migration tool for multiple ASAP deal models. This does NOT regenerate proof models or need the old calculation repository.

## Behavior

Read versioned XLSB from inbox/<DealName>/, inspect ASAP_SQL and Query!tblASAP_SQL_DB, verify mappings and strict row-data equality, rewrite supported full-column Excel formula references, remove ASAP_SQL, apply/verify the Internal label, save new versioned XLSB to output/<DealName>/, and write local reports/history. The original input is never changed.

## Windows setup

Requires Windows desktop Microsoft Excel, signed-in Microsoft 365 account with authorized Internal labeling policy, Python 3.10+, and `pip install pywin32`.

Set approved environment variables locally: ASAP_INTERNAL_LABEL_ID and (when starting unlabeled) ASAP_INTERNAL_SITE_ID. Never commit tenant policy identifiers. A non-Internal existing label is not downgraded.

Commands:

    python tools/remap_proof.py --deal AALC2501
    python tools/remap_proof.py --all
    python tools/remap_proof.py --file inbox/AALC2201/proof_before.xlsb

## Versioning examples

- proof_20261030_165600_001-v05.xlsb -> proof_20261030_165600_001-v06.xlsb
- proof-2502_v0.0.1_release.xlsb -> proof-2502_v0.0.2_release.xlsb
- proof_20261105_050502_601.xlsb -> proof_20261105_050502_601_v01.xlsb

## Scope and safety

The current implementation handles source references consisting of one complete worksheet column (e.g., ASAP_SQL!$H:$H), including SUMIFS/COUNTIFS/INDEX/MATCH expressions. It refuses finite/direct-cell ranges, INDIRECT, OFFSET, arrays, ambiguous mappings, and any detected non-cell workbook references. It requires matching headers and strictly identical row-order data across each affected source/target column; different SQL sort/grain or dates may cause a safe failure.

Exact header mappings can be specified per deal in config/remap_rules.json under mapping_overrides, together with a verified header_row_overrides entry. A mapping override is NOT a waiver of data equivalence. The runtime generates state/processing_history.jsonl automatically.

Validation compares formula definitions, computed results, and worksheet cell values before/after native Excel recalculation; it also checks sheets, tables and the persisted label after reopening. This is not a pixel-level or binary-perfect comparison of every workbook object. Before production use, independently validate a representative proof for each deal family and examine Excel Power Query connections, formatting, and financial tie-outs.

## Public repository warning

This repository is PUBLIC. The XLSB names shown under inbox/ are local-only examples, not files committed to GitHub. Do not upload internal model data, generated reports, credentials, label IDs or other corporate material. Files in inbox/, output/, reports/ and state/ are gitignored except for empty .gitkeep markers.

Replacing the main branch contents does NOT remove old commits from Git history. Follow company policy if a history purge is required.
