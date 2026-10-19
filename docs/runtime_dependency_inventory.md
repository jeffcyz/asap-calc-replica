# Runtime dependency inventory — specification (not yet populated)

**Scope:** Analyze AALC2201's actual local script in Phase 1. Do not infer ASAP runtime semantics from Python method names.

For every ASAP-specific symbol record:

| Column | Meaning |
|---|---|
| deal_id | Deal owning the reference |
| symbol | Exact ASAP expression / method |
| source_file / lines | File and occurrence range |
| kind | Source field, runtime method, collection, historic state, lifecycle or external value |
| evidence_reference | Exact authorized KB path / section / version |
| semantics_documented | Has verified meaning? |
| side_effects_resolved | Known read/write/state behavior and execution timing? |
| requires_runtime_mapping | Must be translated into an SQL/Excel representation? |
| runtime_mapping_documented | Is that representation justified by evidence? |
| required_inputs | Named input fields and runtime values |
| available_inputs | Which are available in the selected run snapshot |
| unresolved_semantics | Explicit unresolved questions |

Distinguish:
1. **Semantic gap** — unclear function/filter/initialization/side effect/rounding/period behavior.
2. **Data gap** — known required opening values, transaction data, index values, historical records or inputs not included in snapshots.
3. **Translation gap** — the semantics are understood but native Excel equivalence has not been demonstrated.

Inventory should cover collection membership, current/previous period access, payment_inputs, critical_dates, index access, transactions, portfolio waterfall phases and source-to-SQL schema mappings.

Do not embed private inventory data in public Git-tracked docs. Export real inventory as ignored files under the Git-tracked `runs/<deal>/<month>/reports/` directory.
