# Phased implementation plan

Each phase requires a review checkpoint; do not proceed automatically.

## Phase 0 — Scaffolding
Directory structure, public-safe templates, documentation, packaging, basic synthetic tests, local run folder helper.

## Phase 1 — AALC2201 dependency inventory
Read complete local ASAP script and schemas; trace every field assignment including mutation order, cross-period dependencies, transaction loops, deal critical dates, payment_inputs, index services and portfolio aggregation. Deliver dependencies report, missing inputs and canonical key design. **No Excel workbook yet.**

## Phase 2 — Excel formula vertical slice
Load synthetic/sample authorized data; test exact INDEX/MATCH and transaction aggregation; implement one asset and a small group of fields with intermediate stages; generate workbook through Windows Excel COM with required formatting. Validate by recalculating.

## Phase 3 — Full AALC2201 computation
Implement all resolvable asset/transaction/portfolio waterfall and carryforward stages with explicit exceptions. Report unsupported formulas; no invented values.

## Phase 4 — SQL reconciliation
Read independently exported SQL field values; output identity columns, Excel, ASAP SQL, Difference (Excel minus SQL); verify keys, row coverage and Excel recalc.

## Phase 5 — Second-deal architecture test
Integrate AALC2404 to test reusability. Adapt shared *infrastructure* only; do not presume business formulas are common.

## Phase 6 — Remaining deals and batch runner
Complete mappings for the remaining seven deals and add month-specific batch orchestration.

Acceptance criteria per phase: reproducible tests, source/field lineage, documented missing inputs, no confidential data in Git, and explicit status in progress.md.
