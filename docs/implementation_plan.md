# Phased implementation plan

Each phase has a review checkpoint; do not proceed automatically. **Do not begin formula generation before runtime semantics and input sufficiency are assessed.**

## Phase 0 — Multi-deal scaffold (complete)
Project structure, safe templates, packaging and synthetic scaffold tests.

## Phase 0.5 — Runtime semantics architecture (complete)
Introduce evidence-backed ASAP Runtime Contracts, dependency registry, explicit period context, collection semantic gates, and provisional A/B/C/D classifier. Add local-only KB conventions and updated documentation. **No real method semantics or calculation formulas are asserted.**

## Phase 1 — AALC2201 ASAP semantics and input sufficiency assessment (next)
Read authorized complete local script and relevant schemas. Inventory objects, dynamic methods, collection filters, lifecycle/side effects, sequential assignments, periods, index and historical state. Link each unknown to **specific** authorized KB sections; distinguish semantic gaps, data gaps and Excel equivalence gaps. Build ordered field dependency graph and classify *every calculated output field* A/B/C/D. Persist private deal-specific details as **ignored report files** inside the Git-tracked `runs/AALC2201/<month>/reports/` directory.

Deliverables: runtime dependency inventory, required input snapshots, missing inputs report, field lineage, coverage assessment and architectural decisions requiring approval. **Do not generate Excel yet.**

## Phase 2 — Excel formula vertical slice
Once Phase 1 is reviewed, implement one real/authorized or synthetic asset across a few fields. Validate exact INDEX/MATCH, transaction aggregation, staged portfolio dependencies and local-default Excel COM formatting. Recalculate and compare against SQL only where data and semantics are complete.

## Phase 3 — Full AALC2201 formula model
Implement all A/B fields and C fields after missing inputs are supplied. Keep D fields unresolved; no guessed mappings. Preserve mutation order, transaction loops, carryforward, portfolio allocation and period exceptions.

## Phase 4 — SQL reconciliation
ASAP SQL benchmark: independent source of expected output. Long-form identity + Excel + ASAP SQL + Difference (Excel - ASAP SQL), no invented zeros, verified joins, matching branches and actual Excel recalculation.

## Phase 5 — Second-deal architecture test
Integrate AALC2404; determine which runtime contracts and shared infrastructure generalize while retaining deal-specific scripts/field definitions.

## Phase 6 — Remaining deals and batch execution
Integrate remaining seven deals, per-month snapshots and batch orchestration.

Acceptance criteria for each phase: accurate scope/status, automated synthetic tests where possible, source evidence and field lineage, explicit unresolved dependencies, no confidential data in public Git, updated `docs/progress.md`.
