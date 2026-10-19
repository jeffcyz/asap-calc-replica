# Architecture

The shared engine in `src/` is separate from per-deal, read-only ASAP sources and mappings in `deals/` and per-month input/SQL/output snapshots in `runs/`. The `runs/` hierarchy is **Git-tracked via `.gitkeep` markers**, while data files inside remain ignored.

## Important constraint: ASAP is not locally executable

The original ASAP scripts rely on host-provided IronPython objects, dynamic data access, execution lifecycle and possibly hidden state. We **do not** attempt to emulate the entire ASAP runtime. The original script is the authoritative source of *visible* business expressions; the knowledge base is needed for host semantics, and authorized data snapshots for concrete observed values.

## Intended pipeline

1. Local deal-specific ASAP scripts + local-only `knowledge/asap/` + explicit input snapshots (`period`, `payment_inputs`, raw SQL tables, date/index/history values).
2. **Semantic inventory**: locate every host-specific field, method, collection, mutation and lifecycle dependency.
3. **Evidence-backed runtime contracts**: describe proven behavior, mapping, required inputs and side effects; refuse undocumented assumptions.
4. **Dependency/sufficiency assessment**: derive ordered field graph and categorize A/B/C/D (see `docs/calculation_coverage.md`).
5. **Excel formula stages**: preserve transaction order, intermediate mutations, asset and portfolio calculations and carryforward with native formulas.
6. Use Windows Excel COM to create a local-default workbook, recalculate and compare against independently exported ASAP SQL benchmarks.

Unknown methods, hidden initialization, period mapping and undocumented filtering result in an unresolved field, **not a guessed formula**.

## Source and workbook boundaries

Raw worksheets: AssetSetup, AssetPeriodic, AssetPeriodicTransaction, optional AssetPeriodicLegacy, PaymentInputs, PeriodConfig. Derived: TransactionCalc, PortfolioCalc, AssetPeriodicCalc, ASAP_SQL, Comparison. Additional intermediate stages allowed to avoid circular references.

Validate composite SQL joins and matching period semantics; do not assume AssetName uniqueness or EffectiveDate = PaymentDate. History, deal critical dates, index values and runtime inputs are first-class dependencies. SQL benchmark outputs never become computation inputs.

## New shared packages

- `runtime/object_contracts.py`: evidence-backed metadata, not execution.
- `runtime/dependency_registry.py`: explicit runtime dependency registry.
- `runtime/collection_semantics.py`: fail-closed validation of collection contracts, not collection implementation.
- `runtime/period_context.py`: explicit period and inputs; no date-derived period guesses.
- `analysis/semantic_resolver.py`: provisional A/B/C/D assessment based on supplied evidence, **not a parser**.

AALC2201 is first. Follow with AALC2404 to test cross-deal reuse. **Current status:** runtime contract scaffold and documented design only; no ASAP source analysis, native Excel calculation engine or reconciliation yet.
