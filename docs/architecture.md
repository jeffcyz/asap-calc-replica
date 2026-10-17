# Architecture

**Separation:** Shared infrastructure in `src/` versus Deal-specific authoritative scripts and mappings in `deals/`, versus monthly inputs and computed outputs in `runs/`.

Intended flow: load raw CSVs and supplied payment_inputs/period/runtime values → validate keys → inspect ASAP Python source and dependencies → compile ordered formula stages → create local-default Excel workbook using COM → recalculate → compare against independently exported SQL benchmark.

Raw worksheets: AssetSetup, AssetPeriodic, AssetPeriodicTransaction, optional AssetPeriodicLegacy, PaymentInputs, PeriodConfig. Derived sheets: TransactionCalc, PortfolioCalc, AssetPeriodicCalc, ASAP_SQL, Comparison. Additional intermediate stages allowed when needed to preserve sequential updates.

Join keys must be discovered from real exported schemas; use DealId/AssetId/period where supported. Do not assume transaction EffectiveDate equals PaymentDate or that AssetName is unique. Detect whether Other31–Other60 belong to AssetPeriodicLegacy. Source calculations may also require prior periods, deal critical dates and index values.

AALC2201 is the first vertical slice. Additional deals use the same engine but independent field lineage, branching and configuration. No generic IronPython-to-Excel compiler is promised.

**Current status:** Design only; modules are unimplemented.
