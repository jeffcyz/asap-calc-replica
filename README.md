# ASAP Calculation Replica

A multi-deal project for **semantically justified** ASAP periodic calculation reconstruction using **native Excel formulas** and independent SQL reconciliation.

> **Current status:** repository scaffolding and a conservative ASAP runtime **contract/coverage metadata layer** are present. No real ASAP deal-script analysis, Excel formula engine, or SQL reconciliation is implemented. Original ASAP/IronPython scripts cannot be assumed executable outside ASAP.

> **Security:** This repository is **public**. Do not commit proprietary ASAP scripts, internal Knowledge Base content, loan-level CSV data, SQL extracts, payment inputs, credentials or output workbooks. Authorized local files live in ignored paths. Check employer/AI tooling policy before making them available to Copilot.

## Repository layout

- `src/asap_calc_replica/runtime/`: documented host-object contracts, explicit dependency registry, collection semantics gates and period context (**not an ASAP emulator**).
- `src/asap_calc_replica/analysis/`: future script/dependency analyzers and a **metadata-only** A/B/C/D semantic readiness assessor.
- `src/asap_calc_replica/calculation/`, `excel/`, `reconciliation/`: planned shared calculation/formula, COM and SQL comparison engines.
- `deals/<deal_id>/`: deal configuration, original scripts (local-only), runtime/field mapping (local-only).
- `knowledge/README.md`: local-only ASAP Knowledge Base intake guidance. Put authorized KB files under `knowledge/asap/` (Git-ignored), optionally linking approved ASAP-Agent knowledge locally.
- `runs/<deal_id>/<YYYY-MM>/`: **tracked folder structure** via `.gitkeep`; data inputs, SQL benchmarks, outputs and reports remain ignored.
- `docs/`: architecture, runtime gaps, coverage, Excel requirements, phased implementation and progress.
- `tests/`: synthetic scaffold and runtime contract tests.

Initial deals: AALC2201, AALC2202, AALC2401, AALC2403, AALC2404, AALC2501, AALC2502, AALC2503, AALC2601.

## Local setup (Windows + Excel)

1. Clone the repository and open in VS Code. Use Python 3.11+.
2. Create a virtual environment; install `pip install -e ".[dev]"`.
3. Put original ASAP source files in `deals/<DEAL>/asap_scripts/` locally. Do not commit them.
4. Put only authorized KB files under `knowledge/asap/` locally. Runtime contracts need source evidence, not guesses.
5. Keep monthly CSV, period, payment_inputs, historical snapshots and SQL benchmark under `runs/<DEAL>/<YYYY-MM>/`.
6. All nine `2026-10` run structures are committed. For a future month use `python scripts/setup_deals.py --all --month 2026-11` or `--deal AALC2201` for one deal; `.gitkeep` markers make these directories visible in Git. `python -m pytest` runs synthetic unit tests. No real calculation commands exist yet.
7. Review `AGENTS.md` and **Phase 1** of `docs/implementation_plan.md` before directing Copilot to analyze AALC2201.

See `docs/run_directory_layout.md` for the complete October folder layout and ignore policy.

## Formula and output constraints (once implemented)

Exact INDEX/MATCH, staged asset/transaction/portfolio calculations, local default Windows Excel workbook through COM, no adjusted column widths or decoration, Short Date for dates, Comma Style for money, General otherwise. Comparison has identity columns, Excel, ASAP SQL, Difference = Excel - ASAP SQL. SQL benchmark outputs are never formula inputs.

## Read next

- `docs/architecture.md`: design boundaries.
- `docs/knowledge_base_integration.md`: evidence and privacy requirements.
- `docs/runtime_dependency_inventory.md`: runtime inventory schema.
- `docs/calculation_coverage.md`: A/B/C/D assessment.
- `docs/implementation_plan.md`: phased roadmap.
