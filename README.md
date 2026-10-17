# ASAP Calculation Replica

A **multi-deal, formula-driven Excel reconciliation project**, currently **scaffolding only**. The shared Python package will eventually reproduce deal-specific ASAP periodic calculations with native Microsoft Excel formulas, then reconcile against SQL-exported ASAP values.

> **Security:** This repository was public when initialized. Do **not** commit proprietary ASAP scripts, loan-level CSV data, SQL benchmark extracts, payment inputs, credentials, or workbooks. Deal IDs are included only as directory identifiers. Before ever publishing real deal scripts or client data, obtain employer authorization and verify repository visibility and policies. Git ignore rules are not a substitute for access controls.

## Layout
- `src/asap_calc_replica/`: future shared CSV, AST/dependency analysis, formula generation, Excel COM and reconciliation implementations
- `deals/{deal_id}/`: deal-specific configuration, original source scripts **stored locally only**, and private field/runtime mappings
- `deals/_template/`: reusable safe example configuration and mappings
- `runs/{deal_id}/{YYYY-MM}/`: local inputs, SQL benchmark, Excel output, reports (ignored by Git)
- `docs/`: architecture, phased plan, Excel constraints and progress
- `tests/`: safe, synthetic tests

## Supported deal directories
AALC2201, AALC2202, AALC2401, AALC2403, AALC2404, AALC2501, AALC2502, AALC2503, AALC2601.

## Getting started (Windows / VS Code)
1. Clone the repository. Use Python 3.11+ and Microsoft Excel for Windows for eventual Excel COM execution.
2. Create a virtual environment and install `pip install -e ".[dev]"`.
3. Put each original periodic script in `deals/<DEAL>/asap_scripts/` **locally**. Files there are ignored by Git.
4. Keep deal-run inputs under `runs/<DEAL>/<YYYY-MM>/inputs/` and SQL results under `benchmark/`.
5. Review `AGENTS.md` and `docs/implementation_plan.md` before directing Copilot to start Phase 1.
6. The calculation CLI and workbook generation **are not implemented yet**. Do not interpret directories/placeholders as working features.

`python scripts/setup_deals.py --deal AALC2201 --month 2026-09` creates local run folders; `python -m pytest` runs only scaffold tests.
