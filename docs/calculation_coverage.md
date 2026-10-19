# Calculation coverage model — A / B / C / D

The semantic resolver exposes a **conservative, provisional readiness classification**. It does not inspect scripts yet, execute ASAP, run Excel formulas, or assert numeric correctness.

| Level | Criteria | Next step |
|---|---|---|
| A | Expression and Excel equivalence understood; all inputs available; no unresolved ASAP mapping needed | Implement formula, independently validate |
| B | As A, but requires a documented ASAP-to-Excel/SQL runtime mapping | Implement mapping/formula and independently validate |
| C | Expression, semantics and translation are understood, but one or more required inputs are absent | Export named inputs/snapshots/history |
| D | Undocumented behavior, unknown side effects, missing source contract or uncertain Excel equivalence | Obtain KB evidence / investigate before translation |

**Precedence:** D > C > B > A. An output field with both undocumented semantics and missing data stays D, with the data gap also reported. Do not use a numeric match against one SQL period to claim semantic equivalence.

Required field assessment metadata: output field, source expression/lines, required symbols, contracts/evidence, expression semantic support, Excel equivalence support, runtime mapping need, required and available inputs, classification, unresolved questions.

Real reports should be written to `runs/<deal>/<month>/reports/calculation_coverage.csv` and are excluded from Git. The committed module `analysis/semantic_resolver.py` supports **metadata-only** classification; automatically extracting source and input evidence is future Phase 1 work.

SQL result columns are **benchmarks only**, never substitutes for unresolved calculation inputs.
