# Calculation dependency analysis

**Not populated yet**. Phase 1 reads the actual, authorized AALC2201 script from the ignored per-deal \`asap_scripts/\` directory.

Capture output field and source expressions/line numbers; assignment chronology and intermediate states; transaction filters/ordering; portfolio aggregation/waterfall sequencing; current and previous periods; critical_dates, payment_inputs and index dependencies; input table/SQL keys; runtime collection behavior; initialization and side effects; data availability; evidence and Excel equivalence status.

Use \`docs/runtime_dependency_inventory.md\` for ASAP-specific symbol contracts and \`docs/calculation_coverage.md\` for A/B/C/D readiness. Keep deal-specific results under ignored \`runs/<deal>/<month>/reports/\` rather than in this public file.

Do not substitute final ASAP SQL benchmark values as interim formula inputs, or assume an undocumented ASAP function has ordinary Python semantics.
