# Agent instructions: ASAP Calculation Replica

## Fundamental constraint
- ASAP/IronPython source scripts are **not standalone executables** in local Python. ASAP supplies objects, collection behavior, data access, execution lifecycle and potentially hidden state. **Do not build a complete ASAP emulator.**
- Target is **observable calculation semantic equivalence in native Excel**, not direct execution or syntax-only translation.
- Never assume behavior from a method name or one month's matching SQL result. Unknown ASAP semantics remain unresolved until supported by authoritative evidence.
- Use \`docs/architecture.md\`, \`docs/knowledge_base_integration.md\`, \`docs/runtime_dependency_inventory.md\` and \`docs/calculation_coverage.md\` before formula development.

## Approved phase discipline
- Work only on the user's approved phase in \`docs/implementation_plan.md\`; record progress in \`docs/progress.md\`.
- Phase 1 first: read full *local authorized* source and schemas; identify runtime APIs, implicit initialization, side effects, ordering, historical state and required inputs.
- For every output field, distinguish (i) semantic gap, (ii) missing data, (iii) Excel translation gap. Report A/B/C/D readiness, with D taking priority over C.
- The \`analysis/semantic_resolver.py\` is a **metadata classifier only**, not a script parser, transpiler, ASAP runtime, proof of formula correctness, or calculator.
- Do not generate formulas for unresolved operations; no guessed defaults, zeros, implicit keys, arbitrary rounding or silent branch dropping.
- The original scripts are read-only authority for visible business expressions; authorized KB contracts explain their host behavior. Neither replaces observed monthly input data.
- Reuse proven runtime contracts across deals, but do not assume deal-specific formulas are equivalent. Begin with AALC2201; then test AALC2404.

## Local sources
- Deal scripts: \`deals/{deal_id}/asap_scripts/\` (read-only, Git-ignored).
- Deal mappings: \`deals/{deal_id}/mappings/\` (private, Git-ignored).
- ASAP KB: \`knowledge/asap/\` (local-only, Git-ignored), optionally approved local ASAP-Agent reference.
- Raw data, payment_inputs, period and SQL benchmark: \`runs/{deal_id}/{YYYY-MM}/\` (Git-ignored).
- Engine: \`src/asap_calc_replica/\`.

## Excel / validation rules
- Native formulas; exact-match INDEX/MATCH instead of XLOOKUP/VLOOKUP/XMATCH; SUMIFS permitted for aggregation.
- Preserve transaction order if material, prior-period carryforward, mutation order, portfolio waterfall, period exceptions, date arithmetic, runtime rounding/null semantics.
- Windows Excel COM \`Workbooks.Add()\` to inherit locally configured New Workbook defaults. Never modify widths/AutoFit or decorate; Short Date actual dates, built-in Comma Style money, General all else.
- SQL benchmarks are for comparison ONLY, never substitutes for missing source values.
- Comparison identity columns + Excel + ASAP SQL + Difference = Excel - ASAP SQL. Distinguish missing matching rows from genuine zero.
- Never claim numeric correctness until Excel formulas are recalculated and independently reconciled for relevant branches.

## Public repository security
- Do **not** stage/push scripts, client data, actual field/rate mappings, private KB documents, screenshots, SQL results, period input records, real workbooks or credentials.
- A .gitignore is an additional guard, not a guarantee: review \`git status\` and staged diff and check employer policy before using external AI tools.
- Keep all public test examples synthetic.
