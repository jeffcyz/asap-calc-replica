# ASAP Knowledge Base integration

## Why this exists

ASAP / IronPython calculation scripts are **not self-contained programs**. They may depend on environment-supplied objects, collection membership, index services, date/period resolution, hidden initialization, mutation semantics, SQL schema relationships and prior-period state. The objective is to **reproduce observable calculation semantics in native Excel**, not to run the original script outside ASAP or emulate the entire runtime.

## Security boundary

- This GitHub repository is currently **public**.
- The documentation here contains *only the proposed knowledge organization*, not an actual ASAP manual.
- Keep authorized knowledge at \`knowledge/asap/\` locally (Git-ignored), or use an approved local read-only link to existing ASAP-Agent knowledge.
- Do not push proprietary source scripts, internal PDF manuals, screenshots, loan data, index snapshots, or deal field definitions without explicit company authorization. Ignoring files in Git does not by itself authorize sending them to Copilot or an external AI provider.

## Prioritized knowledge intake

1. P0: Deal, Asset, AssetPeriodic, AssetPeriodicTransaction object/property contracts and raw SQL mapping.
2. P0: collection behavior such as \`assets.get_active(period)\`: membership, state, timing and side effects.
3. P0: periodic execution lifecycle: setup, initialization, order, persisted vs computed field values.
4. P1: critical dates, periods, payment_inputs and \`deal.get_index_value\` contract.
5. P1: historical carryforward, default/null types, IronPython numerical behavior, rounding and actual/360.
6. P2: less frequently used APIs discovered while analyzing additional deals.

## Evidence requirements

For each runtime symbol, record an exact source (document/path/version/section or controlled experiment), explicit required inputs, read/write side effects and calculation-time behavior. **Knowledge explains operations; it does not supply specific monthly observations.** Obtain actual observed inputs using authorized snapshots separately.

## Work order

1. Inventory unknown runtime symbols in the first local deal script.
2. Look up only the needed authorized KB sections; attach evidence to contract records.
3. Mark semantic gaps, input/data gaps, and translation gaps independently.
4. Classify outputs A/B/C/D using \`docs/calculation_coverage.md\`, with unknowns failing closed.
5. Proceed to Excel formulas only for adequately supported fields; benchmark with multiple branch-covering periods.
6. Generalize shared **contracts** across deals only after confirming their semantics are consistent.
