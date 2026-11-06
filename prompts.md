
# ASAP-SQL-Remap — Complete Final Migration Validation

Continue the existing project.

CURRENT VERIFIED STATUS:

AUDIT_PASS:
- Dependency scan completed.
- Existing audit evidence remains valid.
- Approximately 48,300 ASAP_SQL dependencies identified.
- All reported unsupported audit blockers resolved.

MIGRATION_NOT_VERIFIED:
- Save/reopen/persisted-label final validation
  has not completed to PASS.
- No successful migrated workbook may be claimed.

PRIMARY OBJECTIVE:

Identify the exact incomplete migration stage,
fix it with minimal changes, and complete one
fully validated AALC2201 migration.

Do not refactor the dependency scanner or
formula-remapping engine without evidence
that either is responsible for this failure.

## 1. Inspect Existing Migration Evidence

First read the latest:
- Migration console log
- Failure report
- Worker telemetry
- Processing history
- Staging-file status

Determine the last successfully completed operation
and the first failed, blocked, or unverified operation.

Distinguish:
- Exception
- COM hang
- Supervisor timeout
- Validation mismatch
- Sensitivity label mismatch
- Incomplete or interrupted execution

Do not guess the failure cause.

## 2. Instrument the Finalization Pipeline

Add persistent, flushed BEFORE/AFTER checkpoints
to every step:

1. Formula rewrite completed
2. Financial recalculation completed
3. Pre-save regression comparison passed
4. ASAP_SQL removed
5. Internal label inspected/confirmed
6. XLSB save started
7. XLSB save returned successfully
8. Workbook close completed
9. Excel worker cleanup completed
10. Saved XLSB reopened
11. Internal LabelId read successfully
12. Persisted label verified
13. Final financial regression passed
14. Workbook structure validation passed
15. Final output published

For every checkpoint capture:
- Timestamp
- Stage
- Excel process ID
- Workbook path
- Duration
- Exception or timeout information

All checkpoints must survive process termination.

## 3. Investigate Save/Reopen

Check the actual implementation for:

- Save versus SaveAs behavior
- XLSB FileFormat=50
- Temporary staging-file paths
- File locks
- Read-only status
- Existing destination-file collisions
- Excel alerts and modal prompts
- Workbook calculation state
- Power Query/background refresh activity
- COM object lifecycle
- Save completion before reopen

Do not replace the Excel-native XLSB workflow.

Do not change workbook connections or calculation
semantics merely to make saving faster.

## 4. Verify Internal Label Persistence

Expected label: Internal

Expected LabelId:
dd181445-6ec4-4473-9810-00785f082df0

If the source workbook already carries the
correct authorized Internal label:

- Preserve it.
- Avoid unnecessary reassignment.

If the source is unlabeled:

- Apply Internal through the approved Office API
  using authorized configuration.

If the source has a different label:

- Fail closed.
- Do not overwrite or downgrade it.

After saving, close the workbook and reopen
the staged XLSB.

Use a fresh isolated Excel instance if necessary
to verify persisted metadata independently.

Call GetLabel() on the reopened workbook.

Require exact LabelId equality.

Do not treat a successful pre-save GetLabel()
as proof that the saved file retained the label.

Do not bypass the label check or manipulate
Office protection metadata directly.

## 5. Reuse Existing Audit Safely

Do not rerun the 20-minute audit unnecessarily.

Reuse the existing AUDIT_PASS only when:

- Source workbook SHA-256 is unchanged.
- Source table schema is unchanged.
- Relevant mapping and dependency rules are unchanged.
- The existing report proves complete coverage.

If these conditions do not hold, rerun the
necessary validation.

Do not reuse partial migration results from
a failed staging file without verification.

## 6. Final Publication Gate

Publish the version-incremented XLSB only after:

- All intended formulas migrated
- No unresolved ASAP_SQL dependencies
- ASAP_SQL worksheet removed
- Query!tblASAP_SQL_DB preserved
- Financial results equivalent
- Unrelated formulas preserved
- XLSB saved and reopened
- Internal label persisted and verified
- Workbook structure validation passed

Until every mandatory check passes:

    MIGRATION_NOT_VERIFIED or MIGRATION_FAIL

Do not publish an unverified staging workbook.

Never overwrite the original input.

## 7. Execution

Use the existing AALC2201 input.

Run a single controlled migration with
the existing isolated-worker supervisor.

Keep granular telemetry throughout execution.

Do not prematurely terminate a healthy worker.

If execution fails, retain its final checkpoint,
failure report, and all useful diagnostics.

Do not claim success merely because unit tests pass.

## 8. Required Final Report

Report exactly:

- Audit status
- Source workbook SHA-256
- Formula rewrite status
- Migrated formula count
- Financial comparison status
- XLSB save status
- Close status
- Reopen status
- Original sensitivity label
- Persisted sensitivity label
- Final workbook structure status
- Final output filename and path
- Final publication status
- Overall MIGRATION_PASS / MIGRATION_FAIL

If the run fails, identify the exact failed stage
and the next minimal corrective action.

Execute the investigation and complete the
migration if the evidence permits.
