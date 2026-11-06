
# ASAP SQL Remap — Sensitivity Label COM Blocking Fix

This is a continuation of the existing ASAP-SQL-Remap project.

The reusable migration engine is implemented.
All 14 unit tests pass.

However, real XLSB migration is blocked because:

    Workbook.SensitivityLabel.GetLabel()

does not return when called through Excel COM.

The workbook opens successfully before this call.

Your task is to isolate and resolve this blocking issue
without modifying the existing financial remapping logic.

## 1. Strict Scope

Do not refactor:
- Formula mapping
- Column mapping
- Version management
- Workbook comparison
- Financial regression checks
- Repository architecture

Focus on sensitivity-label handling and its Excel
automation lifecycle.

Do not disable or bypass corporate protection policies.

## 2. Diagnose the Actual COM Blocking Point

Inspect the current sensitivity-label implementation.

Identify:
- Whether Excel is started using DispatchEx
- Whether Application.Visible is False
- Whether DisplayAlerts is disabled
- Whether Excel is running in an interactive user session
- The signed-in Office identity and policy availability
- The exact API call where execution stops
- Whether the workbook already carries a label
- Whether a modal Office dialog is blocking automation

Add timestamped logging immediately before and after
each COM operation.

Distinguish a hanging API from an exception.

Do not assume that GetLabel hanging proves an
authentication failure.

## 3. Add an Independent Label Diagnostic

Create a diagnostic entry point:

    python tools/remap_proof.py --diagnose-label

This must:

1. Start Excel in an isolated process.
2. Run under the current authorized Windows user.
3. Test sensitivity-label initialization.
4. Test GetLabel on an appropriate disposable workbook.
5. Test GetLabel on a private copy of a supplied proof.
6. Record the exact successful and failing stages.
7. Exit with a bounded timeout and actionable diagnosis.

Use visible Excel when diagnosing interactive
authentication or policy prompts.

Never kill or modify an existing user-owned Excel session.

Never report a successful label operation without
actual API confirmation.

## 4. Prevent Indefinite Hangs

Run potentially blocking Excel COM operations in an
isolated worker process.

Do not use a Python thread timeout as if it could
cancel an already-blocked COM call.

Use a supervisor process with a configurable timeout.

Create and use COM objects only inside the owning
worker process/thread.

If a timeout occurs:

- Capture the blocked operation.
- Capture the worker-owned Excel process identity.
- Clean up only processes created by this run.
- Preserve the original input workbook.
- Do not publish output.
- Produce a failure report.
- Return control to the command line.

Do not terminate unrelated Excel processes.

## 5. Preserve Internal Label Requirements

The final workbook must carry the authorized
Internal sensitivity label.

Expected LabelId:

    dd181445-6ec4-4473-9810-00785f082df0

Verify this ID is valid under the current
organization's published labeling policy.

Use Excel-native sensitivity-label APIs.

If the source already has another label,
do not overwrite or downgrade it.

If the label is absent, apply Internal only through
the organization's authorized mechanism.

Do not invent SiteId or authentication metadata.

Do not bypass policy enforcement.

After saving:
- Reopen the XLSB.
- Read the persisted label.
- Confirm the exact Internal LabelId.

## 6. Separate Label Diagnosis from Formula Diagnosis

Add a safe audit-only command:

    python tools/remap_proof.py --audit-only --deal AALC2201

It should inspect:

- ASAP_SQL dependencies
- tblASAP_SQL_DB headers
- Column mappings
- Unsupported formula patterns
- Source data equivalence

It must not:

- Modify source workbooks
- Publish a migrated workbook
- Apply or change sensitivity labels
- Claim migration success

This allows debugging the SQL remapping independently
of Office sensitivity-label initialization.

## 7. Add Focused Regression Tests

Test:

- Label COM timeout
- Visible versus hidden Excel diagnostic handling
- Failure report creation
- Worker cleanup
- Protection of existing Excel sessions
- Existing label conflict
- Label persistence verification
- No output publication after label failure
- Unchanged original XLSB
- Audit-only mode

Preserve all existing tests.

## 8. Deliverables

Provide:

1. Root-cause analysis with actual diagnostic evidence.
2. Minimal code changes.
3. Updated test results.
4. Label diagnostic command and logs.
5. Safe timeout behavior.
6. SQL audit-only capability.
7. Instructions for running the next real XLSB test.

Do not claim a successful migration until an actual
XLSB has been fully migrated, saved, reopened, and
validated with the Internal label.

Execute the troubleshooting and implement the fix.
