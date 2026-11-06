
# ASAP SQL Remapping — Reusable XLSB Migration Engine

You are starting a NEW standalone Copilot Agent project.

This project will be used repeatedly to process proof models
from multiple ASAP deals and reporting periods.

Build a small, reusable Python utility that modifies existing
XLSB proof workbooks directly.

This is NOT a proof generation project.

Do not require the original ASAP calculation repository,
ASAP scripts, original calculation inputs, or historical
Copilot conversation context.

============================================================
1. PRIMARY OBJECTIVE
============================================================

For each versioned XLSB file in:

    inbox/<DealName>/

Perform the following:

1. Open the original XLSB using Microsoft Excel COM.
2. Locate the existing ASAP_SQL worksheet.
3. Locate the Query worksheet.
4. Verify the Excel ListObject tblASAP_SQL_DB.
5. Audit all dependencies on ASAP_SQL.
6. Map original source columns to verified table headers.
7. Redirect all applicable formulas to tblASAP_SQL_DB
   using Excel Structured References.
8. Preserve original financial calculation semantics.
9. Remove ASAP_SQL after successful dependency migration.
10. Apply and verify the Internal sensitivity label.
11. Save the workbook with an incremented version number.
12. Reopen the XLSB and perform regression validation.
13. Produce a migration and validation report.

Do not modify the original file.

Do not use the generic filename proof_after.xlsb.

The output filename MUST be derived from the input filename.

============================================================
2. PROJECT STRUCTURE
============================================================

Create and maintain:

ASAP-SQL-Remap/
    AGENTS.md
    README.md
    inbox/
        <DealName>/
            <versioned original XLSB>
    tools/
        remap_proof.py
    config/
        remap_rules.json
    output/
        <DealName>/
            <versioned migrated XLSB>
    reports/
        <DealName>/
            <migration reports>
    state/
        processing_history.jsonl

Keep this project lightweight.

Do not copy or reconstruct the original ASAP repository.

All migration logic should be reusable across deals.

============================================================
3. AUTOMATIC VERSION MANAGEMENT
============================================================

Preserve the original filename structure.

Increment its existing explicit version identifier.

Examples:

Input:
    proof_20261030_165600_001-v05.xlsb
Output:
    proof_20261030_165600_001-v06.xlsb

Input:
    proof-2404_v0.0.1_release.xlsb
Output:
    proof-2404_v0.0.2_release.xlsb

Input:
    proof_AALC2501_v1.2.9.xlsb
Output:
    proof_AALC2501_v1.2.10.xlsb

Input:
    proof_20261105_050502_601.xlsb
Output:
    proof_20261105_050502_601_v01.xlsb

Rules:

A. Preserve all filename components except the version.

B. For simple numeric versions:
       v05 -> v06
       v09 -> v10
       v099 -> v100

   Preserve the existing minimum numeric width.

C. For semantic versions:
       v1.2.9 -> v1.2.10

   Increment the final numeric component only.

D. Preserve all prefixes, suffixes, timestamps,
   deal identifiers, and extension capitalization.

E. Do not treat a timestamp or numeric run identifier
   as a version if it has no explicit version prefix.

F. If no explicit version exists, append _v01.

G. If multiple explicit version tokens create ambiguity,
   fail with an explanation instead of guessing.

H. Never overwrite an existing output file.

I. Never rename or modify the input workbook.

J. Check the source file fingerprint and processing history.

K. If the exact input was already successfully processed,
   report SKIPPED_ALREADY_PROCESSED.

L. If the workbook is already migrated and ASAP_SQL
   is absent, report ALREADY_MIGRATED and do not
   increment its version unnecessarily.

M. Create a new version only after a real and fully
   validated migration.

============================================================
4. DEAL-INDEPENDENT SOURCE MAPPING
============================================================

Different deals may have different ASAP_SQL layouts,
source columns, table schemas, or formula patterns.

Never hardcode a single deal's column positions
as universal rules.

For each workbook independently:

1. Inspect the original ASAP_SQL structure.
2. Inspect tblASAP_SQL_DB headers.
3. Identify all referenced source columns.
4. Establish the corresponding table header mappings.
5. Verify the mappings against available source evidence.
6. Verify row-grain and calculation compatibility.
7. Apply mappings only when equivalence is established.

Use:

    tblASAP_SQL_DB[VerifiedHeader]

Never use:

    Query!$A:$A

as a replacement for a structured table reference.

Do not assume that column H in ASAP_SQL corresponds
to column H in tblASAP_SQL_DB.

If a verified mapping is reusable, the project may
retain it as a deal-specific mapping record.

However, always revalidate the mapping against the
current workbook before applying it.

If a mapping is missing or ambiguous, fail that
workbook safely.

Do not invent source headers or financial logic.

============================================================
5. FORMULA MIGRATION
============================================================

Audit ALL ASAP_SQL formula dependencies.

Include:

- SUMIFS / SUMIF
- COUNTIFS / COUNTIF
- INDEX / MATCH
- XLOOKUP / VLOOKUP
- SUMPRODUCT
- IFERROR / IFNA
- Other supported formula patterns
- Defined names
- Relevant workbook objects
- Indirect and dynamic references where detectable

Preserve:

- All criteria
- Aggregation semantics
- Exact and approximate matching behavior
- Duplicate handling
- Data types
- Date handling
- Blank and zero handling
- Error handling
- Relative and absolute reference semantics

Do not replace formulas with hardcoded values.

Do not transform unrelated formulas.

Do not replace a whole-column range mechanically
when that would change the result or behavior.

Do not delete ASAP_SQL while unresolved dependencies
still exist.

============================================================
6. STRICT WORKBOOK PRESERVATION
============================================================

The ONLY intended workbook modifications are:

1. Replacement of ASAP_SQL dependencies.
2. Removal of worksheet ASAP_SQL.
3. Assignment or preservation of the required
   Internal sensitivity label.

Everything else must remain unchanged.

Preserve:

- Financial results
- All unrelated formulas
- Worksheets and their relative ordering
- Formatting
- Conditional formatting
- All Step #1/#2/#3 sections
- Cell values
- Row heights and column widths
- Excel tables and connections
- Power Query
- Named ranges
- Calculation settings
- All other workbook features

Do not regenerate any worksheet.

Do not rebuild the workbook from scratch.

Do not intentionally refresh SQL data during
the baseline comparison.

Use the same data snapshot for both comparisons.

============================================================
7. SENSITIVITY LABEL — INTERNAL
============================================================

Every successfully migrated workbook must retain
or receive the Microsoft Office sensitivity label:

    Internal

Expected LabelId:

    dd181445-6ec4-4473-9810-00785f082df0

Use native Microsoft Excel COM APIs:

    Workbook.SensitivityLabel
    CreateLabelInfo()
    SetLabel()
    GetLabel()

Requirements:

1. Inspect the original sensitivity label.

2. If it is already Internal with the expected ID,
   retain and verify it.

3. If unlabeled, apply Internal using the current
   organization's authorized labeling mechanism.

4. If another label exists, do not downgrade or
   overwrite it without authorization.

5. Use required assignment metadata and valid
   organizational policy settings.

6. Do not invent tenant-specific label metadata.

7. Save in XLSB format, FileFormat=50.

8. Close and reopen the saved workbook.

9. Verify the persisted label using GetLabel().

10. Require an exact match with the expected LabelId.

If the label cannot be applied or verified, fail
the workbook's final validation.

Do not bypass Microsoft 365 protection or policy.

============================================================
8. REUSABLE PYTHON ENGINE
============================================================

Create:

    tools/remap_proof.py

Use Python and Windows Excel COM Automation
via pywin32.

Implement reusable components for:

- File discovery
- Filename version parsing and incrementing
- XLSB opening and saving
- SQL table verification
- ASAP_SQL dependency discovery
- Header mapping
- Formula transformations
- Sensitivity label application
- Workbook comparison
- Reporting
- Processing history

Keep the code modular and testable.

Use one implementation for all deals wherever
the actual workbook schemas permit it.

Do not duplicate the entire script for each deal.

Support commands such as:

    python tools/remap_proof.py --deal AALC2501

    python tools/remap_proof.py --all

    python tools/remap_proof.py --file "path/to/file.xlsb"

These are required CLI interfaces.

--deal:
    Process eligible files for the specified deal.

--all:
    Process eligible files in all deal inbox folders.

--file:
    Process one explicitly selected workbook.

Provide clear success, skip, and failure statuses.

Each workbook must be handled independently.

Failure on one deal must not corrupt another.

Do not automatically overwrite existing versions.

============================================================
9. VALIDATION
============================================================

For every input workbook:

Record the baseline before making changes.

After migration verify:

A. ASAP_SQL has been removed.

B. All migratable ASAP_SQL references were replaced.

C. No unresolved dependency on the removed sheet remains.

D. Query and tblASAP_SQL_DB remain intact.

E. All migrated formula results match the baseline.

F. All unrelated formulas remain unchanged.

G. All unrelated financial values remain unchanged.

H. Existing workbook structure and formatting remain
   functionally equivalent.

I. The final workbook remains XLSB.

J. The workbook saves and reopens successfully.

K. Internal sensitivity label verification passes.

Do not use financial tolerances to conceal
unexplained changes.

Recognize and report any unavoidable limitations
in checking cached results or dynamic references.

Use a temporary working file during processing.

Publish the final versioned output only after
all mandatory validations pass.

On failure:
- Retain the original.
- Do not publish a successful output.
- Record the actual blocker.
- Continue with other independent files if --all
  was requested.

============================================================
10. PROCESSING HISTORY
============================================================

Maintain:

    state/processing_history.jsonl

Record for each attempted file:

- Deal name
- Input filename
- Input SHA-256
- Proposed output filename
- Actual output filename, if successful
- Processing timestamp
- Formula migration count
- Mapping verification status
- Financial comparison status
- Sensitivity label status
- Overall status
- Error information, if applicable

Use processing history for idempotency.

Do not unnecessarily create new versions for
previously processed identical input files.

============================================================
11. REPORTING
============================================================

Create an independent report for each workbook.

Example:

    reports/AALC2501/
        proof_20261030_165600_001-v06_remap_report.md

The report must include:

- Original filename
- Final filename
- Deal name
- Original sensitivity label
- Final sensitivity label
- Verified table name
- Column mapping details
- Number and types of formulas migrated
- Unresolved formula references
- Financial comparison summary
- Workbook structure comparison
- Label verification result
- Overall PASS / FAIL

Provide traceable details for failed validations.

Do not claim success without actual verification.

============================================================
12. INITIAL IMPLEMENTATION AND TESTING
============================================================

Use whichever versioned XLSB files currently exist
under inbox/ as the initial test inputs.

Build the reusable utility first.

Test the actual migration with the available workbook.

Add focused automated tests for:

- Version increments
- Filename preservation
- Collision handling
- Already-migrated files
- Column mapping
- Formula rewriting
- Missing headers
- Invalid structured references
- Workbook preservation checks
- Processing history and rerun behavior

Document actual execution requirements in README.md.

Do not introduce unnecessary project dependencies.

============================================================
13. PERMANENT PROJECT RULES
============================================================

Create AGENTS.md.

It must establish these permanent rules:

1. One reusable migration engine for all supported deals.
2. The input is an existing versioned XLSB proof.
3. Never regenerate the underlying financial model.
4. Always use verified table-header references.
5. Never modify business calculations unnecessarily.
6. Always preserve original filenames and increment versions.
7. Never overwrite source or existing output files.
8. Always require Internal label verification.
9. Always perform regression validation.
10. Never claim success if validation fails.
11. Never introduce unnecessary deal-specific code.
12. Keep processing repeatable and idempotent.

============================================================
14. EXECUTION INSTRUCTIONS
============================================================

Work autonomously.

Inspect existing files in inbox/.

Create the reusable Python implementation.

Create any necessary project configuration.

Create AGENTS.md and README.md.

Execute migration against the available sample files.

Validate the resulting workbook.

Produce versioned output and detailed reports.

Do not request the original ASAP repository.

Do not ask for historical Copilot session context.

Do not stop at drafting a plan.

The final project must support repeated processing
of independently supplied XLSB proof files from
different ASAP deals and reporting periods.

Begin implementation now.
