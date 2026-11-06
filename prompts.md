
# ASAP SQL Remap — Dependency Scan Performance Optimization

Continue the existing ASAP-SQL-Remap project.

The sensitivity-label issue has been resolved.

Verified evidence:
- Excel COM initialization succeeds.
- GetLabel() succeeds in hidden Excel.
- Original proof already has the correct Internal label.
- All 23 unit tests pass.

The new blocker is:

    Migration.ASAP_SQL.dependencies

This stage exceeded the 300-second supervisor timeout.

Do not assume that merely increasing the timeout
will solve the problem.

## PRIMARY OBJECTIVE

Optimize dependency discovery and workbook inspection
for large XLSB proof models.

Preserve all existing financial and migration logic.

Do not change working sensitivity-label handling.

## 1. Profile Dependency Discovery

Instrument the dependency scanner with:

- Worksheet name
- UsedRange dimensions
- Number of detected formula cells
- Number of matching ASAP_SQL formulas
- COM call timing
- Elapsed time by worksheet
- Elapsed time by dependency category

Separately measure:

A. Cell formulas
B. Defined names
C. Conditional formatting
D. Data validation
E. Chart references
F. Power Query
G. Workbook connections
H. Other workbook objects

Record progress during execution.

Determine the actual bottleneck before optimizing.

## 2. Eliminate Per-Cell COM Formula Scanning

Do not read every formula cell individually merely
to discover references to ASAP_SQL.

Prefer Excel-native Range.Find / FindNext.

Search for:

    ASAP_SQL

with explicit settings:

    LookIn = xlFormulas
    LookAt = xlPart
    SearchOrder = xlByRows
    SearchDirection = xlNext
    MatchCase = False
    SearchFormat = False

Handle FindNext wraparound correctly.

Inspect and validate matching formulas before migration.

Use targeted matching so unrelated SQL strings
or ordinary text are not treated as sheet references.

Where necessary, use bulk Formula/Formula2 extraction
in bounded chunks instead of one COM call per cell.

Preserve dynamic-array and implicit-intersection
semantics. Do not silently convert Formula2 to Formula
or vice versa.

## 3. Optimize Workbook Object Auditing

Avoid scanning every cell individually for:

- Conditional formatting
- Data validation
- Other workbook metadata

Inspect rules and objects at the appropriate
worksheet/workbook collection level.

Preserve coverage of indirect references.

If an object cannot be inspected reliably,
report an unresolved dependency instead of
assuming it is safe.

Do not disable existing safety checks for speed.

## 4. Optimize Snapshot and Comparison

Inspect the existing baseline snapshot implementation.

Avoid unnecessary per-cell COM calls.

Use efficient bounded bulk reads where safe.

Preserve:
- Exact formula comparison
- Financial-result comparison
- Unchanged-value verification
- Workbook structure verification
- Formatting and connection preservation checks

Do not weaken financial regression requirements.

Any optimized comparison must produce the same
PASS/FAIL conclusions as the existing implementation.

## 5. Keep Separate Execution Stages

Expose timing for:

1. Open workbook
2. Inspect sensitivity label
3. Locate SQL table
4. Discover dependencies
5. Verify column mappings
6. Verify SQL source equivalence
7. Capture baseline
8. Rewrite formulas
9. Recalculate and compare
10. Delete ASAP_SQL
11. Save XLSB
12. Reopen and verify

This will help distinguish slow Excel calculation
from slow dependency discovery.

## 6. Preserve Existing Safety Requirements

Do not:
- Modify original XLSB
- Skip sensitivity-label validation
- Disable financial regression
- Hardcode financial results
- Skip unresolved dependency checks
- Remove necessary workbook objects
- Publish incomplete output
- Kill user-owned Excel processes

Keep the existing isolated-worker timeout protection.

## 7. Validation

First run the optimized audit-only mode
against AALC2201.

Record actual timings and discovered dependencies.

Compare the optimized scanner against the existing
scanner on small controlled test workbooks to verify
that no dependencies are missed.

After the audit passes, attempt a full migration.

Do not label the migration PASS until:
- All relevant dependencies have been remapped
- ASAP_SQL is removed
- Query table remains intact
- Financial results match
- Internal label is verified after reopening
- Versioned XLSB is successfully published

## 8. Deliverables

Provide:

1. Actual cause of the timeout
2. Before/after dependency scan timings
3. Number of detected formulas
4. Number of ASAP_SQL-dependent formulas
5. Optimization changes
6. Regression test results
7. Actual AALC2201 migration results
8. Any remaining blockers

Do not stop after implementation.
Run the real audit and full migration if possible.

Execute now.
