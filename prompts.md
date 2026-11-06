
# ASAP SQL Remap — Targeted COM Audit Optimization

Continue the existing reusable XLSB migration project.

We now have real AALC2201 profiling evidence.

CURRENT MEASUREMENTS:

- Total elapsed: approximately 568.59 seconds
- Audit was still running
- dependency.worksheet_enumeration: ~437.10s cumulative
- dependency.other_workbook_objects: ~84.60s
- dependency.chart_references: ~83.84s
- dependency.conditional_formatting: ~82.19s

Individual worksheet enumeration:
- Config: ~53.20s
- AggregateAudit: ~47.32s
- Diagnostics: ~47.30s

Individual expensive calls:
- Diagnostics worksheet enumeration: ~29.76s
- Comparison worksheet enumeration: ~29.46s

No ASAP_SQL matches have yet been observed in the
incomplete audit.

The current unit tests pass.

OBJECTIVE:

Eliminate excessive Excel COM calls and repeated
worksheet/object traversal without weakening
dependency detection or financial validation.

## 1. Optimize Worksheet Enumeration

Inspect the entire worksheet enumeration path.

Determine:
- How often each worksheet is enumerated
- How often UsedRange is requested
- Whether formula counts cause full range scans
- Whether each dependency category independently
  traverses the same expensive Excel objects
- Whether COM collections are repeatedly materialized

Acquire and reuse worksheet references and UsedRange
where safe within the same scan phase.

Avoid redundant COM requests.

Do not cache data across workbook mutations or
refresh events if doing so could produce stale results.

## 2. Remove Expensive Diagnostic-Only Calls

Identify calls such as:

    UsedRange.SpecialCells(xlCellTypeFormulas).Count

If they exist only to calculate progress statistics,
remove them from the hot path.

Do not enumerate every formula merely to report
the number of formulas examined.

Do not compromise mandatory dependency coverage.

Counts that are expensive to obtain should be
reported as unavailable or estimated, clearly labeled.

## 3. Optimize Formula Discovery

Use Excel-native Range.Find / FindNext for:

    ASAP_SQL

Search formulas, not just displayed values.

Use explicit Find parameters.

Handle first-address wraparound correctly.

Collect only matching cell addresses and formulas
needed for the source migration.

Avoid per-cell COM access across unrelated formulas.

Preserve detection for quoted and unquoted references.

Preserve checks for unsupported and dynamic formulas.

Do not interpret zero observed matches before scan
completion as evidence of zero dependencies.

## 4. Optimize Workbook Object Auditing

Inspect these categories independently:

- Defined names
- Conditional formatting
- Data validation
- Chart references
- Power Query
- Workbook connections
- Other relevant objects

Avoid repeated collection traversal where possible.

Use collection-level inspection instead of per-cell
inspection where correctness permits.

Do not assume that inspecting the first rule in a
range represents all rules.

Do not skip hidden worksheets or unsupported objects.

If exhaustive inspection is not possible, fail closed
and report the unresolved object.

## 5. Measure COM Calls and Timing

Before and after each category, record:

- Number of COM operations where measurable
- Per-worksheet elapsed time
- Category elapsed time
- Cumulative elapsed time
- Formula matches
- Unresolved dependencies

Maintain real-time flushed checkpoints.

Clearly distinguish exclusive phase durations
from cumulative elapsed durations.

Do not double-count nested timing categories.

## 6. Strict Preservation

Do not change:

- Formula remapping semantics
- SQL table header mapping
- Financial calculations
- Sensitivity label handling
- Workbook versioning
- Excel save format
- Worker isolation
- Original file protection
- Existing failure handling

Do not reduce verification coverage solely
to improve speed.

## 7. Focused Tests

Add regression tests demonstrating:

A. Cached versus uncached worksheet inspection
   discovers the same dependencies.

B. Find/FindNext discovers every matching formula
   in controlled workbooks.

C. No-match worksheets complete correctly.

D. Hidden worksheets are included.

E. Defined names and workbook-object dependencies
   remain detectable.

F. Unsupported formulas still fail closed.

G. Existing tests continue passing.

## 8. Real Workbook Validation

Do not launch a second audit concurrently with
an existing active audit worker.

After implementation, perform one controlled
AALC2201 audit.

Compare actual category timings with the
previous measurements.

If the audit completes, proceed to full migration
only after confirming dependency coverage.

If it remains slow, identify the specific
remaining COM operation using persistent telemetry.

## 9. Deliverables

Report:

1. Exact code paths causing repeated COM calls
2. COM operations removed or reduced
3. Per-sheet before/after timing comparison
4. Total dependency scan duration
5. ASAP_SQL formula references discovered
6. Unresolved workbook dependencies
7. Unit test results
8. Actual XLSB audit result
9. Full migration result, if completed

Do not claim optimization success solely because
unit tests pass.

Actual workbook timing evidence is required.

Implement this targeted optimization now.
