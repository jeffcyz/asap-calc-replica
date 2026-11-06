
# ASAP-SQL-Remap — Audit Performance Root Cause

Continue the existing project.

Current evidence:

- 23 unit tests pass.
- Python compilation passes.
- Excel COM sensitivity-label access works.
- Real AALC2201 audit has not completed.
- Previous audit sessions were terminated before
  producing end-to-end timing evidence.

PRIMARY OBJECTIVE:

Identify the exact performance bottleneck of the
real workbook dependency audit.

Do NOT perform another broad refactoring.

Do NOT simply increase the timeout without evidence.

## 1. Persistent Real-Time Diagnostics

Instrument every dependency scan stage.

Immediately write progress to a persistent local log.

Log:

- Timestamp
- Worksheet name
- Operation name
- UsedRange dimensions
- Number of formula cells examined
- Number of ASAP_SQL matches found
- Current elapsed seconds
- Total elapsed seconds

Flush logs immediately after every checkpoint.

Do not rely exclusively on a final Markdown report.

## 2. Detailed Scan Profiling

Measure independently:

A. Worksheet enumeration
B. Formula reference discovery
C. Defined names
D. Conditional formatting
E. Data validation
F. Charts
G. Power Query definitions
H. Workbook connections

For every worksheet, record start and completion.

Identify the exact operation consuming the most time.

Before potentially blocking COM calls, log the
operation and target object.

## 3. Investigate Excessive COM Calls

Inspect for inefficient patterns:

- Individual cell.Formula calls
- Repeated UsedRange enumeration
- Enumerating entire columns
- Per-cell conditional formatting inspection
- Repeated Formula/Formula2 reads
- Repeated workbook object enumeration
- Expensive whole-workbook snapshots

Prefer Range.Find/FindNext for locating ASAP_SQL
references in formula cells.

Preserve dependency-detection correctness.

Do not skip hidden worksheets or indirect references.

## 4. Preserve Existing Functionality

Do NOT alter:

- Formula migration rules
- Verified header mapping
- Financial calculations
- Sensitivity label enforcement
- Version naming
- Excel workbook preservation
- Existing fail-closed behavior

## 5. Audit Execution

Run one controlled audit:

python tools/remap_proof.py --deal AALC2201 --audit-only --com-timeout 1200

Use the correct existing CLI path and arguments.

Do not prematurely terminate a healthy running audit.

If execution is interrupted or times out,
preserve all progress logs and checkpoints.

Never terminate pre-existing user Excel processes.

## 6. Required Results

Report:

1. Total audit execution time, if completed
2. The slowest worksheet
3. The slowest dependency category
4. The exact expensive COM operation
5. Number of formulas inspected
6. Number of ASAP_SQL dependencies found
7. Whether audit completed
8. Evidence-based optimization recommendation

If audit fails, provide the last completed checkpoint
and the exact operation in progress.

Do not claim performance improvement without
real workbook timing evidence.

Do not start another broad optimization before
identifying the bottleneck.

Execute the instrumented audit now.
