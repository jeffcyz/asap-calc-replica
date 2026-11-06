
# ASAP-SQL-Remap — Fix Comparison FindNext Bottleneck

Continue the existing project.

Latest real-workbook telemetry establishes that:

1. Worksheet enumeration performance improved.
2. Repeated UsedRange retrieval was reduced.
3. Diagnostic-only formula counting was removed.
4. 28 unit tests pass.
5. The current audit still has not completed.
6. The current bottleneck is the Comparison
   worksheet's FindNext("ASAP_SQL") loop.

Do not start another broad optimization.

PRIMARY OBJECTIVE:

Fix the Comparison FindNext bottleneck and complete
one real AALC2201 dependency audit.

## 1. Identify the Exact FindNext Behavior

Inspect the existing formula discovery implementation.

For the Comparison worksheet, log:

- UsedRange address and dimensions
- Initial Find result
- First match address
- Current match address
- FindNext iteration count
- Time spent in each FindNext call
- Total FindNext elapsed time
- Number of unique addresses found
- Whether the search wrapped to the first address

Log before and after the actual COM calls.

Determine whether the problem is:

A. Excessive legitimate matches
B. Incorrect FindNext wraparound handling
C. Repeated visits to the same cells
D. One or more unusually slow COM calls
E. Oversized or corrupted UsedRange
F. Another verified cause

Do not assume an infinite loop without evidence.

## 2. Implement Bounded Bulk Formula Discovery

Evaluate replacing the expensive FindNext loop with
chunked Range.Formula2 reads.

Requirements:

- Read formulas in bounded rectangular chunks.
- Avoid individual COM calls for every cell.
- Keep chunk size configurable.
- Start with approximately 5,000 cells per chunk.
- Extract formula strings in Python memory.
- Search for ASAP_SQL references.
- Retain exact worksheet and cell addresses.
- Preserve original formula text for subsequent rewriting.
- Include hidden rows, columns, and worksheets.
- Do not skip areas based only on visible values.

Use actual UsedRange boundaries.

Handle single-cell, single-row, and 2D COM return
shapes correctly.

Check that a formula begins with "=" before
treating it as a formula dependency.

Detect both direct references and unsupported
indirect references.

Do NOT write bulk Formula2 arrays back to Excel.

Only individually rewrite formulas that have been
confirmed safe to migrate.

## 3. Handle Large or Sparse UsedRanges

Inspect whether Comparison has an inflated UsedRange
because of formatting or historical data.

Do not assume all UsedRange cells contain formulas.

Avoid massive single-call Formula2 extraction.

If the range is unusually large:

- Use bounded chunks.
- Minimize COM calls.
- Preserve exhaustive coverage.
- Report actual cell count and chunk count.
- Fail safely if exhaustive scanning is impossible.

Do not silently skip large worksheets.

## 4. Correct FindNext Behavior Where Retained

If FindNext remains in any supported code path:

- Save the first match address.
- Explicitly handle wraparound.
- Detect repeated addresses.
- Track iteration count.
- Record per-call duration.
- Bound iteration count using a justified limit.
- Fail closed rather than returning partial results.

Use explicit Find parameters.

Do not remove the existing dependency checks.

## 5. Prevent Incomplete Audit Results

Never interpret:

    ASAP_SQL matches = 0

as a final result unless all relevant worksheets
and workbook objects were fully inspected.

Reports must distinguish:

    COMPLETED
    INCOMPLETE
    TIMEOUT
    FAILED

Only COMPLETED may establish the final number
of dependencies.

## 6. Test Both Discovery Methods

Create controlled Excel workbooks containing:

- SUMIFS references
- COUNTIFS references
- INDEX/MATCH references
- Quoted ASAP_SQL references
- Multiple references in the same formula
- Formulas containing INDIRECT
- No-match worksheets
- Hidden worksheets
- Large formula blocks
- Sparse UsedRanges
- Repeated FindNext matches

Compare the old and optimized discovery results.

They must identify identical dependencies.

Unsupported cases must remain explicit blockers.

## 7. Run One Controlled Real Audit

First verify no previous audit worker is active.

Then execute:

python tools/remap_proof.py --deal AALC2201 --audit-only --com-timeout 1200

Record checkpoints to persistent local files.

Do not terminate a healthy run prematurely.

If the optimized scan fails, identify the exact
remaining operation.

Do not claim success based only on unit tests.

## 8. Required Final Evidence

Provide:

1. Comparison worksheet UsedRange dimensions
2. Original FindNext iteration count and timings
3. Confirmed cause of the bottleneck
4. New formula discovery method
5. Number of chunks read
6. Comparison worksheet scan duration
7. Total AALC2201 audit duration
8. Number of actual ASAP_SQL formula dependencies
9. Unsupported dependencies
10. Final audit status

After the audit is fully completed, perform one
full AALC2201 migration if safe.

Verify:
- Migrated formulas
- SQL table preservation
- Financial result equivalence
- ASAP_SQL removal
- Version increment
- Persisted Internal sensitivity label
- Final XLSB reopen

Never publish an unverified output.

Execute the targeted fix and validation.
