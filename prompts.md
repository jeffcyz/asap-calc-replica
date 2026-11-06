
# ASAP SQL Remap — Resolve Final Audit Blocker and Complete First Migration

Continue the existing ASAP-SQL-Remap project.

The latest real AALC2201 dependency audit has completed.

Verified results:

- Dependency discovery completed.
- Total reported dependencies: 48,301.
- Dependency scan duration: 1229.920865 seconds.
- Comparison UsedRange: A1:UF279.
- Comparison formula discovery: 226.669034 seconds.
- Chunked Formula2 scanning is operational.
- The previous FindNext bottleneck was eliminated.
- Current audit result: AUDIT_FAIL.
- Only one blocking condition was reported:

    Comparison!$AC$4

    Unsupported finite or direct source reference.

The original workbook remains unchanged.

PRIMARY OBJECTIVE:

Resolve the remaining unsupported reference safely,
achieve AUDIT_PASS, and execute the first complete
AALC2201 XLSB migration.

Do not perform another broad refactoring.

==================================================
1. INSPECT THE EXACT BLOCKING FORMULA
==================================================

Read the actual formula at:

    Comparison!$AC$4

Report:

- Complete original formula
- Referenced ASAP_SQL cells/ranges
- Absolute and relative reference semantics
- Original calculated result
- Whether it references a header, data row,
  control value, or summary row
- Whether the referenced source data exists
  in tblASAP_SQL_DB
- Whether equivalent behavior can be guaranteed

Do not assume the reference is a normal data cell.

==================================================
2. IMPLEMENT SAFE FINITE/DIRECT REFERENCES
==================================================

Extend formula migration only as required.

Potential patterns include:

    ASAP_SQL!$H$15
    ASAP_SQL!H15
    ASAP_SQL!$H$15:$H$25

These are examples, not confirmed workbook formulas.

Support only patterns for which a verified
semantic transformation is available.

For direct data-row references, consider:

    INDEX(tblASAP_SQL_DB[VerifiedHeader], row_index)

ONLY when all required conditions are proven:

- Source row mapping is unambiguous.
- Data rows correspond exactly.
- Row ordering semantics are preserved.
- Refresh behavior remains equivalent.
- Relative/absolute reference behavior remains valid.
- No hidden filters or additional SQL rows
  change the intended result.

If these conditions cannot be proven, use another
verified semantic mapping or fail closed.

Do not create a positional INDEX mapping merely
to eliminate an audit blocker.

Do not introduce hardcoded financial values.

==================================================
3. HANDLE SPECIAL REFERENCE TYPES
==================================================

If the formula references:

A. A table header:
   Use the verified structured header reference
   only when semantically equivalent.

B. A source data cell:
   Use verified row/field lookup semantics.

C. A finite aggregate range:
   Preserve exact inclusion, aggregation,
   and row-boundary behavior.

D. A source summary/control cell:
   Determine whether an equivalent value exists
   in the SQL table without inventing a mapping.

E. A copied relative reference:
   Preserve intended fill/copy behavior,
   not merely the current displayed result.

If exact equivalence is impossible,
stop and identify the specific limitation.

==================================================
4. VALIDATE THE 48,301 DEPENDENCIES
==================================================

Distinguish:

- Number of unique formula cells
- Number of ASAP_SQL reference occurrences
- Number of source columns referenced
- Number of supported formula patterns
- Number of unsupported formula patterns

Do not assume 48,301 represents unique formulas.

Ensure that all original dependencies have been
identified and accounted for.

No dependency may be silently skipped.

Preserve the existing completed audit evidence.

==================================================
5. MINIMAL IMPLEMENTATION CHANGES
==================================================

Modify only the necessary formula parsing,
rewriting, mapping, and validation components.

Do not change:

- Existing SUMIFS migration behavior
- Existing COUNTIFS migration behavior
- Existing lookup behavior
- Excel COM lifecycle
- Sensitivity label handling
- Versioning
- Workbook preservation rules
- Worker isolation
- Fail-closed safeguards

Add focused tests for the newly supported reference
pattern and its edge cases.

==================================================
6. RERUN COMPLETE AUDIT
==================================================

Execute the real AALC2201 audit.

Use a supervisor timeout supported by the actual
measured duration and retain persistent checkpoints.

Confirm:

    AUDIT_PASS

Required evidence:

- Total dependency count
- Unsupported dependency count = 0
- All required SQL headers verified
- Data equivalence verified
- All workbook-object checks passed

Do not attempt final migration if audit fails.

==================================================
7. EXECUTE FIRST FULL MIGRATION
==================================================

After AUDIT_PASS, perform one real AALC2201
migration using the existing versioned input.

Requirements:

1. Preserve original XLSB.
2. Rewrite all supported ASAP_SQL references.
3. Preserve unrelated formulas.
4. Preserve financial results.
5. Delete ASAP_SQL only after dependency validation.
6. Preserve Query and tblASAP_SQL_DB.
7. Preserve workbook formatting and structure.
8. Preserve or verify the Internal sensitivity label.
9. Save as the next appropriate version.
10. Reopen and validate the saved XLSB.

Do not publish a partially verified workbook.

==================================================
8. LARGE-WORKBOOK PERFORMANCE
==================================================

Because approximately 48,301 dependencies were found,
measure formula replacement and regression performance.

Avoid unnecessary one-cell-at-a-time COM reads.

Use bounded bulk reads and safe targeted writes
where they preserve formula semantics.

Do not bulk overwrite unrelated formulas.

Record:

- Formula rewrite duration
- Recalculation duration
- Financial comparison duration
- Workbook structure comparison duration
- Save duration
- Reopen duration
- Sensitivity label verification duration

Do not weaken validation solely to improve speed.

==================================================
9. FINAL REPORT
==================================================

Provide:

A. Exact original Comparison!AC4 formula.

B. Its verified replacement formula.

C. Explanation of why the mapping is equivalent.

D. Final audit status.

E. Number of unique affected formula cells.

F. Number of reference occurrences migrated.

G. Number of unsupported references remaining.

H. Full migration execution time.

I. Financial comparison result.

J. Workbook preservation result.

K. Internal label persistence result.

L. Original and final versioned filenames.

M. Final output path.

N. Overall PASS / FAIL.

If final migration fails, report the exact
next blocker and preserve the original workbook.

Do not claim success merely because unit tests pass.

EXECUTION:

Perform the targeted fix, run the audit, and
attempt the first complete migration.

Do not stop after describing a plan.
