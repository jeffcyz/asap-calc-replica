# ASAP-SQL-Remap — AGENTS.md

Project purpose: make safe, reproducible changes to existing XLSB workbooks; never rebuild financial proofs.

1. Treat original workbook as immutable. Never overwrite it or an existing version.
2. Only intended workbook edits are replacement of ASAP_SQL references, removal of ASAP_SQL, and authorized application/verification of Internal label.
3. Resolve columns using actual tblASAP_SQL_DB headers, never column positions alone.
4. Preserve calculation criteria, matching, duplicate behavior, dates, nulls, lookup ordering, and Excel formatting.
5. Fail closed for unsupported formulas, inconsistent data, indirect dependencies, label mismatch, or changed financial results.
6. Support all deals with the same reusable engine; do not hardcode specific financial data or source paths.
7. Version output names: v05->v06, v0.0.1->v0.0.2, no version->_v01.
8. Apply Internal with Office SensitivityLabel and verify exact label GUID after reopening. Never downgrade or bypass sensitivity policy.
9. Run on a company-approved Windows Excel desktop and report actual validation, not assumptions.
10. This GitHub repository is PUBLIC. Never commit XLSB, confidential reports, raw deal records, tenant GUIDs, or local processing history.
11. Keep supported migration scope explicit and test any newly supported Excel formula structures.
