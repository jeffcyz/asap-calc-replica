# Excel requirements

- Excel for Windows COM `Application.Workbooks.Add()`; local default New Workbook theme and Normal style.
- Do not change default column widths or call AutoFit. No custom aesthetic formatting or Excel Table styles.
- Dates: actual date cells in Excel's local built-in **Short Date** style. Beware that COM NumberFormat/NumberFormatLocal semantics vary by locale.
- Money: apply built-in Excel **Comma Style** to actual monetary amount fields (not to identifiers/rates). Classify by business meaning, not value magnitude.
- Everything else: **General**, including rates and percentage-like raw numbers.
- Use INDEX + MATCH with exact matching for lookup fields. SUMIFS may be used for group aggregates.
- Model ordered intermediate stages when ASAP scripts mutate a field multiple times.
- Comparison sheet: identity fields, Excel, ASAP SQL, Difference = Excel - ASAP SQL; missing benchmark values remain visibly missing, not 0.
- Preserve original workbook date serial semantics, numeric rounding and actual/360.
