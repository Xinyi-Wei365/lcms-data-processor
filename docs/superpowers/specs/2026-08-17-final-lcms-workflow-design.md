# Final LC-MS/MS Workflow Design

## Goal

Make the web processor language-complete, analyte-general, and scientifically auditable before regenerating the Demo workbook.

## Confirmed Rules

1. The selected language controls all web labels, notices, validation messages, tables, downloadable workbook sheet names, headings, column labels, and calculation notes.
2. Every imported analyte is retained. The system proposes type, chain length, and role from its name, then provides an editable classification table where the user can correct type, chain length, and choose Target, SS, or IS.
3. Output ordering is Type, then chain length, then name. Unknown type/chain values remain explicitly labelled `Other` / `NA`; the processor must not invent chemical identity.
4. DDAC is displayed as DADMAC everywhere in the processed result and summary. Raw uploaded values are not changed.
5. Target analyte recovery is measured concentration divided by that analyte's concentration in the matching MS column, times 100 percent.
6. SS recovery is measured SS concentration divided by that SS's own concentration in the matching MS column, times 100 percent. SS rows appear in the designated SS recovery section.
7. IS addition concentrations are entered and exported as an independent analyte-by-MS table. IS values are record-only and never create recovery percentages or alter final concentration calculation. IS correction remains controlled only by the existing IS-corrected yes/no setting.
8. `blank = 0` remains as implemented: a numeric zero in all applicable blank cells uses user-entered calibration concentration and S/N, with MDL = 3 times calibration concentration divided by S/N. Blank cells are missing data, not zero.
9. The separate descriptive-statistics sheet is a final user result: Name, chain length, DF percent, Median (Q1-Q3), MDL, MQL. It retains formulas and displays three significant figures. It uses all valid final concentrations; DF remains true detection rate only.
10. CSV/XLSX/XLS/TSV import and CSV/XLSX export remain supported. CSV is a flat numeric report; XLSX preserves the multi-sheet formulas.
11. The Demo workbook is regenerated only after all rules above are implemented and is then used for end-to-end validation.

## Verification

- Unit tests cover language dictionaries and translated workbook labels, classification overrides, IS exclusion from recovery, per-MS SS/target calculation, and final descriptive statistics.
- An end-to-end test loads the regenerated Demo through Streamlit, processes it, previews numeric results, and produces XLSX and CSV downloads.
- The generated XLSX is inspected for the final summary sheet, IS record table, formulas, and absence of IS recovery rows.
