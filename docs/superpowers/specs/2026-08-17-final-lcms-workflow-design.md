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
8. For non-zero blanks, MDL is `MAX(t(0.99, n_spike-1) × SD(low-level spike replicates), mean(blank) + t(0.99, n_blank-1) × SD(blank))`; the one-sided 99% t value uses the actual replicate count. The low-level spike values must be same-concentration, full-method replicates and cannot be mixed from MS1/MS2/MS3 at different concentrations. When these low-spike values are unavailable, the workbook transparently applies only the blank branch and the interface warns that the full two-branch MDL method was not available. MQL is `MAX(10 × SD(low-level spike replicates), mean(blank) + 10 × SD(blank))`.
9. `blank = 0` is a visible main-interface rule: when all applicable blank cells for a compound are numeric zero, that compound is automatically suggested in the blank-zero setting. The user confirms it and enters the S/N from the corresponding calibration point. The interface also records that calibration point concentration because a concentration-unit MDL cannot be calculated from a dimensionless S/N alone. The calculation is `MDL = 3 × calibration-point concentration ÷ S/N` and `MQL = 10 × calibration-point concentration ÷ S/N`; blank cells are missing data, not zero.
10. A true detection is an original vial concentration greater than or equal to its vial MDL. DF is true detections divided by valid samples. Original blank sample cells stay blank and are excluded; numeric zero is a valid non-detect and is replaced by one-half of the final-unit MDL for concentration statistics. A true-detect final concentration is `(vial concentration − blank mean) × conversion factor`.
11. The separate descriptive-statistics sheet is a final user result: Name, chain length, DF percent, Median (Q1-Q3), MDL, MQL. It retains formulas and displays three significant figures. It uses all valid final concentrations; DF remains true detection rate only.
10. CSV/XLSX/XLS/TSV import and CSV/XLSX export remain supported. CSV is a flat numeric report; XLSX preserves the multi-sheet formulas.
11. The Demo workbook is regenerated only after all rules above are implemented and is then used for end-to-end validation.

## Verification

- Unit tests cover language dictionaries and translated workbook labels, classification overrides, IS exclusion from recovery, per-MS SS/target calculation, and final descriptive statistics.
- An end-to-end test loads the regenerated Demo through Streamlit, processes it, previews numeric results, and produces XLSX and CSV downloads.
- The generated XLSX is inspected for the final summary sheet, IS record table, formulas, and absence of IS recovery rows.
