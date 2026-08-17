# SS/IS Example Tables Design

## Goal

Show detailed SS and IS per-MS concentration examples on the main interface before a file is uploaded.

## Confirmed content

The SS example table contains `SS替代物`, `MS1基质加标浓度（ppb）`, `MS2基质加标浓度（ppb）`, and `MS3基质加标浓度（ppb）`. It shows `d7-C12-BAC` values `4`, `8`, `12` and `d9-C10-ATMAC` values `2`, `2`, `4`. These values are the SS compound's own matrix-spike concentrations used as the denominator of recovery; they are not measured MS concentrations. Measured concentrations are read automatically from the uploaded source.

The IS example table contains only `IS内标化合物名称`, with `IS-A` and `IS-B`. Users enter IS names only. The app matches those names to uploaded compounds; IS has no matrix-spike concentration input and no recovery calculation.

Both tables are bilingual and always visible. MS1/MS2/MS3 are examples only. The actual compound-by-MS editor remains file-driven and creates fewer or more columns to match every real MS column in the uploaded file. That editor contains target compounds and SS compounds, but excludes IS compounds.

## User input punctuation guidance

The SS and IS custom-name input areas must show the exact accepted name syntax, not only the rendered example tables. The IS workflow ends with name selection. The SS name box identifies which compounds are surrogates; after upload, the user must also enter every SS compound's own matrix-spike concentration for every detected MS column. Multiple names may be separated by any one of these delimiters:

- English comma: `,`
- Chinese comma: `，`
- English semicolon: `;`
- Chinese semicolon: `；`
- Tab character

The interface must include copyable SS and IS name examples. After upload, the left sidebar dynamically creates one real concentration input for every SS-by-MS pair; the read-only main-page example is never an input control. The main-page compound-by-MS editor remains the source of target-analyte spike concentrations only. Every sidebar SS input is required and must start empty; the app must not assume a default such as 4 ppb. Processing remains disabled until all SS inputs contain positive values. The SS recovery calculation pairs each measured MS value from the source with the corresponding sidebar SS matrix-spike concentration.

The parser and its validation messages must accept all five delimiters so the displayed guidance matches actual behavior. Invalid lines must identify the line number and restate the accepted input pattern.
