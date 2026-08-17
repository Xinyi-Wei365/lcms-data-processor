# SS/IS Example Tables Design

## Goal

Show detailed SS and IS per-MS concentration examples on the main interface before a file is uploaded.

## Confirmed content

The SS example table contains `SS替代物`, `MS1理论加入`, `MS2理论加入`, and `MS3理论加入`, with `d7-C12-BAC` values `4`, `8`, and `12`. The caption states that these are theoretical additions used as the denominator of SS recovery; measured MS concentrations are read automatically from the uploaded source.

The IS example table contains `IS内标`, `MS1加入浓度（ppb）`, `MS2加入浓度（ppb）`, and `MS3加入浓度（ppb）`, with `IS-A` values `5`, `5`, `10` and `IS-B` values `2`, `4`, `4`. The caption states that these are experimental additions recorded in the output and do not calculate recovery.

Both tables are bilingual and always visible. The actual compound-by-MS editor remains file-driven and uses the uploaded file's real MS columns.

## User input punctuation guidance

The SS and IS custom-name input areas must show the exact accepted syntax, not only the rendered example tables. Each line represents one compound and contains a compound name followed by one positive numeric default concentration. The name and concentration may be separated by any one of these delimiters:

- English comma: `,`
- Chinese comma: `，`
- English semicolon: `;`
- Chinese semicolon: `；`
- Tab character

The interface must tell users not to enter quotation marks, table headers, or the text/unit `ppb` in the concentration field. It must include copyable SS and IS examples and explain that the single concentration entered here is only the initial/default value; after upload, the compound-by-MS table is the source of truth for different MS1, MS2, and MS3 additions.

The parser and its validation messages must accept all five delimiters so the displayed guidance matches actual behavior. Invalid lines must identify the line number and restate the accepted input pattern.
