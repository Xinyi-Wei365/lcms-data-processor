# SS/IS Example Tables Design

## Goal

Show detailed SS and IS per-MS concentration examples on the main interface before a file is uploaded.

## Confirmed content

The SS example table contains `SS替代物`, `MS1理论加入`, `MS2理论加入`, and `MS3理论加入`, with `d7-C12-BAC` values `4`, `8`, and `12`. The caption states that these are theoretical additions used as the denominator of SS recovery; measured MS concentrations are read automatically from the uploaded source.

The IS example table contains `IS内标`, `MS1加入浓度（ppb）`, `MS2加入浓度（ppb）`, and `MS3加入浓度（ppb）`, with `IS-A` values `5`, `5`, `10` and `IS-B` values `2`, `4`, `4`. The caption states that these are experimental additions recorded in the output and do not calculate recovery.

Both tables are bilingual and always visible. The actual compound-by-MS editor remains file-driven and uses the uploaded file's real MS columns.
