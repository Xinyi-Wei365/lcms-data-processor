# Blank/MDL Calculation Evidence Design

## Goal

Add a main-interface calculation-evidence area that lets the user audit the Blank/MDL decision and result for every target compound before processing.

## Confirmed calculation rules

- Blank status is determined independently for each target compound.
- If every blank cell is numeric zero, the compound is `blank=0`. The user supplies the calibration concentration and corresponding S/N, and vial MDL is `3 × calibration concentration ÷ S/N`.
- If at least one valid blank value is non-zero and at least two valid blank values exist, the compound is `blank≠0`. Vial MDL is `mean(blank) + t(0.99,n-1) × SD(blank)`, with the one-sided 99% t value selected from the actual number of valid blanks.
- Empty cells are missing values, not numeric zero.
- If all blank cells are empty, or fewer than two valid values are available for the non-zero path, the interface reports why MDL cannot be calculated and does not invent a result.
- The former low-level-spike-replicate input and MAX branch are removed from the interface and calculation path.

## Interface

After compound-role classification and matrix-spike settings, show a Blank/MDL section for target compounds. It contains a compact status table with compound, blank values, valid count, non-zero count, status, and required action.

Below it, add a `Calculation evidence` / `计算依据` expander. Each compound displays:

- parsed blank values;
- valid blank count and non-zero count;
- selected path (`blank=0`, `blank≠0`, or unable to calculate);
- for `blank=0`, calibration concentration, S/N, substituted formula, and vial MDL preview;
- for `blank≠0`, degrees of freedom, one-sided 99% t value, blank mean, sample SD, substituted formula, and vial MDL preview;
- a clear validation message when calculation is unavailable.

The evidence display, online numeric previews, and exported Excel formulas must share the same calculation helpers.

## Validation

- Blank-zero compounds require positive calibration concentration and S/N before processing.
- Blank-nonzero compounds require at least two valid blank results.
- Missing blanks remain missing and are never converted to zero.
- Tests cover the three states, dynamic t selection, evidence values, Excel formula generation, bilingual UI labels, and removal of the low-spike interface.
