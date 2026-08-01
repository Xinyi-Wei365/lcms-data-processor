# LC-MS/MS Generalization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generalize the LC-MS/MS processor from the fixed QAC Demo layout to compatible CSV/XLSX inputs, configurable analyte roles and MDL rules, numeric preview, and a formula-driven descriptive-statistics sheet.

**Architecture:** Keep the existing workbook output formulas and layout as the calculation backend, but replace fixed input assumptions with a normalized tabular reader that returns detected analytes, chain metadata, blanks, matrix spikes, and samples. Streamlit will expose detected-role and method-rule configuration, then pass a serializable configuration into the processor. A separate summary sheet will reference the final-concentration and MDL sheets while the web preview uses independently calculated display values.

**Tech Stack:** Python 3, Streamlit, openpyxl, pandas, unittest.

---

### Task 1: Unified input detection

**Files:**
- Modify: `process_lcms_data.py`
- Modify: `streamlit_app.py`
- Test: `tests/test_input_detection.py`

- [x] Support CSV bytes with encoding/delimiter detection.
- [x] Support XLSX workbooks whose data sheet is not named `Sheet1`.
- [x] Normalize DDAC input names to DADMAC while preserving all associated data.
- [ ] Allow the Streamlit uploader to accept `.csv` and `.xlsx`.
- [ ] Display detected column counts before processing.

### Task 2: Dynamic analyte roles and ordering

**Files:**
- Modify: `process_lcms_data.py`
- Modify: `streamlit_app.py`
- Test: `tests/test_dynamic_roles.py`

- [ ] Build target/IS/SS lists from detected rows plus user selections, not constants.
- [ ] Sort known type and chain labels while retaining unknown compounds in an `Other` section.
- [x] Preserve associated data while displaying DDAC as DADMAC.
- [ ] Pass selected SS concentrations by exact analyte name.

### Task 3: Blank-zero S/N MDL

**Files:**
- Modify: `process_lcms_data.py`
- Modify: `streamlit_app.py`
- Test: `tests/test_mdl_rules.py`

- [ ] Add per-analyte blank status, calibration concentration, and S/N inputs.
- [ ] Use `3 * STDEVA(blank)` when valid blank variation exists.
- [ ] Use `3 * calibration_concentration / signal_to_noise` when configured for blank-zero.
- [ ] Reject missing or non-positive S/N inputs instead of silently returning zero.

### Task 4: Numeric preview

**Files:**
- Modify: `streamlit_app.py`
- Modify: `process_lcms_data.py`
- Test: `tests/test_preview_values.py`

- [ ] Keep formulas in downloaded XLSX.
- [ ] Render computed numeric values in Streamlit preview.
- [ ] Show clear errors for formula cells that cannot be evaluated.

### Task 5: Descriptive statistics sheet

**Files:**
- Modify: `process_lcms_data.py`
- Test: `tests/test_summary_sheet.py`

- [ ] Add a separate sheet with `名称`, `链长`, `DF (%)`, `Median (Q1-Q3)`, `MDL`, and `MQL`.
- [ ] Reference source sheets with formulas rather than hardcoding derived values.
- [ ] Preserve three significant figures in display formatting.
- [ ] Keep MQL formula configurable until the laboratory rule is confirmed.

### Task 6: Verification and publish

**Files:**
- Modify: `requirements.txt` only if a parser dependency is required.

- [ ] Run all unit tests and Python compilation.
- [ ] Process the real Demo through the full pipeline.
- [ ] Inspect output sheet names, formulas, and numeric preview values.
- [ ] Push only verified changes to `master`.
