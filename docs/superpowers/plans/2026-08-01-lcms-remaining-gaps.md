# LC-MS Remaining Gaps Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the remaining usability and validation gaps without inventing IS correction data.

**Architecture:** Keep the workbook engine formula-driven for XLSX and numeric for CSV/online preview. Add explicit metadata validation and role detection, unify summary semantics, and make the UI explain the limits of IS correction from concentration-only exports.

**Tech Stack:** Python, openpyxl, pandas, Streamlit, unittest.

---

### Task 1: Add regression tests for role detection and IS-input validation

**Files:**
- Modify: `tests/test_dynamic_rules.py`
- Modify: `tests/test_csv_output_and_ui.py`

- [ ] Add tests that names containing `IS`, `internal standard`, `SS`, or `surrogate` are auto-detected as roles while ordinary analytes remain targets.
- [ ] Add a test that concentration-only input is reported as unable to perform response-ratio IS correction unless explicit response columns/configuration exist.
- [ ] Run the focused tests and confirm they fail before implementation.

### Task 2: Implement robust role detection and explicit IS correction capability reporting

**Files:**
- Modify: `process_lcms_data.py`
- Modify: `streamlit_app.py`

- [ ] Extend `classify_compounds` with conservative role keywords and preserve manual overrides.
- [ ] Add `detect_is_response_inputs` / `validate_is_correction_inputs` that detects response-ratio columns only when the source contains explicit target/IS response data.
- [ ] Make processing reject a requested response-ratio correction when those inputs are absent, with a clear Chinese/English message.
- [ ] Keep the existing MassHunter-already-IS-corrected path unchanged and label the uncorrected path as a sample-volume conversion only.

### Task 3: Unify descriptive-statistics semantics and add non-QAC Demo coverage

**Files:**
- Modify: `process_lcms_data.py`
- Modify: `demo_urine_qac_masshunter.xlsx`
- Modify: `tests/test_summary_sheet.py`

- [ ] Ensure the legacy final-concentration statistics do not apply a contradictory DF > 50% gate when the standalone summary is selected as the reporting result.
- [ ] Add one clearly named non-QAC analyte to Demo and verify it is retained as `Other`.
- [ ] Add tests that summary DF is always populated and MQL equals MDL times the configured factor.

### Task 4: Improve format and language diagnostics

**Files:**
- Modify: `streamlit_app.py`
- Modify: `README.md`

- [ ] Add a visible input-format diagnostic for CSV delimiter/encoding and workbook header detection.
- [ ] Add a visible note that CSV is a flat, formula-free export and XLSX is required for multi-sheet formulas.
- [ ] Add a language smoke test for required English labels and a Chinese fallback.

### Task 5: Verify and publish

**Files:**
- No source changes beyond Tasks 1-4.

- [ ] Run all unit tests, syntax checks, and `git diff --check`.
- [ ] Process Demo as XLSX and CSV and verify the non-QAC row, summary sheet, and no formula text in CSV.
- [ ] Commit and push to `origin/master`.
