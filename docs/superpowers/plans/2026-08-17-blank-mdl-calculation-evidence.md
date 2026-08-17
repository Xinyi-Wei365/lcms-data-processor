# Blank/MDL Calculation Evidence Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add auditable, per-compound Blank/MDL calculation evidence and remove the obsolete low-level-spike input path.

**Architecture:** Put blank classification and numeric evidence in pure helpers in `process_lcms_data.py`. Reuse those helpers for Streamlit display, preview calculations, validation, and Excel formula generation so all outputs implement the same rule.

**Tech Stack:** Python, Streamlit, openpyxl, unittest.

---

### Task 1: Unified Blank/MDL evidence helper

**Files:**
- Modify: `process_lcms_data.py`
- Test: `tests/test_blank_mdl_evidence.py`

- [ ] Write failing tests for all-zero, non-zero, missing, and insufficient blank states.
- [ ] Run `python -m unittest tests.test_blank_mdl_evidence -v` and verify failures are caused by the missing helper.
- [ ] Implement `build_blank_mdl_evidence()` with parsed values, counts, status, mean, SD, degrees of freedom, t value, formula text, and vial MDL.
- [ ] Make preview and Excel MDL paths use the blank-only non-zero rule.
- [ ] Run the focused tests and verify they pass.

### Task 2: Calculation-evidence interface

**Files:**
- Modify: `streamlit_app.py`
- Test: `tests/test_csv_output_and_ui.py`

- [ ] Write failing source-level UI assertions for bilingual labels, the evidence expander, per-compound status table, and absence of the low-spike input.
- [ ] Run the focused UI tests and verify failure.
- [ ] Replace the old blank-zero multiselect and low-spike textbox with per-target automatic classification, blank-zero C/SN inputs, and calculation evidence.
- [ ] Disable processing when an applicable target has unresolved MDL evidence.
- [ ] Run the focused UI tests and verify they pass.

### Task 3: Regression and end-to-end verification

**Files:**
- Modify: tests whose expected legacy MDL rule conflicts with the confirmed rule.
- Verify: `demo_urine_qac_masshunter.xlsx`

- [ ] Run the complete unit-test suite.
- [ ] Process the Demo to XLSX and CSV and confirm numeric preview calculations succeed.
- [ ] Inspect representative Blank/MDL and descriptive-statistics formulas.
- [ ] Run Python syntax checks and `git diff --check`.
- [ ] Commit only the feature files, preserving the user's existing `README.md` modification.
