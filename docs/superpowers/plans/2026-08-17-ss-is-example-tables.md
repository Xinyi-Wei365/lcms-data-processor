# SS/IS Example Tables Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add detailed always-visible SS and IS per-MS concentration examples to the main Streamlit interface.

**Architecture:** Define bilingual labels and pure example-row helpers in `streamlit_app.py`, render two disabled dataframes before upload, and leave the uploaded-file editor unchanged.

**Tech Stack:** Python, Streamlit, pandas, unittest.

---

### Task 1: Example table UI

**Files:**
- Modify: `streamlit_app.py`
- Test: `tests/test_csv_output_and_ui.py`

- [ ] Write failing assertions for both example tables, exact rows, bilingual explanations, and pre-upload visibility.
- [ ] Run the focused test and confirm it fails.
- [ ] Add bilingual labels and render the two detailed dataframes on the main interface.
- [ ] Run the focused test and confirm it passes.
- [ ] Run the full suite and Streamlit pre-upload/Demo smoke tests.

### Task 2: Name-only SS/IS input and dynamic MS concentrations

**Files:**
- Modify: `process_lcms_data.py:442-469`
- Modify: `streamlit_app.py:142-149`
- Test: `tests/test_dynamic_rules.py`
- Test: `tests/test_csv_output_and_ui.py`

- [ ] Add a failing parser test that supplies names separated by English/Chinese comma, English/Chinese semicolon, Tab, and newline, and asserts that all names are accepted and duplicates are removed without reordering.
- [ ] Run the focused parser tests and confirm they fail because the name-only parser does not exist.
- [ ] Add `parse_compound_name_entries()` for name-only SS/IS input.
- [ ] Add failing UI assertions for bilingual punctuation instructions, copyable SS/IS name examples, dynamic MS columns, and IS exclusion from the concentration editor.
- [ ] Run the focused UI test and confirm the new assertions fail before changing production text.
- [ ] Update the bilingual SS/IS help and placeholder text so both inputs accept names only. Rename SS example and editor columns to matrix-spike concentration, dynamically label all detected MS columns, and exclude IS from the editor.
- [ ] Run both focused tests and confirm they pass.
- [ ] Run `python -m unittest discover -s tests -v` and `git diff --check` before committing.
