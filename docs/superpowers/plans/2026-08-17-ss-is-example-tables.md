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

### Task 2: Punctuation-compatible SS and IS input guidance

**Files:**
- Modify: `process_lcms_data.py:442-469`
- Modify: `streamlit_app.py:142-149`
- Test: `tests/test_dynamic_rules.py`
- Test: `tests/test_csv_output_and_ui.py`

- [ ] Add a failing parser test that supplies `SS-A, 1`, `SS-B，2`, `SS-C;3`, `SS-D；4`, and `SS-E\t5`, and asserts that all five names and concentrations are accepted.
- [ ] Run `python -m unittest tests.test_dynamic_rules.DynamicRulesTests.test_custom_ss_entries_accepts_all_documented_delimiters -v` and confirm it fails for the Chinese delimiters.
- [ ] Extend the delimiter expression in `parse_custom_ss_entries()` to accept English/Chinese comma, English/Chinese semicolon, and Tab while retaining one-compound-per-line validation.
- [ ] Add failing UI assertions for bilingual punctuation instructions, the prohibition on entering `ppb`, and copyable SS/IS examples.
- [ ] Run the focused UI test and confirm the new assertions fail before changing production text.
- [ ] Update the bilingual SS/IS help and placeholder text so the accepted punctuation and input restrictions are explicit, and clarify that the post-upload compound-by-MS grid controls per-MS values.
- [ ] Run both focused tests and confirm they pass.
- [ ] Run `python -m unittest discover -s tests -v` and `git diff --check` before committing.
