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
