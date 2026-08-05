import io
import sys
import unittest

sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.dirname(__file__)))
import process_lcms_data as processor


class CsvOutputAndUiTests(unittest.TestCase):
    def test_export_csv_contains_summary_columns_and_numeric_preview(self):
        raw = (
            'Name,Ion,BLANK1,BLANK2,MS1,F1,F2\n'
            'C8-BAC,248>91,0.1,0.2,10,0.4,0.5\n'
            'd7-C12-BAC,311>98,0.1,0.2,4,4,4\n'
        ).encode('utf-8')
        output, _ = processor.process({
            'input_bytes': raw, 'input_file': '', 'output_format': 'csv',
            'output_file': 'processed.csv', 'sample_type': '尿液',
            'is_compounds': [], 'ss_compounds': ['d7-C12-BAC'],
            'ss_spike_concentrations': {'d7-C12-BAC': 4},
            'spike_conc_ppb': 10, 'conversion_factor': 1,
        }, return_bytes=True)
        self.assertIn(b'Final concentration summary', output)
        self.assertIn(b'Median (Q1-Q3)', output)
        self.assertNotIn(b'=', output)

    def test_language_labels_are_available_in_both_languages(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        self.assertIn("'zh': '上传原始数据（XLSX 或 CSV）'", source)
        self.assertIn("'en': 'Upload Raw Data (XLSX or CSV)'", source)
        self.assertIn("'en': 'Compound Roles'", source)
        self.assertIn("'en': 'Blank-zero MDL settings'", source)

    def test_ui_keeps_one_is_correction_question(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        self.assertNotIn("'is_mode'", source)
        self.assertNotIn('is_correction_mode', source)
        self.assertNotIn('response_factors', source)
        self.assertEqual(source.count("'is_correction':"), 1)

    def test_ui_exposes_custom_ss_input_with_example_and_validation(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        self.assertIn("'custom_ss'", source)
        self.assertIn('d7-C12-BAC, 4', source)
        self.assertIn('parse_custom_ss_entries(custom_ss_text)', source)
        self.assertIn('missing_custom_ss', source)

    def test_ui_exposes_custom_is_concentration_input_and_keeps_correction_separate(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        self.assertIn("'custom_is'", source)
        self.assertIn('parse_custom_ss_entries(custom_is_text)', source)
        self.assertIn('is_spike_concentrations', source)
        self.assertIn("'is_corrected': is_corrected", source)

    def test_ui_does_not_show_fixed_d7_and_d9_surrogate_concentration_boxes(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        self.assertNotIn("ss_spike_d7 = st.number_input", source)
        self.assertNotIn("ss_spike_d9 = st.number_input", source)
        self.assertNotIn("'ss_spike_d7_ppb':", source)
        self.assertNotIn("'ss_spike_d9_ppb':", source)

    def test_custom_ss_input_is_visible_in_sidebar_before_file_upload(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        sidebar_start = source.index('with st.sidebar:')
        main_region = source.index('# 主区域')
        sidebar_source = source[sidebar_start:main_region]
        self.assertIn('custom_ss_text = st.text_area(', sidebar_source)
        self.assertIn("key='custom_ss_text'", sidebar_source)

    def test_ui_preview_accepts_legacy_xls_uploads(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        self.assertIn("file_bytes.startswith(b'\\xd0\\xcf\\x11\\xe0')", source)
        self.assertIn("pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd')", source)

    def test_ui_shows_a_table_of_type_chain_and_role_for_detected_compounds(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        self.assertIn('compound_classification_rows', source)
        self.assertIn('classification_rows', source)
        self.assertIn('pd.DataFrame(classification_rows)', source)

    def test_custom_ss_spike_value_is_shown_as_the_value_used_for_calculation(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        self.assertIn('value=float(custom_ss.get(name, 4.0))', source)
        self.assertNotIn('if name in custom_ss:\n                        ss_concentrations[name] = custom_ss[name]', source)

    def test_classification_table_updates_after_user_confirms_is_and_ss_roles(self):
        with open(__import__('os').path.join(__import__('os').path.dirname(__file__), '..', 'streamlit_app.py'), encoding='utf-8-sig') as handle:
            source = handle.read()
        self.assertIn('compound_classification_rows(all_c, selected_is, selected_ss)', source)

    def test_is_addition_concentrations_are_recorded_without_changing_final_concentration_rule(self):
        raw = (
            'Name,Ion,BLANK1,BLANK2,MS1,F1\n'
            'C8-BAC,248>91,0.1,0.2,10,0.5\n'
            'My Internal Standard,300>100,0.1,0.2,4,4\n'
        ).encode('utf-8')
        output, _ = processor.process({
            'input_bytes': raw, 'input_file': '', 'output_file': 'is_record.xlsx',
            'is_compounds': ['My Internal Standard'],
            'is_spike_concentrations': {'My Internal Standard': 4},
            'is_corrected': True,
        }, return_bytes=True)
        import openpyxl
        workbook = openpyxl.load_workbook(io.BytesIO(output), data_only=False)
        info = workbook['计算说明']
        rows = [[info.cell(r, c).value for c in range(1, 5)] for r in range(1, info.max_row + 1)]
        self.assertIn(['IS addition record', 'My Internal Standard', '4 ppb', 'Recorded only; IS correction applied: yes. Does not change concentration formulas.'], rows)


if __name__ == '__main__':
    unittest.main()
