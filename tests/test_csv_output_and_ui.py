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


if __name__ == '__main__':
    unittest.main()
