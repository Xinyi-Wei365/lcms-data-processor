import sys
import unittest

sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.dirname(__file__)))
import process_lcms_data as processor


class NameNormalizationTests(unittest.TestCase):
    def test_ddac_is_displayed_as_dadmac(self):
        self.assertEqual(processor.normalize_analyte_name('C12-DDAC'), 'C12-DADMAC')
        self.assertEqual(processor.normalize_analyte_name('C8-10-DDAC'), 'C8-10-DADMAC')
        self.assertEqual(processor.normalize_analyte_name('C12-DADMAC'), 'C12-DADMAC')

    def test_reader_returns_dadmac_name_and_data(self):
        csv_bytes = (
            'Name,Ion,BLANK1,F1\n'
            'C12-DDAC,382>214,0.1,0.2\n'
        ).encode('utf-8')
        raw, _, _, _, targets, _, _, all_compounds = processor.read_raw(csv_bytes)
        self.assertIn('C12-DADMAC', raw)
        self.assertIn('C12-DADMAC', targets)
        self.assertIn('C12-DADMAC', all_compounds)
        self.assertNotIn('C12-DDAC', raw)


if __name__ == '__main__':
    unittest.main()
