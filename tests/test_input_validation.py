import sys
import unittest

sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.dirname(__file__)))
import process_lcms_data as processor
from pathlib import Path


class InputValidationTests(unittest.TestCase):
    def test_demo_file_is_a_directly_usable_masshunter_layout(self):
        demo = Path(__file__).resolve().parents[1] / 'demo_urine_qac_masshunter.xlsx'
        raw, blanks, mss, samps, targets, is_comps, ss_comps, _ = processor.read_raw(demo)
        report = processor.validate_input_layout(blanks, mss, samps, targets, is_comps, ss_comps)
        self.assertTrue(report['ready'])
        self.assertEqual(len(blanks), 6)
        self.assertEqual(len(mss), 2)
        self.assertEqual(len(samps), 10)
        self.assertIn('C8-DADMAC', targets)
        self.assertEqual(len(is_comps), 2)
        self.assertEqual(len(ss_comps), 2)
        self.assertTrue(all('DDAC' not in name for name in raw))
        self.assertIn('C12-DADMAC', raw)

    def test_validation_reports_a_ready_demo_layout(self):
        report = processor.validate_input_layout(
            blanks=[(3, 'C', 'F91-BLANK1')] * 6,
            mss=[(10, 'J', 'F89-MS1'), (11, 'K', 'F90-MS2')],
            samps=[(12, 'L', 'F1'), (13, 'M', 'F2')],
            target_compounds=['C8-BAC'],
            is_compounds=['C10-BAC[C13]'],
            ss_compounds=['d7-C12-BAC'],
        )
        self.assertTrue(report['ready'])
        self.assertEqual(report['errors'], [])
        self.assertIn('6 BLANK', report['summary'])

    def test_validation_flags_missing_blank_and_sample_columns(self):
        report = processor.validate_input_layout(
            blanks=[], mss=[], samps=[], target_compounds=['C8-BAC'],
            is_compounds=[], ss_compounds=[],
        )
        self.assertFalse(report['ready'])
        self.assertTrue(any('BLANK' in message for message in report['errors']))
        self.assertTrue(any('sample' in message.lower() for message in report['errors']))


if __name__ == '__main__':
    unittest.main()
