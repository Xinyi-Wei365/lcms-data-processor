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
        self.assertEqual(len(blanks), 4)
        self.assertEqual(len(mss), 2)
        self.assertEqual(len(samps), 5)
        self.assertIn('C8-DADMAC', targets)
        self.assertEqual(len(is_comps), 1)
        self.assertEqual(len(ss_comps), 2)
        self.assertTrue(all('DDAC' not in name for name in raw))
        self.assertIn('C8-DADMAC', raw)

    def test_demo_runs_with_default_roles_and_ms_spike_settings(self):
        demo = Path(__file__).resolve().parents[1] / 'demo_urine_qac_masshunter.xlsx'
        raw, _, mss, _, _, is_comps, ss_comps, _ = processor.read_raw(demo)
        output, _ = processor.process({
            'input_bytes': demo.read_bytes(), 'input_file': '',
            'is_compounds': is_comps, 'ss_compounds': ss_comps,
            'matrix_spike_concentrations': {
                name: {header: 4 for _, _, header in mss}
                for name in [*is_comps, *ss_comps]
            },
        }, return_bytes=True)
        self.assertGreater(len(output), 1000)
        self.assertIn('C8-DADMAC', raw)

    def test_streamlit_uses_only_the_regenerated_demo_workbook(self):
        app_source = (Path(__file__).resolve().parents[1] / 'streamlit_app.py').read_text(encoding='utf-8-sig')
        self.assertNotIn('def get_demo_bytes', app_source)
        self.assertIn("demo_urine_qac_masshunter.xlsx", app_source)

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

    def test_process_rejects_a_file_without_required_blank_or_sample_columns(self):
        raw = (
            'Compound Name,Transition,MatrixSpike_1\n'
            'C8-PFAS,499>80,1.0\n'
        ).encode('utf-8')

        with self.assertRaisesRegex(ValueError, 'BLANK|sample'):
            processor.process({
                'input_bytes': raw,
                'input_file': '',
                'output_file': 'invalid.xlsx',
            }, return_bytes=True)


if __name__ == '__main__':
    unittest.main()
