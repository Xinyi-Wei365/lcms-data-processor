import unittest

import process_lcms_data as processor


class BlankMdlEvidenceTests(unittest.TestCase):
    def test_all_zero_blanks_require_calibration_and_signal_to_noise(self):
        evidence = processor.build_blank_mdl_evidence('C8-BAC', [0, 0.0, '0'], {})

        self.assertEqual(evidence['status'], 'blank_zero')
        self.assertEqual(evidence['valid_count'], 3)
        self.assertEqual(evidence['nonzero_count'], 0)
        self.assertFalse(evidence['ready'])
        self.assertIsNone(evidence['mdl'])

    def test_all_zero_blanks_show_substituted_signal_to_noise_formula(self):
        evidence = processor.build_blank_mdl_evidence('C8-BAC', [0, 0, 0], {
            'mdl_overrides': {'C8-BAC': {
                'blank_zero': True,
                'calibration_concentration': 1.0,
                'signal_to_noise': 10.0,
            }}
        })

        self.assertTrue(evidence['ready'])
        self.assertEqual(evidence['formula'], '3 × 1 ÷ 10')
        self.assertAlmostEqual(evidence['mdl'], 0.3)

    def test_nonzero_blanks_use_dynamic_t_and_blank_only_formula(self):
        evidence = processor.build_blank_mdl_evidence(
            'C10-BAC', [0, 0.012, 0.018, None, 0.015, 0.011], {}
        )

        self.assertEqual(evidence['status'], 'blank_nonzero')
        self.assertEqual(evidence['valid_count'], 5)
        self.assertEqual(evidence['nonzero_count'], 4)
        self.assertEqual(evidence['degrees_of_freedom'], 4)
        self.assertAlmostEqual(evidence['t_value'], 3.747, places=3)
        self.assertTrue(evidence['ready'])
        self.assertAlmostEqual(
            evidence['mdl'], evidence['mean'] + evidence['t_value'] * evidence['sd']
        )
        self.assertIn('3.747', evidence['formula'])

    def test_all_missing_blanks_are_not_treated_as_zero(self):
        evidence = processor.build_blank_mdl_evidence('Missing', [None, '', 'ND'], {})

        self.assertEqual(evidence['status'], 'missing')
        self.assertEqual(evidence['valid_count'], 0)
        self.assertFalse(evidence['ready'])

    def test_one_nonzero_blank_is_insufficient_for_standard_deviation(self):
        evidence = processor.build_blank_mdl_evidence('Insufficient', [None, 0.02, None], {})

        self.assertEqual(evidence['status'], 'insufficient')
        self.assertEqual(evidence['valid_count'], 1)
        self.assertFalse(evidence['ready'])

    def test_excel_formula_ignores_legacy_low_spike_values(self):
        formula = processor.mdl_formula('C8-BAC', 'B5:F5', {
            'mdl_spike_values': {'C8-BAC': [0.9, 1.0, 1.1]}
        })

        self.assertEqual(
            formula,
            '=AVERAGE(B5:F5)+T.INV(0.99,COUNT(B5:F5)-1)*STDEV.S(B5:F5)',
        )


if __name__ == '__main__':
    unittest.main()
