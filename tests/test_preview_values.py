import sys
import unittest

sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.dirname(__file__)))
import process_lcms_data as processor


class PreviewValueTests(unittest.TestCase):
    def test_numeric_preview_calculates_final_concentration_and_statistics(self):
        raw_data = {
            'C8-BAC': {
                'B': 1.0, 'C': 1.2, 'D': 0.8,
                'E': 2.0, 'F': 3.0,
            }
        }
        blank_cols = [(2, 'B', 'blank_1'), (3, 'C', 'blank_2'), (4, 'D', 'blank_3')]
        sample_cols = [(5, 'E', 'F1'), (6, 'F', 'F2')]
        cfg = {
            'target_compounds': ['C8-BAC'],
            'conversion_factor': 1.0,
            'mql_factor': 3.333333,
        }

        rows = processor.compute_preview_summary(raw_data, blank_cols, sample_cols, cfg)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['名称'], 'C8-BAC')
        self.assertEqual(rows[0]['链长'], 'C8')
        self.assertAlmostEqual(rows[0]['DF (%)'], 50.0)
        # With the confirmed rule, the vial MDL is 2.393 ppb.  F1 (2.0)
        # is therefore a valid non-detect and becomes 1/2 MDL; F2 (3.0) is
        # a true detection and becomes 3.0 - blank mean (1.0) = 2.0.
        self.assertEqual(rows[0]['Median (Q1-Q3)'], '1.6 (1.4-1.8)')
        self.assertAlmostEqual(rows[0]['MDL'], 2.39, places=2)
        self.assertAlmostEqual(rows[0]['MQL'], 3.0, places=6)

    def test_numeric_final_preview_has_sample_values_without_formula_text(self):
        raw_data = {
            'C8-BAC': {'B': 1.0, 'C': 1.0, 'D': 1.0, 'E': 2.0, 'F': 3.0},
        }
        blank_cols = [(2, 'B', 'blank_1'), (3, 'C', 'blank_2'), (4, 'D', 'blank_3')]
        sample_cols = [(5, 'E', 'F1'), (6, 'F', 'F2')]
        rows = processor.compute_preview_final_table(
            raw_data, blank_cols, sample_cols,
            {'target_compounds': ['C8-BAC'], 'conversion_factor': 1.0}
        )
        # In this fixture blank SD is zero, so vial MDL is 1.0 and both
        # samples are true detections: (2-1)=1 and (3-1)=2.
        self.assertEqual(rows, [{'名称': 'C8-BAC', 'F1': 1.0, 'F2': 2.0}])

    def test_df_counts_only_true_detections_not_half_mdl_substitutions(self):
        raw_data = {
            'C8-BAC': {
                'B': 1.0, 'C': 1.0, 'D': 1.0,
                'E': 0.5, 'F': 0.6, 'G': 2.0, 'H': None,
            }
        }
        blank_cols = [(2, 'B', 'blank_1'), (3, 'C', 'blank_2'), (4, 'D', 'blank_3')]
        sample_cols = [(5, 'E', 'F1'), (6, 'F', 'F2'), (7, 'G', 'F3'), (8, 'H', 'F4')]
        rows = processor.compute_preview_summary(
            raw_data, blank_cols, sample_cols,
            {'target_compounds': ['C8-BAC'], 'conversion_factor': 1.0, 'mql_factor': 3.333333}
        )

        # Only F3 is a true detection. F1/F2 are 1/2 MDL substitutions and
        # F4 is missing, so DF is 1 true detection out of 3 valid samples.
        self.assertAlmostEqual(rows[0]['DF (%)'], 33.3)
        # DF is shown independently. All available final concentrations,
        # including 1/2 MDL substitutions, are summarized regardless of DF.
        self.assertEqual(rows[0]['Median (Q1-Q3)'], '0.5 (0.5-0.75)')

    def test_blank_zero_preview_reports_mdl_and_mql_in_sample_units(self):
        raw_data = {'C12-Other': {'B': 0, 'C': 0, 'D': 0, 'E': 2}}
        blanks = [(2, 'B', 'blank_1'), (3, 'C', 'blank_2'), (4, 'D', 'blank_3')]
        samples = [(5, 'E', 'F1')]
        rows = processor.compute_preview_summary(raw_data, blanks, samples, {
            'target_compounds': ['C12-Other'],
            'conversion_factor': 0.25,
            'mql_factor': 3.333333,
            'mdl_overrides': {'C12-Other': {
                'blank_zero': True,
                'calibration_concentration': 1.0,
                'signal_to_noise': 30.0,
            }},
        })
        self.assertAlmostEqual(rows[0]['MDL'], 0.025, places=6)
        self.assertAlmostEqual(rows[0]['MQL'], 0.0833, places=3)


if __name__ == '__main__':
    unittest.main()
