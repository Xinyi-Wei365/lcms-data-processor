import sys
import unittest

sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.dirname(__file__)))
import process_lcms_data as processor


class DynamicRulesTests(unittest.TestCase):
    def test_analyte_metadata_normalizes_ddac_and_extracts_type_and_chain(self):
        metadata = processor.analyte_metadata('C12-DDAC')
        self.assertEqual(metadata['name'], 'C12-DADMAC')
        self.assertEqual(metadata['type'], 'DADMAC')
        self.assertEqual(metadata['chain_length'], 'C12')

    def test_dadmac_is_recognized_as_distinct_name_without_renaming_ddac(self):
        metadata = processor.analyte_metadata('C10-DDAC')
        self.assertEqual(metadata['name'], 'C10-DADMAC')
        self.assertEqual(metadata['type'], 'DADMAC')
        targets, _, _ = processor.classify_compounds(['C10-DDAC'])
        self.assertEqual(targets, ['C10-DADMAC'])

    def test_role_selection_uses_all_detected_compounds(self):
        roles = processor.resolve_roles(
            ['C8-BAC', 'Surrogate-X', 'Internal-X'],
            is_compounds=['Internal-X'],
            ss_compounds=['Surrogate-X'],
        )
        self.assertEqual(roles['target_compounds'], ['C8-BAC'])
        self.assertEqual(roles['is_compounds'], ['Internal-X'])
        self.assertEqual(roles['ss_compounds'], ['Surrogate-X'])

    def test_role_keywords_detect_custom_is_and_ss_names(self):
        targets, is_compounds, ss_compounds = processor.classify_compounds([
            'C8-PFAS', 'PFAS internal standard', 'PFAS surrogate'
        ])
        self.assertEqual(targets, ['C8-PFAS'])
        self.assertEqual(is_compounds, ['PFAS internal standard'])
        self.assertEqual(ss_compounds, ['PFAS surrogate'])

    def test_mdl_formula_uses_signal_to_noise_for_blank_zero(self):
        formula = processor.mdl_formula(
            'C12-Other',
            'B4:G4',
            {'mdl_overrides': {'C12-Other': {
                'blank_zero': True,
                'calibration_concentration': 1.0,
                'signal_to_noise': 30.0,
            }}}
        )
        self.assertEqual(formula, '=3*1.0/30.0')

    def test_mdl_report_formula_applies_conversion_factor_once(self):
        formula = processor.mdl_report_formula(
            'C12-Other', 'J5',
            {'conversion_factor': 0.25, 'mdl_overrides': {
                'C12-Other': {
                    'blank_zero': True,
                    'calibration_concentration': 1.0,
                    'signal_to_noise': 30.0,
                }
            }}
        )
        self.assertEqual(formula, '=3*1.0/30.0*0.25')

    def test_mdl_formula_keeps_blank_standard_deviation_rule_by_default(self):
        formula = processor.mdl_formula('C8-BAC', 'B4:G4', {})
        self.assertEqual(formula, '=3*STDEVA(B4:G4)')

    def test_selected_ss_concentration_is_resolved_by_exact_name(self):
        cfg = {'ss_spike_concentrations': {'Surrogate-X': 2.0}}
        self.assertEqual(processor.resolve_ss_spike('Surrogate-X', cfg), 2.0)

    def test_unknown_compounds_are_sorted_by_chain_then_name(self):
        targets, _, _ = processor.classify_compounds(['C12-PFOS', 'C8-PFOS', 'C10-Other'])
        self.assertEqual(targets, ['C8-PFOS', 'C10-Other', 'C12-PFOS'])

    def test_non_isotope_cooh_metabolite_is_not_forced_to_is(self):
        targets, is_compounds, _ = processor.classify_compounds(['C10-BAC-COOH', 'C10-BAC-COOH[C13]'])
        self.assertIn('C10-BAC-COOH', targets)
        self.assertNotIn('C10-BAC-COOH', is_compounds)
        self.assertIn('C10-BAC-COOH[C13]', is_compounds)

    def test_duplicate_detected_names_are_removed(self):
        targets, is_compounds, ss_compounds = processor.classify_compounds(
            ['C8-BAC', 'C8-BAC', 'd7-C12-BAC', 'd7-C12-BAC']
        )
        self.assertEqual(targets, ['C8-BAC'])
        self.assertEqual(is_compounds, [])
        self.assertEqual(ss_compounds, ['d7-C12-BAC'])

    def test_blank_zero_detection_requires_all_blank_cells_to_be_numeric_zero(self):
        raw_data = {
            'Zero': {'B': 0, 'C': 0.0, 'D': '0'},
            'Mixed': {'B': 0, 'C': 0.1, 'D': 0},
            'Missing': {'B': 0, 'C': None, 'D': 0},
        }
        blanks = [(2, 'B', 'blank_1'), (3, 'C', 'blank_2'), (4, 'D', 'blank_3')]
        self.assertEqual(processor.detect_blank_zero_compounds(raw_data, blanks), ['Zero'])

    def test_blank_zero_mdl_requires_manual_snr_override(self):
        with self.assertRaises(ValueError):
            processor.validate_blank_zero_mdl('Zero', [0, 0, 0], {})
        self.assertIsNone(processor.validate_blank_zero_mdl(
            'Zero', [0, 0, 0],
            {'mdl_overrides': {'Zero': {
                'blank_zero': True,
                'calibration_concentration': 1,
                'signal_to_noise': 30,
            }}}
        ))

    def test_blank_zero_configuration_validation_covers_all_target_compounds(self):
        raw_data = {'Zero': {'B': 0, 'C': 0}, 'Nonzero': {'B': 0, 'C': 0.1}}
        blanks = [(2, 'B', 'blank_1'), (3, 'C', 'blank_2')]
        with self.assertRaises(ValueError):
            processor.validate_blank_zero_configuration(raw_data, blanks, ['Zero'], {})
        processor.validate_blank_zero_configuration(
            raw_data, blanks, ['Zero'],
            {'mdl_overrides': {'Zero': {
                'blank_zero': True,
                'calibration_concentration': 1,
                'signal_to_noise': 30,
            }}}
        )


if __name__ == '__main__':
    unittest.main()
