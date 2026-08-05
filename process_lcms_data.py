#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LC-MS/MS 数据处理智能体 v2
从 MassHunter 导出的原始数据生成最终计算浓度表格

用法：
  python process_lcms_data.py              # 使用 CONFIG 中的参数直接运行
  from process_lcms_data import process     # 作为模块导入，供 Streamlit 调用
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import statistics
import math
import re
import io
import csv
import os
import codecs
import pandas as pd

# ============================================================
# 可配置参数
# ============================================================
CONFIG = {
    'sample_type': '尿液',
    'sample_volume_ml': 2,
    'final_volume_ml': 0.5,
    'extra_dilution': 1,
    'conversion_factor': 1,
    'spike_conc_ppb': 10,
    'masshunter_unit': 'ppb',
    'output_unit': 'ng/mL',
    'blank_handling': 'ND',
    'input_file': r'C:\Users\HP-PC\Desktop\未整理 - 副本.xlsx',
    'output_file': r'C:\Users\HP-PC\Desktop\已处理_尿液数据.xlsx',
}

# ============================================================
# 样式常量
# ============================================================
FONT_COLOR = '384350'
HEADER_BG = 'E1E3E5'
BLUE_FONT = '1F4E78'
BLUE_BG = 'D9EAF7'
GOLD_FONT = '7F6000'
GOLD_BG = 'FFF2CC'
RED_FONT = 'C00000'
YELLOW_BG = 'FFFF00'

# ============================================================
# 化合物分组（按输出顺序）
# ============================================================
BAC_SERIES   = ['C8-BAC','C10-BAC','C12-BAC','C14-BAC','C16-BAC','C18-BAC']
DDAC_SERIES  = ['C8-DADMAC','C8-10-DADMAC','C10-DADMAC','C12-DADMAC','C14-DADMAC','C16-DADMAC','C18-DADMAC']
ATMAC_SERIES = ['C8-ATMAC','C10-ATMAC','C12-ATMAC','C14-ATMAC','C16-ATMAC','C18-ATMAC']
BAC_COOH     = ['C10-BAC+2O-2H','C12-BAC+2O-2H','C14-BAC+2O-2H']
BAC_OH       = ['C10-BAC+O','C12-BAC+O','C14-BAC+O']
IS_COMPS     = ['C10-BAC-COOH[C13]','C12-BAC+COOH-d6','C12-BAC+OH-d3']
SS_COMPS     = ['d7-C12-BAC','d9-C10-ATMAC']

TARGET_COMPS = BAC_SERIES + DDAC_SERIES + ATMAC_SERIES + BAC_COOH + BAC_OH
ALL_COMPS    = TARGET_COMPS + IS_COMPS + SS_COMPS


# ============================================================
# 工具函数
# ============================================================
def safe_float(v):
    if v is None: return None
    try: return float(v)
    except: return None


def normalize_analyte_name(name):
    """Normalize the requested DDAC display spelling to DADMAC."""
    original = str(name or '').strip()
    return re.sub(r'(?i)(?<![A-Z])DDAC(?![A-Z])', 'DADMAC', original)

def round6(v):
    if v is None: return None
    return round(float(v), 6)

def round_int(v):
    if v is None: return None
    return round(float(v))


def extract_chain_length(name):
    """Return the first C-number chain label without changing the analyte name."""
    match = re.search(r'(?<![A-Za-z])C(\d+)(?!\d)', str(name or ''), flags=re.I)
    return f'C{match.group(1)}' if match else None


def infer_analyte_type(name):
    """Infer a sortable chemical family from the imported analyte name.

    QAC family names keep their established labels.  For other chemicals with
    a chain-length prefix (for example ``C8-PFAS`` or ``C10-Phthalate``), the
    text after the chain length becomes a family label.  Unknown free-text
    names remain ``Other`` rather than being silently treated as QACs.
    """
    upper = normalize_analyte_name(name).upper()
    if 'ATMAC' in upper:
        return 'ATMAC'
    if 'DADMAC' in upper:
        return 'DADMAC'
    if 'BAC' in upper:
        return 'BAC'

    # Prefer well-known multi-character families when present anywhere in a
    # method name, including names without an explicit C-number prefix.
    for family in ('PFOS', 'PFOA', 'PFHxS', 'PFNA', 'PFDA', 'PFAS', 'PAH', 'PCB'):
        if family.upper() in upper:
            return family.upper()

    # Generic non-QAC family: capture the chemical-type part after C<number>.
    # Remove isotope/oxidation decorations so related analytes stay together.
    match = re.search(r'(?:^|[-_\s])C\d+(?:-\d+)?[-_\s]+([A-Za-z][A-Za-z0-9 ]*)', upper)
    if match:
        family = re.split(r'\s*(?:\+|-)(?:\d*O|\d*H|D\d+|\[.*)', match.group(1), maxsplit=1)[0]
        family = re.sub(r'\s+', ' ', family).strip(' -_')
        if family and family not in {'OTHER', 'UNKNOWN'}:
            return family
    return 'Other'


def analyte_metadata(name):
    """Return display-safe metadata while preserving the original analyte name."""
    original = normalize_analyte_name(name)
    return {
        'name': original,
        'type': infer_analyte_type(original),
        'chain_length': extract_chain_length(original),
    }


def resolve_roles(compounds, is_compounds=None, ss_compounds=None):
    """Resolve roles from the detected list; unselected compounds remain targets."""
    all_compounds = list(dict.fromkeys(normalize_analyte_name(value) for value in compounds if str(value).strip()))
    is_set = {normalize_analyte_name(value) for value in (is_compounds or [])}
    ss_set = {normalize_analyte_name(value) for value in (ss_compounds or [])}
    overlap = is_set & ss_set
    if overlap:
        raise ValueError(f'An analyte cannot be both IS and SS: {sorted(overlap)}')
    is_list = [name for name in all_compounds if name in is_set]
    ss_list = [name for name in all_compounds if name in ss_set]
    targets = [name for name in all_compounds if name not in is_set and name not in ss_set]
    return {'target_compounds': targets, 'is_compounds': is_list, 'ss_compounds': ss_list}


def configured_compound_lists(cfg, detected_compounds, detected_is=None, detected_ss=None):
    """Return target/IS/SS lists using user selections with detected-role defaults."""
    roles = resolve_roles(
        detected_compounds,
        cfg.get('is_compounds', detected_is or []),
        cfg.get('ss_compounds', detected_ss or []),
    )
    return roles['target_compounds'], roles['is_compounds'], roles['ss_compounds'], list(detected_compounds)


def compound_classification_rows(compounds, is_compounds=None, ss_compounds=None):
    """Create an inspectable classification table for the Streamlit interface.

    The imported names remain unchanged apart from the requested DDAC ->
    DADMAC display normalization.  Roles always come from the user-confirmed
    IS/SS selections, while targets are sorted by type and chain length.
    """
    roles = resolve_roles(compounds, is_compounds, ss_compounds)
    rows = []
    for role, items in (
        ('目标物', sort_compounds(roles['target_compounds'])),
        ('IS', sort_compounds(roles['is_compounds'])),
        ('SS', sort_compounds(roles['ss_compounds'])),
    ):
        for compound in items:
            metadata = analyte_metadata(compound)
            rows.append({
                '名称': metadata['name'],
                '类型': metadata['type'],
                '链长': metadata['chain_length'] or 'NA',
                '角色': role,
            })
    return rows


def sort_compounds(compounds):
    """Sort detected analytes by recognized family, chain length, then name."""
    family_order = {'BAC': 0, 'DADMAC': 1, 'ATMAC': 2}
    def family_sort_key(name):
        analyte_type = analyte_metadata(name)['type']
        if analyte_type in family_order:
            return family_order[analyte_type], ''
        # Non-QAC analytes are arranged by chain length first, then their
        # inferred type.  This keeps e.g. C8-PFOS, C10-Other, C12-PFOS in a
        # practical carbon-chain order while still exposing their types.
        return 3, ''
    return sorted(
        list(dict.fromkeys(compounds)),
        key=lambda name: (
            *family_sort_key(name),
            int(extract_chain_length(name)[1:]) if extract_chain_length(name) else 10**9,
            analyte_metadata(name)['type'],
            str(name).upper(),
        ),
    )


def _cell_text(value):
    return str(value or '').strip()


def _normalise_header_text(value):
    """Normalize harmless header punctuation from CSV/XLSX conversion tools."""
    return re.sub(r'[_\-]+', ' ', _cell_text(value).lower()).strip()


def _is_compound_header(value):
    return _normalise_header_text(value) in {
        'name', 'analyte', 'analyte name', 'compound', 'compound name',
        'compound id', '名称', '化合物', '化合物名称',
    }


def _looks_like_header_row(row):
    texts = [_normalise_header_text(value) for value in row]
    has_name = any(_is_compound_header(value) for value in row)
    has_data_hint = any(('blank' in value or 'sample' in value or 'ms' in value or value.startswith('f')) for value in texts)
    return has_name and has_data_hint


def _csv_rows(raw_bytes):
    # UTF-16 is common when a CSV is opened and re-saved from Excel.  Prefer
    # BOM-aware decoding before legacy East-Asian encodings.
    for encoding in ('utf-8-sig', 'utf-16', 'utf-8', 'gb18030', 'big5'):
        try:
            text = raw_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            text = None
    if text is None:
        raise ValueError('CSV encoding is not supported; please save as UTF-8, UTF-16, or GB18030.')
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=',;\t')
    except csv.Error:
        dialect = csv.excel
    return [row for row in csv.reader(io.StringIO(text), dialect)]


def _normalise_rows(rows):
    rows = [[value.strip() if isinstance(value, str) else value for value in row] for row in rows]
    header_index = next((i for i, row in enumerate(rows) if _looks_like_header_row(row)), None)
    if header_index is None:
        header_index = 0
    header = rows[header_index]
    body = rows[header_index + 1:]
    name_col = next((i for i, value in enumerate(header) if _is_compound_header(value)), 0)
    data_start = name_col + 1
    if data_start < len(header) and any(token in _cell_text(header[data_start]).lower()
                                        for token in ('ion', 'transition', '离子', '母离子', '子离子')):
        data_start += 1
    return header, body, name_col, data_start


def _read_csv_source(raw_bytes):
    rows = _csv_rows(raw_bytes)
    header, body, name_col, data_start = _normalise_rows(rows)
    return _parse_tabular_rows(header, body, name_col, data_start)


def _parse_tabular_rows(header, body, compound_col, data_start_col):
    max_cols = max([len(header)] + [len(row) for row in body] or [0])
    header = list(header) + [''] * (max_cols - len(header))
    blanks, mss, samps = [], [], []
    for index in range(data_start_col, max_cols):
        hdr = _cell_text(header[index])
        upper = hdr.upper()
        letter = get_column_letter(index + 1)
        if 'BLANK' in upper:
            blanks.append((index + 1, letter, hdr))
        elif re.search(r'\bMS(?:\d+)?\b|MATRIX\s*SPIKE', upper):
            mss.append((index + 1, letter, hdr))
        elif re.search(r'\d+\s*PPB', upper):
            continue
        elif hdr:
            samps.append((index + 1, letter, hdr))

    data = {}
    compounds = []
    for row in body:
        row = list(row) + [None] * (max_cols - len(row))
        name = normalize_analyte_name(_cell_text(row[compound_col]))
        if not name:
            continue
        compounds.append(name)
        data[name] = {get_column_letter(i + 1): row[i] for i in range(data_start_col, max_cols)}
    target, is_c, ss_c = classify_compounds(compounds)
    return data, blanks, mss, samps, target, is_c, ss_c, target + is_c + ss_c


def resolve_ss_spike(name, cfg, ms_column=None):
    """Resolve a surrogate's own spike concentration by exact analyte name."""
    # The compound-by-MS table is the most specific source when supplied.
    if ms_column is not None:
        configured_matrix = cfg.get('matrix_spike_concentrations') or {}
        _, letter, header = ms_column
        per_ms = configured_matrix.get(name)
        if isinstance(per_ms, dict):
            for key in (header, letter):
                value = safe_float(per_ms.get(key))
                if value is not None and value > 0:
                    return value
    configured = cfg.get('ss_spike_concentrations') or {}
    if name in configured:
        value = safe_float(configured[name])
        if value is not None and value > 0:
            return value
    # Backward-compatible defaults for the current Demo only.
    if 'd7' in name.lower():
        return safe_float(cfg.get('ss_spike_d7_ppb', 4))
    if 'd9' in name.lower():
        return safe_float(cfg.get('ss_spike_d9_ppb', 4))
    return safe_float(cfg.get('ss_spike_conc_ppb'))


def resolve_matrix_spike_concentration(compound, ms_column, cfg):
    """Return one compound's concentration in one matrix-spike column.

    The preferred configuration is ``{compound: {MS1: value, MS2: value}}``.
    A flat header map remains supported for old saved configurations.
    """
    _, letter, header = ms_column
    configured = cfg.get('matrix_spike_concentrations') or {}
    per_ms = configured.get(compound)
    if isinstance(per_ms, dict):
        for key in (header, letter):
            value = safe_float(per_ms.get(key))
            if value is not None and value > 0:
                return value
    for key in (header, letter):
        value = safe_float(configured.get(key))
        if value is not None and value > 0:
            return value
    value = safe_float(cfg.get('spike_conc_ppb', 10))
    if value is None or value <= 0:
        raise ValueError(f'{header}: matrix-spike concentration must be positive.')
    return value


def parse_custom_ss_entries(text):
    """Parse one user-defined surrogate per line: name, spike concentration.

    Example: ``d7-C12-BAC, 4``.  The caller validates that each supplied name
    actually exists in the imported MassHunter table before it is used.
    """
    entries = {}
    errors = []
    for line_number, raw_line in enumerate(str(text or '').splitlines(), 1):
        line = raw_line.strip()
        if not line:
            continue
        parts = [part.strip() for part in re.split(r'[,\t;]', line, maxsplit=1)]
        if len(parts) != 2 or not parts[0]:
            errors.append(f'Line {line_number}: enter "name, concentration".')
            continue
        name = normalize_analyte_name(parts[0])
        concentration = safe_float(parts[1])
        if concentration is None or concentration <= 0:
            errors.append(f'Line {line_number}: spike concentration must be positive.')
            continue
        if name in entries:
            errors.append(f'Line {line_number}: duplicate SS name "{name}".')
            continue
        entries[name] = concentration
    return entries, errors


def mdl_formula(name, blank_range, cfg):
    """Return the auditable Excel MDL formula for one analyte."""
    override = (cfg.get('mdl_overrides') or {}).get(name) or {}
    if override.get('blank_zero'):
        concentration = safe_float(override.get('calibration_concentration'))
        signal_to_noise = safe_float(override.get('signal_to_noise'))
        if concentration is None or concentration <= 0:
            raise ValueError(f'{name}: calibration concentration must be positive for S/N MDL.')
        if signal_to_noise is None or signal_to_noise <= 0:
            raise ValueError(f'{name}: signal-to-noise must be positive for S/N MDL.')
        return f'=3*{concentration}/{signal_to_noise}'
    return f'=3*STDEVA({blank_range})'


def mdl_report_formula(name, bottle_ref, cfg):
    """Return MDL in the same sample/report unit as final concentrations."""
    override = (cfg.get('mdl_overrides') or {}).get(name) or {}
    conversion_factor = safe_float(cfg.get('conversion_factor', 1))
    if conversion_factor is None:
        conversion_factor = 1.0
    if override.get('blank_zero'):
        concentration = safe_float(override.get('calibration_concentration'))
        signal_to_noise = safe_float(override.get('signal_to_noise'))
        if concentration is None or concentration <= 0 or signal_to_noise is None or signal_to_noise <= 0:
            raise ValueError(f'{name}: calibration concentration and S/N must be positive.')
        return f'=3*{concentration}/{signal_to_noise}*{conversion_factor:g}'
    return f'={bottle_ref}*{conversion_factor:g}'


def significant_digits_formula(ref, digits=3):
    """Return an Excel formula that displays a value to significant digits."""
    return f'IFERROR(ROUND({ref},{digits}-1-INT(LOG10(ABS({ref})))),0)'


def _round_significant(value, digits=3):
    """Round a numeric value to significant figures without changing its type."""
    value = safe_float(value)
    if value is None or value == 0:
        return 0.0 if value == 0 else None
    places = digits - 1 - int(math.floor(math.log10(abs(value))))
    return round(value, places)


def _format_significant(value, digits=3):
    """Format a number compactly for the human-readable median range."""
    rounded = _round_significant(value, digits)
    if rounded is None:
        return 'NA'
    return f'{rounded:g}'


def _numeric_values(raw_data, compound, columns):
    return [safe_float(raw_data.get(compound, {}).get(column_letter))
            for _, column_letter, _ in columns
            if safe_float(raw_data.get(compound, {}).get(column_letter)) is not None]


def detect_blank_zero_compounds(raw_data, blank_cols):
    """Detect analytes whose complete blank series is numeric zero."""
    detected = []
    for compound in raw_data:
        values = [safe_float(raw_data.get(compound, {}).get(column_letter))
                  for _, column_letter, _ in blank_cols]
        if values and all(value is not None and value == 0 for value in values):
            detected.append(compound)
    return detected


def validate_blank_zero_mdl(compound, blank_values, cfg):
    """Validate that a blank-zero analyte has the required manual S/N rule."""
    values = [safe_float(value) for value in blank_values]
    if not values or not all(value is not None and value == 0 for value in values):
        return None
    override = (cfg.get('mdl_overrides') or {}).get(compound) or {}
    concentration = safe_float(override.get('calibration_concentration'))
    signal_to_noise = safe_float(override.get('signal_to_noise'))
    if not override.get('blank_zero') or concentration is None or concentration <= 0 or signal_to_noise is None or signal_to_noise <= 0:
        raise ValueError(f'{compound}: blank=0 requires positive calibration concentration and S/N.')
    return None


def validate_blank_zero_configuration(raw_data, blank_cols, target_compounds, cfg):
    """Reject processing when any all-zero target lacks manual S/N inputs."""
    for compound in target_compounds:
        blank_values = [raw_data.get(compound, {}).get(column_letter)
                        for _, column_letter, _ in blank_cols]
        validate_blank_zero_mdl(compound, blank_values, cfg)


def _preview_mdl(compound, blank_cols, cfg):
    override = (cfg.get('mdl_overrides') or {}).get(compound) or {}
    if override.get('blank_zero'):
        concentration = safe_float(override.get('calibration_concentration'))
        signal_to_noise = safe_float(override.get('signal_to_noise'))
        if concentration is None or concentration <= 0 or signal_to_noise is None or signal_to_noise <= 0:
            raise ValueError(f'{compound}: calibration concentration and S/N must be positive.')
        return 3 * concentration / signal_to_noise
    blanks = _numeric_values(cfg.get('_raw_data', {}), compound, blank_cols)
    if len(blanks) < 2:
        return None
    return 3 * statistics.stdev(blanks)


def compute_preview_summary(raw_data, blank_cols, sample_cols, cfg):
    """Calculate numeric summary values for Streamlit without relying on Excel recalculation."""
    cfg = dict(cfg or {})
    cfg['_raw_data'] = raw_data
    conversion_factor = safe_float(cfg.get('conversion_factor', 1))
    if conversion_factor is None:
        conversion_factor = 1.0
    mql_factor = safe_float(cfg.get('mql_factor', 3.333333))
    if mql_factor is None or mql_factor <= 0:
        raise ValueError('mql_factor must be positive.')

    rows = []
    compounds = cfg.get('target_compounds') or []
    for compound in compounds:
        blanks = _numeric_values(raw_data, compound, blank_cols)
        blank_average = statistics.mean(blanks) if blanks else None
        mdl = _preview_mdl(compound, blank_cols, cfg)
        report_mdl = mdl * conversion_factor if mdl is not None else None
        half_mdl = (report_mdl / 2) if report_mdl is not None else None
        final_values = []
        true_detections = 0
        valid_samples = 0
        for _, column_letter, _ in sample_cols:
            value = safe_float(raw_data.get(compound, {}).get(column_letter))
            if value is None:
                continue
            valid_samples += 1
            if blank_average is not None and value > blank_average:
                true_detections += 1
                final_values.append((value - blank_average) * conversion_factor)
            elif half_mdl is not None:
                final_values.append(half_mdl)

        df_pct = (true_detections / valid_samples * 100) if valid_samples else 0.0
        # DF reports true detections only. Descriptive concentration values use
        # every available final value, including the approved 1/2 MDL substitutes.
        if final_values:
            ordered = sorted(final_values)
            median = statistics.median(ordered)
            if len(ordered) == 1:
                q1 = q3 = ordered[0]
            else:
                quartiles = statistics.quantiles(ordered, n=4, method='inclusive')
                q1, q3 = quartiles[0], quartiles[2]
            median_iqr = f'{_format_significant(median)} ({_format_significant(q1)}-{_format_significant(q3)})'
        else:
            median_iqr = 'NC'

        rows.append({
            '名称': compound,
            '链长': extract_chain_length(compound) or 'NA',
            'DF (%)': _round_significant(df_pct),
            'Median (Q1-Q3)': median_iqr,
            'MDL': _round_significant(report_mdl),
            'MQL': _round_significant(report_mdl * mql_factor) if report_mdl is not None else None,
        })
    return rows


def compute_preview_final_table(raw_data, blank_cols, sample_cols, cfg):
    """Return numeric final-concentration rows for an online sample preview."""
    cfg = dict(cfg or {})
    conversion_factor = safe_float(cfg.get('conversion_factor', 1)) or 1.0
    rows = []
    for compound in cfg.get('target_compounds') or []:
        blanks = _numeric_values(raw_data, compound, blank_cols)
        blank_average = statistics.mean(blanks) if blanks else None
        mdl = _preview_mdl(compound, blank_cols, cfg)
        half_mdl = mdl / 2 * conversion_factor if mdl is not None else None
        result = {'名称': compound}
        for _, column_letter, header in sample_cols:
            value = safe_float(raw_data.get(compound, {}).get(column_letter))
            if value is None:
                result[header] = None
            elif blank_average is not None and value > blank_average:
                result[header] = round6((value - blank_average) * conversion_factor)
            else:
                result[header] = round6(half_mdl)
        rows.append(result)
    return rows


# ============================================================
# 样式
# ============================================================
def make_styles():
    base = Font(name='宋体', size=11, color=FONT_COLOR)
    bold = Font(name='宋体', size=11, color=FONT_COLOR, bold=True)
    red_bold = Font(name='宋体', size=11, color=RED_FONT, bold=True)
    blue = Font(name='宋体', size=11, color=BLUE_FONT)
    gold = Font(name='宋体', size=11, color=GOLD_FONT)

    hdr_fill  = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type='solid')
    blue_fill = PatternFill(start_color=BLUE_BG, end_color=BLUE_BG, fill_type='solid')
    gold_fill = PatternFill(start_color=GOLD_BG, end_color=GOLD_BG, fill_type='solid')
    yell_fill = PatternFill(start_color=YELLOW_BG, end_color=YELLOW_BG, fill_type='solid')
    no_fill   = PatternFill(fill_type=None)

    ca = Alignment(horizontal='center', vertical='center')
    la = Alignment(horizontal='left', vertical='center')

    return {
        'hdr':  {'font': base, 'fill': hdr_fill,  'alignment': ca},
        'hdrL': {'font': base, 'fill': hdr_fill,  'alignment': la},
        'cmpd': {'font': base, 'fill': no_fill,   'alignment': la},
        'data': {'font': base, 'fill': no_fill,   'alignment': ca},
        'rec':  {'font': red_bold, 'fill': no_fill, 'alignment': ca},
        'stat': {'font': blue, 'fill': blue_fill, 'alignment': ca},
        'statB':{'font': blue, 'fill': blue_fill, 'alignment': ca},  # stat header bold
        'gold': {'font': gold, 'fill': gold_fill, 'alignment': ca},
        'yell': {'font': base, 'fill': yell_fill, 'alignment': ca},
        'yellL':{'font': bold, 'fill': yell_fill, 'alignment': la},
        'bold': {'font': bold, 'fill': no_fill,   'alignment': ca},
    }

def sty(cell, s):  # apply style dict
    for k in ('font','fill','alignment'):
        if s.get(k): setattr(cell, k, s[k])


# ============================================================
# 原始数据读取
# ============================================================
def read_raw(filepath_or_bytes):
    if isinstance(filepath_or_bytes, (bytes, bytearray)):
        raw_bytes = bytes(filepath_or_bytes)
        if not raw_bytes.startswith(b'PK'):
            if raw_bytes.startswith(b'\xd0\xcf\x11\xe0'):
                frame = pd.read_excel(io.BytesIO(raw_bytes), header=None, engine='xlrd')
                rows = frame.where(pd.notna(frame), None).values.tolist()
                header, body, compound_col, data_start = _normalise_rows(rows)
                return _parse_tabular_rows(header, body, compound_col, data_start)
            return _read_csv_source(raw_bytes)
    if isinstance(filepath_or_bytes, bytes):
        wb = openpyxl.load_workbook(io.BytesIO(filepath_or_bytes), data_only=True)
    else:
        if os.fspath(filepath_or_bytes).lower().endswith('.xls'):
            frame = pd.read_excel(filepath_or_bytes, header=None, engine='xlrd')
            rows = frame.where(pd.notna(frame), None).values.tolist()
            header, body, compound_col, data_start = _normalise_rows(rows)
            return _parse_tabular_rows(header, body, compound_col, data_start)
        wb = openpyxl.load_workbook(filepath_or_bytes, data_only=True)
    ws = next((candidate for candidate in wb.worksheets
               if any(_looks_like_header_row(list(row))
                      for row in candidate.iter_rows(min_row=1, max_row=min(candidate.max_row, 20), values_only=True))), None)
    if ws is None and 'Sheet1' in wb.sheetnames:
        ws = wb['Sheet1']
    if ws is None:
        ws = next((candidate for candidate in wb.worksheets
                   if candidate.max_row > 0 and candidate.max_column > 0), None)
    if ws is None:
        wb.close()
        raise ValueError('No non-empty worksheet was found.')

    preview_rows = [list(row) for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=True)]
    header_index = next((i for i, row in enumerate(preview_rows) if _looks_like_header_row(row)), None)
    if header_index is not None:
        # The header is detected from a small preview, but the body must be
        # read through the worksheet's actual last row.
        rows = [list(row) for row in ws.iter_rows(min_row=1, values_only=True)]
        header, body, compound_col, data_start = _normalise_rows(rows)
        result = _parse_tabular_rows(header, body, compound_col, data_start)
        wb.close()
        return result

    compound_col = 2; data_start_col = 4
    for col in range(1, ws.max_column + 1):
        v = str(ws.cell(row=2, column=col).value or '').strip()
        if v == '名称':
            compound_col = col; data_start_col = col + 1
            next_v = str(ws.cell(row=2, column=col+1).value or '').strip()
            if '离子' in next_v or next_v == '':
                data_start_col = col + 2 if col + 2 <= ws.max_column else col + 1
            break

    blanks, mss, samps = [], [], []
    for col in range(data_start_col, ws.max_column + 1):
        hdr = str(ws.cell(row=1, column=col).value or '')
        cl = get_column_letter(col)
        if 'BLANK' in hdr.upper():
            blanks.append((col, cl, hdr))
        elif 'MS' in hdr.upper():
            mss.append((col, cl, hdr))
        elif re.search(r'\dPPB', hdr.upper()):
            continue
        else:
            samps.append((col, cl, hdr))

    data = {}
    for row in range(3, ws.max_row + 1):
        nm = normalize_analyte_name(str(ws.cell(row=row, column=compound_col).value or '').strip())
        if not nm: continue
        data[nm] = {}
        for col in range(data_start_col, ws.max_column + 1):
            data[nm][get_column_letter(col)] = ws.cell(row=row, column=col).value
    # 读取化合物名称用于分类
    compounds_raw = []
    for row in range(3, ws.max_row + 1):
        nm = normalize_analyte_name(str(ws.cell(row=row, column=compound_col).value or '').strip())
        if nm: compounds_raw.append(nm)
    wb.close()

    target, is_c, ss_c = classify_compounds(compounds_raw)
    return data, blanks, mss, samps, target, is_c, ss_c, target + is_c + ss_c


def validate_input_layout(blanks, mss, samps, target_compounds, is_compounds, ss_compounds):
    """Return a human-readable compatibility check for an imported MassHunter file."""
    errors = []
    warnings = []
    if not blanks:
        errors.append('未识别到 BLANK 列；请检查列名是否包含 BLANK。')
    if not samps:
        errors.append('未识别到 sample 列；样品列应使用 F1、F2 或 Sample-1 等名称。')
    if not target_compounds:
        errors.append('未识别到目标化合物行；请检查化合物名称列和数据行。')
    if not mss:
        warnings.append('未识别到 MS/基质加标列；基质加标回收率将无法计算。')
    if len(blanks) < 2:
        warnings.append('BLANK 列少于 2 个；标准差和 MDL 的稳定性需要确认。')
    if not is_compounds:
        warnings.append('未识别到 IS 内标；请在界面中确认是否使用内标校正。')
    if not ss_compounds:
        warnings.append('未识别到 SS 替代物；SS 回收率不会自动生成。')
    all_headers = [item[2].strip().upper() for item in [*blanks, *mss, *samps] if item[2]]
    duplicate_headers = sorted({header for header in all_headers if all_headers.count(header) > 1})
    if duplicate_headers:
        warnings.append('发现重复数据列名：' + ', '.join(duplicate_headers))
    return {
        'ready': not errors,
        'errors': errors,
        'warnings': warnings,
        'summary': f'{len(blanks)} BLANK + {len(mss)} MS + {len(samps)} sample + '
                   f'{len(is_compounds)} IS + {len(ss_compounds)} SS + {len(target_compounds)} 个目标物',
    }


# ============================================================
# 化合物自动分类
# ============================================================
def classify_compounds(compounds):
    bac, ddac, atmac, metabolites, is_list, ss_list, others = [], [], [], [], [], [], []
    for c in list(dict.fromkeys(str(value).strip() for value in compounds if str(value).strip())):
        cn = normalize_analyte_name(c)
        if not cn: continue
        lower = cn.lower()
        if (any(kw in cn for kw in ['[C13]', '-d6', '-d3', 'OH-d3'])
                or re.search(r'\binternal[ _-]?standard\b|\bis\b', lower)):
            is_list.append(cn)
        elif re.match(r'd\d+-', cn) or re.search(r'\bsurrogate\b|\bss\b', lower):
            ss_list.append(cn)
        elif 'DADMAC' in cn.upper(): ddac.append(cn)
        elif 'ATMAC' in cn.upper(): atmac.append(cn)
        elif 'BAC' in cn.upper(): bac.append(cn)
        else: others.append(cn)
    def sk(n):
        nums = re.findall(r'C(\d+)', n)
        return int(nums[0]) if nums else 0
    bac.sort(key=sk); ddac.sort(key=sk); atmac.sort(key=sk)
    metabolites.sort(key=sk); is_list.sort(key=sk); ss_list.sort(key=sk)
    pure_bac = []; meta = []
    for c in bac:
        if '+O' in c or '+2O' in c: meta.append(c)
        else: pure_bac.append(c)
    metabolites = meta + metabolites
    target = sort_compounds(pure_bac + ddac + atmac + metabolites + others)
    return target, is_list, ss_list


# ============================================================
# Sheet 1: 基质加标浓度
# ============================================================
def build_sheet1(wb, raw_data, ms_cols, S, cfg):
    ws = wb.create_sheet('Matrix spike  基质加标浓度')
    n_ms = len(ms_cols)
    # 列布局: A | B~(n_ms) MS data | 空 | Recoveries标签 | n_ms个回收率% | 空 | avg | SD | SE
    # SS 的理论加标浓度仅作为处理参数，用于其独立回收率计算；不增加输出列。
    ms_start = 2
    mid1 = ms_start + n_ms           # 空列
    rec_lbl = mid1 + 1               # "Recoveries"
    rec_start = rec_lbl + 1          # 回收率% 起始
    mid2 = rec_start + n_ms          # 空列
    stat_col = mid2 + 1              # avg
    sd_col = stat_col + 1            # SD
    se_col = stat_col + 2            # SE
    last_col = se_col

    # Row 1
    ws.cell(row=1, column=1, value='化合物方法'); sty(ws.cell(row=1,column=1), S['hdrL'])
    for i in range(n_ms):
        ws.cell(row=1, column=ms_start+i, value=f'matrix spike_{i+1}'); sty(ws.cell(row=1,column=ms_start+i), S['hdr'])
    ws.cell(row=1, column=mid1, value=None)
    ws.cell(row=1, column=rec_lbl, value='Recoveries'); sty(ws.cell(row=1,column=rec_lbl), S['hdr'])
    for i in range(n_ms):
        ws.cell(row=1, column=rec_start+i, value=f'matrix spike_{i+1}'); sty(ws.cell(row=1,column=rec_start+i), S['hdr'])
    ws.cell(row=1, column=mid2, value=None)
    ws.cell(row=1, column=stat_col, value='average'); sty(ws.cell(row=1,column=stat_col), S['hdr'])
    ws.cell(row=1, column=sd_col, value='SD'); sty(ws.cell(row=1,column=sd_col), S['hdr'])
    ws.cell(row=1, column=se_col, value='SE'); sty(ws.cell(row=1,column=se_col), S['hdr'])

    # Row 2
    ws.cell(row=2, column=1, value='分组'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_ms):
        ws.cell(row=2, column=ms_start+i, value='ppb'); sty(ws.cell(row=2,column=ms_start+i), S['hdr'])
    ws.cell(row=2, column=mid1, value=None)
    ws.cell(row=2, column=rec_lbl, value=None)
    for i in range(n_ms):
        ws.cell(row=2, column=rec_start+i, value='%'); sty(ws.cell(row=2,column=rec_start+i), S['hdr'])
    ws.cell(row=2, column=mid2, value=None)
    ws.cell(row=2, column=stat_col, value='%'); sty(ws.cell(row=2,column=stat_col), S['hdr'])
    ws.cell(row=2, column=sd_col, value='%'); sty(ws.cell(row=2,column=sd_col), S['hdr'])
    ws.cell(row=2, column=se_col, value='%'); sty(ws.cell(row=2,column=se_col), S['hdr'])

    rec_cl = get_column_letter(rec_start)
    rec_cr = get_column_letter(rec_start + n_ms - 1)

    # 数据行（跳过 SS，SS 单独放底部）
    row = 3
    all_compounds = cfg.get('all_compounds', ALL_COMPS)
    selected_ss = cfg.get('ss_compounds', SS_COMPS)
    non_ss = [c for c in all_compounds if c not in selected_ss]
    for comp in non_ss:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])

        # MS 原始数据
        for i, (_, cl, _) in enumerate(ms_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=ms_start+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])

        ws.cell(row=row, column=mid1, value=None)
        ws.cell(row=row, column=rec_lbl, value=None)

        # 回收率 % = (MS值/spike)*100, 四舍五入取整
        for i, ms_col in enumerate(ms_cols):
            _, cl, _ = ms_col
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=rec_start+i)
            if v is not None:
                c.value = round_int((v / resolve_matrix_spike_concentration(comp, ms_col, cfg)) * 100)
                c.number_format = '0'
            sty(c, S['rec'])

        ws.cell(row=row, column=mid2, value=None)

        # 统计公式 (基于回收率%列)
        rr = f'{rec_cl}{row}:{rec_cr}{row}'
        c_avg = ws.cell(row=row, column=stat_col)
        c_avg.value = f'=ROUND(AVERAGE({rr}),0)'; c_avg.number_format = '0'
        sty(c_avg, S['data'])

        c_sd = ws.cell(row=row, column=sd_col)
        c_sd.value = f'=ROUND(STDEV({rr}),0)'; c_sd.number_format = '0'
        sty(c_sd, S['data'])

        c_se = ws.cell(row=row, column=se_col)
        c_se.value = f'=ROUND({get_column_letter(sd_col)}{row}/SQRT(COUNT({rr})),0)'; c_se.number_format = '0'
        sty(c_se, S['data'])

        row += 1

    # --- SS 替代物回收率部分（独立于上方主列表） ---
    row += 1
    ws.cell(row=row, column=1, value='SS recoveries, %  替代物回收率')
    sty(ws.cell(row=row,column=1), S['yellL'])
    for c in range(2, last_col+1):
        sty(ws.cell(row=row,column=c), S['yell'])
    row += 1

    selected_ss = cfg.get('ss_compounds', SS_COMPS)
    for ss in selected_ss:
        ws.cell(row=row, column=1, value=ss); sty(ws.cell(row=row,column=1), S['cmpd'])
        # MS数据列：照抄原始浓度（不除以任何值）
        for i, (_, cl, _) in enumerate(ms_cols):
            v = safe_float(raw_data.get(ss, {}).get(cl))
            c = ws.cell(row=row, column=ms_start+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])
        ws.cell(row=row, column=mid1, value=None)
        ws.cell(row=row, column=rec_lbl, value=None)
        # 回收率列：SS实测浓度 ÷ SS理论加标浓度 × 100%
        for i, ms_col in enumerate(ms_cols):
            _, cl, _ = ms_col
            v = safe_float(raw_data.get(ss, {}).get(cl))
            c = ws.cell(row=row, column=rec_start+i)
            if v is not None:
                this_ss_spike = resolve_ss_spike(ss, cfg, ms_col)
                if this_ss_spike is None or this_ss_spike <= 0:
                    raise ValueError(f'{ss}: missing positive SS spike concentration for {ms_col[2]}.')
                c.value = round_int((v / this_ss_spike) * 100)
                c.number_format = '0'
            sty(c, S['rec'])
        for c in range(mid2, last_col+1):
            sty(ws.cell(row=row,column=c), S['data'])
        rr = f'{rec_cl}{row}:{rec_cr}{row}'
        c_avg = ws.cell(row=row, column=stat_col)
        c_avg.value = f'=ROUND(AVERAGE({rr}),0)'; c_avg.number_format = '0'
        sty(c_avg, S['data'])
        c_sd = ws.cell(row=row, column=sd_col)
        c_sd.value = f'=ROUND(STDEV({rr}),0)'; c_sd.number_format = '0'
        sty(c_sd, S['data'])
        c_se = ws.cell(row=row, column=se_col)
        c_se.value = f'=ROUND({get_column_letter(sd_col)}{row}/SQRT(COUNT({rr})),0)'; c_se.number_format = '0'
        sty(c_se, S['data'])
        row += 1

    # 说明
    row += 1
    note = rec_lbl + n_ms
    ws.cell(row=row, column=note, value='此表格计算方法：回收率，用测得浓度除以加标浓度')
    sty(ws.cell(row=row,column=note), S['yell'])
    row += 1
    ws.cell(row=row, column=note, value='每一个 matrix spike 列使用该列设置的加标浓度；例如 MS1=10 ppb、MS2=20 ppb 时，分别以测得浓度÷10 和测得浓度÷20 计算回收率。SS 仍只除以其自身理论加标浓度。')
    sty(ws.cell(row=row,column=note), S['yell'])

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0

    return ws


# ============================================================
# Sheet 2: 空白基质检出限
# ============================================================
def build_sheet2(wb, raw_data, blank_cols, S, cfg):
    ws = wb.create_sheet('Blanks_MDL 空白基质检出限')
    n_b = len(blank_cols)
    cf = cfg['conversion_factor']
    unit = cfg['output_unit']

    # 换算因子存放在 A1
    ws.cell(row=1, column=1, value=cf)
    ws.cell(row=1, column=1).number_format = '0.000000'

    # Row 1 说明
    ws.cell(row=1, column=9, value='仅为检出率>50% ND取代')
    sty(ws.cell(row=1,column=9), S['yell'])

    # Row 2 表头
    ws.cell(row=2, column=1, value='化合物方法'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_b):
        ws.cell(row=2, column=2+i, value=f'blank_{i+1}'); sty(ws.cell(row=2,column=2+i), S['hdr'])

    blank_end = 2 + n_b         # 空列
    avg_c = blank_end + 1       # I: average
    mdl_c = blank_end + 2       # J: MDL
    half_c = blank_end + 3      # K: 1/2 MDL
    last_c = half_c

    ws.cell(row=2, column=avg_c, value='procedural blank average'); sty(ws.cell(row=2,column=avg_c), S['hdr'])
    ws.cell(row=2, column=mdl_c, value='MDL'); sty(ws.cell(row=2,column=mdl_c), S['hdr'])
    ws.cell(row=2, column=half_c, value='1/2 MDL'); sty(ws.cell(row=2,column=half_c), S['hdr'])

    # Row 3 单位
    ws.cell(row=3, column=1, value='分组'); sty(ws.cell(row=3,column=1), S['hdr'])
    for i in range(n_b):
        ws.cell(row=3, column=2+i, value='ppb'); sty(ws.cell(row=3,column=2+i), S['hdr'])
    ws.cell(row=3, column=blank_end, value=None)
    ws.cell(row=3, column=avg_c, value='bottle ppb'); sty(ws.cell(row=3,column=avg_c), S['hdr'])
    ws.cell(row=3, column=mdl_c, value='bottle ppb'); sty(ws.cell(row=3,column=mdl_c), S['hdr'])
    ws.cell(row=3, column=half_c, value=unit); sty(ws.cell(row=3,column=half_c), S['hdr'])

    # Row 4 样本标签
    row = 4
    ws.cell(row=row, column=1, value=cfg['sample_type']); sty(ws.cell(row=row,column=1), S['cmpd'])
    for c in range(2, last_c+1): sty(ws.cell(row=row,column=c), S['data'])
    row += 1

    avg_l = get_column_letter(avg_c)
    blank_l = get_column_letter(2)
    blank_r = get_column_letter(2 + n_b - 1)

    first_data_row = row
    for comp in cfg.get('all_compounds', ALL_COMPS):
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])

        # Blank 值
        for i, (_, cl, _) in enumerate(blank_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=2+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])

        ws.cell(row=row, column=blank_end, value=None)

        # I: Average (公式)
        br = f'{blank_l}{row}:{blank_r}{row}'
        c_avg = ws.cell(row=row, column=avg_c)
        c_avg.value = f'=AVERAGE({br})'
        c_avg.number_format = '0.000000'
        sty(c_avg, S['data'])

        # J: MDL = 3*STDEVA(blanks) — 公式
        c_mdl = ws.cell(row=row, column=mdl_c)
        c_mdl.value = mdl_formula(comp, br, cfg)
        c_mdl.number_format = '0.000000'
        sty(c_mdl, S['data'])

        # K: 1/2 MDL = ROUND(MDL/2 * $A$1, 6) — 公式，引用A1换算因子
        mdl_l = get_column_letter(mdl_c)
        c_half = ws.cell(row=row, column=half_c)
        c_half.value = f'=ROUND({mdl_l}{row}/2*$A$1,6)'
        c_half.number_format = '0.000000'
        sty(c_half, S['data'])

        row += 1

    # SS 替代物部分
    row += 1
    ws.cell(row=row, column=1, value='SS recoveries, %替代物回收率')
    sty(ws.cell(row=row,column=1), S['yellL'])
    for c in range(2, last_c+1):
        sty(ws.cell(row=row,column=c), S['yell'])
    row += 1

    selected_ss = cfg.get('ss_compounds', SS_COMPS)
    for ss in selected_ss:
        ws.cell(row=row, column=1, value=ss); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i, (_, cl, _) in enumerate(blank_cols):
            v = safe_float(raw_data.get(ss, {}).get(cl))
            c = ws.cell(row=row, column=2+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])
        for c in range(blank_end, last_c+1):
            sty(ws.cell(row=row,column=c), S['data'])
        row += 1

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0

    info = {
        'avg_c': avg_c, 'mdl_c': mdl_c, 'half_c': half_c,
        'avg_l': avg_l, 'mdl_l': get_column_letter(mdl_c), 'half_l': get_column_letter(half_c),
        'first_row': first_data_row,
        'row_map': {comp: first_data_row + index for index, comp in enumerate(cfg.get('all_compounds', ALL_COMPS))},
    }
    return ws, info


# ============================================================
# Sheet 3: 瓶内实测浓度
# ============================================================
def build_sheet3(wb, raw_data, sample_cols, S, cfg):
    ws = wb.create_sheet('Conc. in bottle 瓶内实测浓度')
    n_s = len(sample_cols)

    ws.cell(row=1, column=1, value='化合物方法'); sty(ws.cell(row=1,column=1), S['hdr'])
    for i, (_, _, hdr) in enumerate(sample_cols):
        ws.cell(row=1, column=2+i, value=hdr.replace('0527 Urine-',''))
        sty(ws.cell(row=1,column=2+i), S['hdr'])

    ws.cell(row=2, column=1, value='分组'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_s):
        ws.cell(row=2, column=2+i, value=cfg['masshunter_unit']); sty(ws.cell(row=2,column=2+i), S['hdr'])

    row = 3
    ws.cell(row=row, column=1, value=cfg['sample_type']); sty(ws.cell(row=row,column=1), S['cmpd'])
    for i in range(n_s): sty(ws.cell(row=row,column=2+i), S['data'])
    row += 1
    first_row = row

    for comp in cfg.get('target_compounds', TARGET_COMPS):
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i, (_, cl, _) in enumerate(sample_cols):
            v = safe_float(raw_data.get(comp, {}).get(cl))
            c = ws.cell(row=row, column=2+i)
            if v is not None:
                c.value = round6(v); c.number_format = '0.000000'
            sty(c, S['data'])
        row += 1

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0
    return ws, first_row


# ============================================================
# Sheet 4: 最终计算浓度 (v2: 删除D/E列, 合并统计表头, A38标注)
# ============================================================
def build_sheet4(wb, raw_data, sample_cols, blank_info, s3_first, S, cfg):
    ws = wb.create_sheet('Final. conc 最终计算浓度')
    n_s = len(sample_cols)
    cf = cfg['conversion_factor']
    unit = cfg['output_unit']

    # 列布局 (删除D/E后):
    # A:化合物 B:BLANK avg C:MDL D:DF E:MEAN F:Geomean G:MEDIAN H:MIN I:MAX
    # J:5TH K:25TH L:75TH M:95TH N:空 O:说明 P~CX:样品数据(88列)
    sample_start = 16  # P列
    last_sample = sample_start + n_s - 1

    blanks_name = 'Blanks_MDL 空白基质检出限'
    bottle_name = 'Conc. in bottle 瓶内实测浓度'
    al = blank_info['avg_l']
    ml = blank_info['mdl_l']
    hl = blank_info['half_l']

    # Row 1
    ws.cell(row=1, column=1, value='化合物方法'); sty(ws.cell(row=1,column=1), S['hdr'])
    ws.cell(row=1, column=2, value='bottle'); sty(ws.cell(row=1,column=2), S['hdr'])
    ws.cell(row=1, column=3, value='bottle'); sty(ws.cell(row=1,column=3), S['hdr'])
    for c in range(4, 14): ws.cell(row=1, column=c, value=None)
    ws.cell(row=1, column=14, value=None)
    ws.cell(row=1, column=15, value=None)

    # 样品列头
    for i, (_, _, hdr) in enumerate(sample_cols):
        ws.cell(row=1, column=sample_start+i, value=hdr.replace('0527 Urine-',''))
        sty(ws.cell(row=1,column=sample_start+i), S['hdr'])

    # Row 2
    ws.cell(row=2, column=1, value='分组'); sty(ws.cell(row=2,column=1), S['hdr'])
    ws.cell(row=2, column=2, value='BLANK average'); sty(ws.cell(row=2,column=2), S['hdr'])
    ws.cell(row=2, column=3, value='MDL'); sty(ws.cell(row=2,column=3), S['hdr'])

    stat_names = ['DF 检出率','MEAN','Geometric Mean','MEDIAN','MIN','MAX',
                  '5TH','25TH','75TH','95TH']
    for i, nm in enumerate(stat_names):
        ws.cell(row=2, column=4+i, value=nm)
        sty(ws.cell(row=2,column=4+i), S['stat'])

    ws.cell(row=2, column=15, value='称样量 g, 请根据你的实际称样量填写')
    sty(ws.cell(row=2,column=15), S['yell'])

    for i in range(n_s):
        ws.cell(row=2, column=sample_start+i, value=cfg.get('sample_volume_ml', 2))
        sty(ws.cell(row=2,column=sample_start+i), S['data'])

    # 统计列 D-M (cols 4-13): Row1+Row2 合并
    stat_end = 4 + len(stat_names) - 1  # col 13
    for col in range(4, stat_end + 1):
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        # 合并后，合并区域的左上角单元格保留值，其他被清除
        # 把 row2 的值移到 row1
        cl = get_column_letter(col)
        merged_cell = ws.cell(row=1, column=col)
        merged_cell.value = stat_names[col - 4]
        sty(merged_cell, S['stat'])

    # Row 3 单位
    ws.cell(row=3, column=1, value=cfg['sample_type']); sty(ws.cell(row=3,column=1), S['cmpd'])
    ws.cell(row=3, column=2, value='ng/mL'); sty(ws.cell(row=3,column=2), S['hdr'])
    ws.cell(row=3, column=3, value='ng/mL'); sty(ws.cell(row=3,column=3), S['hdr'])
    ws.cell(row=3, column=4, value='%'); sty(ws.cell(row=3,column=4), S['stat'])
    for c in range(5, 14):
        ws.cell(row=3, column=c, value=unit); sty(ws.cell(row=3,column=c), S['stat'])

    # A38 标注, B38 存储换算因子
    ws.cell(row=38, column=1, value='换算因子')
    sty(ws.cell(row=38,column=1), S['bold'])
    ws.cell(row=38, column=2, value=cf)
    ws.cell(row=38,column=2).number_format = '0.000000'
    sty(ws.cell(row=38,column=2), S['data'])

    # 行号映射
    s2_first = blank_info['first_row']
    all_compounds = cfg.get('all_compounds', ALL_COMPS)
    target_compounds = cfg.get('target_compounds', TARGET_COMPS)
    comp2row_s2 = {c: s2_first + i for i, c in enumerate(all_compounds)}
    comp2row_s3 = {c: s3_first + i for i, c in enumerate(target_compounds)}

    row = 4
    for comp in target_compounds:
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        s2r = comp2row_s2.get(comp)
        s3r = comp2row_s3.get(comp)

        # B: BLANK average
        cb = ws.cell(row=row, column=2)
        if s2r: cb.value = f"='{blanks_name}'!{al}{s2r}"; cb.number_format = '0.000000'
        sty(cb, S['data'])

        # C: MDL in the report/sample unit (bottle MDL is converted once).
        cc = ws.cell(row=row, column=3)
        if s2r: cc.value = f"='{blanks_name}'!{ml}{s2r}*$B$38"; cc.number_format = '0.000000'
        sty(cc, S['data'])

        # Final concentration range (P~CX) and hidden detection-status range.
        # Status is 1 only when the original bottle value exceeds the blank average.
        sl = get_column_letter(sample_start)
        el = get_column_letter(last_sample)
        sr = f'{sl}{row}:{el}{row}'
        detection_start = last_sample + 1
        detection_end = detection_start + n_s - 1
        dsl = get_column_letter(detection_start)
        del_ = get_column_letter(detection_end)
        dsr = f'{dsl}{row}:{del_}{row}'

        # D: DF 检出率
        cd = ws.cell(row=row, column=4)
        cd.value = f'=IFERROR(COUNTIF({dsr},">0")/COUNT({dsr}),0)'
        cd.number_format = '0.00%'
        sty(cd, S['stat'])

        # E~I: MEAN, GEOMEAN, MEDIAN, MIN, MAX
        funcs = {5:'AVERAGE', 6:'GEOMEAN', 7:'MEDIAN', 8:'MIN', 9:'MAX'}
        for col, func in funcs.items():
            c = ws.cell(row=row, column=col)
            c.value = f'=IF(COUNT({sr})>0,{func}({sr}),"NC")'
            c.number_format = '0.000000'
            sty(c, S['stat'])

        # J~M: Percentiles
        for col, pct in [(10,0.05),(11,0.25),(12,0.75),(13,0.95)]:
            c = ws.cell(row=row, column=col)
            c.value = f'=IF(COUNT({sr})>0,PERCENTILE({sr},{pct}),"NC")'
            c.number_format = '0.000000'
            sty(c, S['stat'])

        # 样品数据列 P~CX
        for i in range(n_s):
            col = sample_start + i
            detection_col = detection_start + i
            s3_cl = get_column_letter(2 + i)

            if s2r and s3r:
                ws.cell(row=row, column=detection_col).value = (
                    f"=IF('{bottle_name}'!{s3_cl}{s3r}=\"\",\"\","
                    f"IF('{bottle_name}'!{s3_cl}{s3r}>'{blanks_name}'!{al}{s2r},1,0))"
                )
                ws.cell(row=row, column=detection_col).number_format = '0'
                formula = (
                    f"=IF('{bottle_name}'!{s3_cl}{s3r}=\"\",\"\","
                    f"IF('{bottle_name}'!{s3_cl}{s3r}>'{blanks_name}'!{al}{s2r},"
                    f"('{bottle_name}'!{s3_cl}{s3r}-'{blanks_name}'!{al}{s2r})*$B$38,"
                    f"'{blanks_name}'!{hl}{s2r}))"
                )
                ws.cell(row=row, column=col).value = formula
                ws.cell(row=row, column=col).number_format = '0.000000'
            sty(ws.cell(row=row, column=col), S['data'])
            ws.column_dimensions[get_column_letter(detection_col)].hidden = True
        row += 1

    ws.row_dimensions[1].height = 19.5
    ws.row_dimensions[2].height = 17.25
    ws.column_dimensions['A'].width = 30.0
    return ws


# ============================================================
# Sheet 5: 统计计算数据
# ============================================================
def build_sheet5(wb, sample_cols, S, cfg=None):
    ws = wb.create_sheet('统计计算结果')
    n_s = len(sample_cols)
    fn = 'Final. conc 最终计算浓度'
    ss = 16  # sample start col P

    ws.cell(row=1, column=1, value='统计数据源（原始空单元格不参与统计；未检出有效样品按该化合物 1/2 MDL 替代）')

    ws.cell(row=2, column=1, value='样品名称'); sty(ws.cell(row=2,column=1), S['hdr'])
    for i in range(n_s):
        ws.cell(row=2, column=2+i, value=f"='{fn}'!{get_column_letter(ss+i)}$1")
        sty(ws.cell(row=2,column=2+i), S['hdr'])

    ws.cell(row=3, column=1, value='称样量(g)'); sty(ws.cell(row=3,column=1), S['hdr'])
    for i in range(n_s):
        ws.cell(row=3, column=2+i, value=f"='{fn}'!{get_column_letter(ss+i)}$2")
        sty(ws.cell(row=3,column=2+i), S['data'])

    row = 4
    cfg = cfg or {}
    for idx, comp in enumerate(cfg.get('target_compounds', TARGET_COMPS)):
        fr = 4 + idx
        ws.cell(row=row, column=1, value=comp); sty(ws.cell(row=row,column=1), S['cmpd'])
        for i in range(n_s):
            ws.cell(row=row, column=2+i, value=f"='{fn}'!{get_column_letter(ss+i)}{fr}")
            sty(ws.cell(row=row,column=2+i), S['data'])
        row += 1

    ws.column_dimensions['A'].width = 30.0
    return ws


def build_summary_sheet(wb, final_sheet, blank_info, sample_cols, S, cfg):
    """Create a compact, formula-driven descriptive-statistics sheet."""
    title = '\u63cf\u8ff0\u6027\u7edf\u8ba1'
    ws = wb.create_sheet(title)
    final_name = final_sheet.title
    blanks_name = 'Blanks_MDL \u7a7a\u767d\u57fa\u8d28\u68c0\u51fa\u9650'
    sample_start = 16
    last_sample = sample_start + len(sample_cols) - 1
    first_letter = get_column_letter(sample_start)
    last_letter = get_column_letter(last_sample)
    target_compounds = cfg.get('target_compounds', TARGET_COMPS)
    mql_factor = safe_float(cfg.get('mql_factor', 3.333333333))
    if mql_factor is None or mql_factor <= 0:
        raise ValueError('mql_factor must be positive.')

    headers = ['\u540d\u79f0', '\u94fe\u957f', 'DF (%)', 'Median (Q1-Q3)', 'MDL', 'MQL']
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.cell(row=1, column=1, value='\u63cf\u8ff0\u6027\u7edf\u8ba1\uff08\u4fdd\u75593\u4f4d\u6709\u6548\u6570\u5b57\uff09')
    sty(ws.cell(row=1, column=1), S['hdr'])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws.cell(row=2, column=1, value=f'MQL = MDL × {mql_factor:g}；DF 始终展示真实检出率；Median(Q1-Q3) 基于所有最终浓度且不受 DF 门槛限制；MDL 引用 Blanks_MDL；公式保留。')
    sty(ws.cell(row=2, column=1), S['yell'])
    for col, header in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=header)
        sty(ws.cell(row=3, column=col), S['hdr'])

    mdl_letter = blank_info['mdl_l']
    for index, compound in enumerate(target_compounds):
        row = 4 + index
        final_row = 4 + index
        sample_range = f"'{final_name}'!{first_letter}{final_row}:{last_letter}{final_row}"
        ws.cell(row=row, column=1, value=compound)
        ws.cell(row=row, column=2, value=extract_chain_length(compound) or 'NA')
        # DF is displayed directly as a percentage number (for example 33.3),
        # consistent with the online preview and the "DF (%)" heading.
        ws.cell(row=row, column=3, value=f'={significant_digits_formula(f"\'{final_name}\'!D{final_row}*100")}')
        # The result sheet summarizes all final concentrations independently
        # from DF. Final values include approved 1/2 MDL substitutions.
        median_formula = significant_digits_formula(f'MEDIAN({sample_range})')
        q1_formula = significant_digits_formula(f'PERCENTILE({sample_range},0.25)')
        q3_formula = significant_digits_formula(f'PERCENTILE({sample_range},0.75)')
        ws.cell(row=row, column=4, value=(
            f'=IF(COUNT({sample_range})>0,{median_formula}&" ("&{q1_formula}&"-"&{q3_formula}&")","NC")'
        ))
        blank_row = blank_info['row_map'].get(compound)
        if blank_row is None:
            raise ValueError(f'{compound}: missing MDL row in blank sheet.')
        mdl_ref = f"'{blanks_name}'!{mdl_letter}{blank_row}"
        # Use the same report-unit MDL as Final. conc; blank-zero analytes use
        # the explicit calibration/SN formula and are converted exactly once.
        report_mdl_formula = mdl_report_formula(compound, mdl_ref, cfg)
        ws.cell(row=row, column=5, value=f'={significant_digits_formula(report_mdl_formula[1:])}')
        ws.cell(row=row, column=6, value=f'={significant_digits_formula(f"({report_mdl_formula[1:]})*{mql_factor}")}')
        for col in range(1, 7):
            sty(ws.cell(row=row, column=col), S['cmpd'] if col in (1, 2) else S['stat'])
        # Each formula has already rounded the result to three significant
        # digits. General avoids a fixed six-decimal display when users paste
        # the compact descriptive statistics into manuscripts.
        for col in (3, 5, 6):
            ws.cell(row=row, column=col).number_format = 'General'

    for col, width in enumerate([30, 12, 12, 28, 14, 14], 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f'A3:F{3 + len(target_compounds)}'
    return ws


# ============================================================
# 说明书
# ============================================================
def build_info_sheet(wb, S, cfg):
    ws = wb.create_sheet('计算说明')
    rows = [
        ['LC-MS/MS 数据处理说明 v2'],
        ['区域','公式/规则','来源','说明'],
        ['Sheet1: MS数据','从原始数据含"MS"列提取','原始数据',f'单位:{cfg["masshunter_unit"]}'],
        ['Sheet1: 回收率%','MS值÷对应 MS 列的加标浓度×100,四舍五入取整','Sheet1 MS列','单位:%；每一个 MS 列可设置不同浓度'],
        ['Sheet1: avg/SD/SE','AVERAGE/STDEV/SD/SQRT(COUNT) 公式','Sheet1 回收率%列','基于回收率百分比值'],
        ['Sheet2: Blank avg','=AVERAGE(空白) 公式','Sheet2 B~G列',f'单位:bottle {cfg["masshunter_unit"]}'],
        ['Sheet2: MDL','=3*STDEVA(空白) 公式','Sheet2 J列',f'单位:bottle {cfg["masshunter_unit"]}'],
        ['Sheet2: 1/2 MDL','=ROUND(MDL/2*$A$1,6) 公式','Sheet2 K列',f'单位:{cfg["output_unit"]}，$A$1=换算因子={cfg["conversion_factor"]}'],
        ['Sheet3','原始数据实际样品列直接迁移','原始数据','空白单元格保持空白'],
        ['Sheet4 B/C','引用Sheet2 I/J列','Sheet2',''],
        ['Sheet4 R~DA','IF(瓶内值>空白平均,(瓶内值-空白平均)×$B$38, 1/2MDL)','Sheet2/Sheet3','$B$38=换算因子；1/2 MDL 已在 Sheet2 转换'],
        ['Sheet4 DF','COUNTIF(隐藏检出状态,">0")/COUNT(隐藏检出状态)','Sheet4 样品列','仅原始瓶内值>空白平均值计为真实检出；1/2 MDL 不计入 DF'],
        ['Sheet4 统计','IF(COUNT(最终浓度)>0,统计函数,"NC")','Sheet4','空值自动忽略；描述性统计页不使用 DF>50% 门槛'],
        ['换算因子位置','Sheet2 $A$1 + Sheet4 $B$38','','两处均可独立修改'],
        ['本次参数',f'样本:{cfg["sample_type"]} 取样:{cfg["sample_volume_ml"]}mL 定容:{cfg["final_volume_ml"]}mL 换算因子:{cfg["conversion_factor"]}','',''],
    ]
    is_corrected = bool(cfg.get('is_corrected', False))
    ms_headers = cfg.get('matrix_spike_headers') or []
    matrix_cfg = cfg.get('matrix_spike_concentrations') or {}
    configured_is = cfg.get('is_spike_concentrations') or {}
    for name in cfg.get('is_compounds', []):
        per_ms = matrix_cfg.get(name)
        if isinstance(per_ms, dict) and ms_headers:
            values = '; '.join(
                f'{header}: {safe_float(per_ms.get(header)):g} ppb'
                for header in ms_headers
                if safe_float(per_ms.get(header)) is not None and safe_float(per_ms.get(header)) > 0
            )
            if values:
                rows.append([
                    'IS addition record', name, values,
                    'Recorded only; IS correction applied: ' + ('yes' if is_corrected else 'no') + '. Does not change concentration formulas.',
                ])
        else:
            value = safe_float(configured_is.get(name))
            if value is not None and value > 0:
                rows.append([
                    'IS addition record', name, f'{value:g} ppb',
                    'Recorded only; IS correction applied: ' + ('yes' if is_corrected else 'no') + '. Does not change concentration formulas.',
                ])
    for name, configured in matrix_cfg.items():
        if isinstance(configured, dict):
            values = '; '.join(
                f'{header}: {safe_float(configured.get(header)):g} ppb'
                for header in ms_headers
                if safe_float(configured.get(header)) is not None and safe_float(configured.get(header)) > 0
            )
            if values:
                rows.append(['Matrix spike concentration', name, values, 'Used for this compound in the corresponding MS columns; IS values are record-only.'])
        else:
            value = safe_float(configured)
            if value is not None and value > 0:
                rows.append(['Matrix spike concentration', name, f'{value:g} ppb', 'Used only for this matrix-spike recovery column.'])
    for r, rd in enumerate(rows, 1):
        for c, val in enumerate(rd, 1):
            cell = ws.cell(row=r, column=c, value=val)
            sty(cell, S['hdr'] if r <= 2 else S['data'])
    for c, w in zip([1,2,3,4], [35,65,20,60]):
        ws.column_dimensions[get_column_letter(c)].width = w
    return ws


def export_csv_bytes(raw_data, blank_cols, sample_cols, cfg):
    """Create a formula-free CSV containing summary and final concentrations."""
    preview_cfg = dict(cfg)
    preview_cfg['_raw_data'] = raw_data
    rows = [['Name', 'Chain length', 'DF (%)', 'Median (Q1-Q3)', 'MDL', 'MQL']]
    for item in compute_preview_summary(raw_data, blank_cols, sample_cols, preview_cfg):
        rows.append([item.get('名称') or item.get('鍚嶇О'), item.get('链长') or item.get('閾鹃暱'), item.get('DF (%)'), item.get('Median (Q1-Q3)'), item.get('MDL'), item.get('MQL')])
    sample_headers = [header for _, _, header in sample_cols]
    rows.append([])
    rows.append(['Final concentration summary'] + sample_headers)
    for item in compute_preview_final_table(raw_data, blank_cols, sample_cols, preview_cfg):
        name = item.get('名称') or item.get('鍚嶇О')
        rows.append([name] + [item.get(header) for header in sample_headers])
    rows.append([])
    rows.append(['IS addition record', 'Addition concentration (ppb)', 'IS correction applied', 'Note'])
    configured_is = cfg.get('is_spike_concentrations') or {}
    for name in cfg.get('is_compounds', []):
        value = safe_float(configured_is.get(name))
        if value is not None and value > 0:
            rows.append([name, value, 'yes' if cfg.get('is_corrected', False) else 'no', 'Recorded only; does not change concentration formulas.'])
    text = io.StringIO()
    csv.writer(text, lineterminator='\n').writerows(rows)
    return text.getvalue().encode('utf-8-sig')


# ============================================================
# 主处理函数 (可被 Streamlit 调用)
# ============================================================
def process(config=None, return_bytes=False):
    """
    执行完整的数据处理流程。

    Args:
        config: 配置字典，None 则使用默认 CONFIG
        return_bytes: True 返回 bytes (供 Streamlit), False 保存到文件

    Returns:
        如果 return_bytes=True，返回 (bytes, filename)
        否则返回 output_filepath
    """
    cfg = {**CONFIG, **(config or {})}
    input_src = cfg.get('input_file', '')
    input_bytes = cfg.get('input_bytes', None)
    src = input_bytes if input_bytes else input_src
    if not src:
        raise ValueError('No input file or data provided')
    raw_data, blanks, mss, samps, detected_target, detected_is, detected_ss, detected_all = read_raw(src)
    cfg['matrix_spike_headers'] = [header for _, _, header in mss]
    target, is_c, ss_c, all_c = configured_compound_lists(cfg, detected_all, detected_is, detected_ss)
    cfg['target_compounds'] = target
    cfg['is_compounds'] = is_c
    cfg['ss_compounds'] = ss_c
    cfg['all_compounds'] = all_c
    layout_report = validate_input_layout(blanks, mss, samps, target, is_c, ss_c)
    if not layout_report['ready']:
        raise ValueError('Input layout is not processable: ' + '; '.join(layout_report['errors']))
    validate_blank_zero_configuration(raw_data, blanks, target, cfg)

    # 验证
    missing = [c for c in all_c if c not in raw_data]
    if missing:
        print(f"WARNING: {len(missing)} compounds not found: {missing}")

    S = make_styles()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    print("[1/6] Matrix spike...")
    build_sheet1(wb, raw_data, mss, S, cfg)

    print("[2/6] Blanks_MDL...")
    _, binfo = build_sheet2(wb, raw_data, blanks, S, cfg)

    print("[3/6] Conc. in bottle...")
    _, s3_first = build_sheet3(wb, raw_data, samps, S, cfg)

    print("[4/6] Final conc...")
    final_sheet = build_sheet4(wb, raw_data, samps, binfo, s3_first, S, cfg)

    print("[5/6] Stats helper...")
    build_sheet5(wb, samps, S, cfg)

    print("[6/7] Descriptive summary...")
    build_summary_sheet(wb, final_sheet, binfo, samps, S, cfg)

    print("[7/7] Info sheet...")
    build_info_sheet(wb, S, cfg)

    if return_bytes:
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        xlsx_bytes = output.getvalue()
        wb.close()
        if cfg.get('output_format', 'xlsx').lower() == 'csv':
            return export_csv_bytes(raw_data, blanks, samps, cfg), cfg.get('output_file', 'processed_data.csv').replace('.xlsx', '.csv')
        return xlsx_bytes, cfg.get('output_file', 'processed_data.xlsx')
    else:
        out = cfg['output_file']
        print(f"Saving to: {out}")
        wb.save(out)
        wb.close()
        print("Done!")
        return out


# ============================================================
# 入口
# ============================================================
if __name__ == '__main__':
    process()
