#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LC-MS/MS 数据处理智能体 — Streamlit 可视化界面
"""

import streamlit as st
import pandas as pd
import tempfile
import os
import io
import importlib
import process_lcms_data as processor

# Streamlit Cloud can hot-reload this entry file while retaining the previous
# helper module in the worker process. Reload before binding names so a new UI
# and its matching processor helpers are always deployed atomically.
processor = importlib.reload(processor)
process = processor.process
read_raw = processor.read_raw
classify_compounds = processor.classify_compounds
resolve_roles = processor.resolve_roles
compute_preview_summary = processor.compute_preview_summary
compute_preview_final_table = processor.compute_preview_final_table
build_blank_mdl_evidence = processor.build_blank_mdl_evidence
parse_custom_ss_entries = processor.parse_custom_ss_entries
parse_compound_name_entries = processor.parse_compound_name_entries
missing_matrix_spike_entries = processor.missing_matrix_spike_entries
compound_classification_rows = processor.compound_classification_rows
compound_metadata_for = processor.compound_metadata_for

APP_VERSION = '2026.08.17-ss-input-example-v7'

try:
    validate_input_layout = processor.validate_input_layout
except AttributeError:
    # Keep the app bootable while Streamlit Cloud refreshes an older module cache.
    def validate_input_layout(blanks, mss, samps, target_compounds, is_compounds, ss_compounds):
        errors = []
        warnings = []
        if not blanks:
            errors.append('未识别到 BLANK 列；请检查列名是否包含 BLANK。')
        if not samps:
            errors.append('未识别到 sample 列；请检查样品列名称。')
        if not target_compounds:
            errors.append('未识别到目标化合物行；请检查化合物名称列。')
        if not mss:
            warnings.append('未识别到 MS/基质加标列；基质加标回收率将无法计算。')
        if len(blanks) < 2:
            warnings.append('BLANK 列少于 2 个；请确认 MDL 计算所需的空白数量。')
        if not is_compounds:
            warnings.append('未识别到 IS 内标；请确认是否使用内标校正。')
        if not ss_compounds:
            warnings.append('未识别到 SS 替代物；SS 回收率不会自动生成。')
        return {
            'ready': not errors,
            'errors': errors,
            'warnings': warnings,
            'summary': f'{len(blanks)} BLANK + {len(mss)} MS + {len(samps)} sample + '
                       f'{len(is_compounds)} IS + {len(ss_compounds)} SS + {len(target_compounds)} 个目标物',
        }


def read_preview_table(file_bytes):
    """Read XLSX/XLS or a delimited MassHunter export for display."""
    if file_bytes.startswith(b'PK'):
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)
    if file_bytes.startswith(b'\xd0\xcf\x11\xe0'):
        return pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd')
    for encoding in ('utf-8-sig', 'utf-16', 'utf-8', 'gb18030', 'big5'):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), header=None, encoding=encoding, sep=None, engine='python')
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    raise ValueError('无法识别 CSV 编码或分隔符，请保存为 UTF-8、UTF-16、GB18030 或制表符分隔的 CSV。')

# ============================================================
# 中英文对照字典
# ============================================================
T = {
    'page_title':           {'zh': 'LC-MS/MS 数据处理智能体',                    'en': 'LC-MS/MS Data Processing Agent'},
    'page_subtitle':        {'zh': '上传 MassHunter 原始数据 → 一键生成完整分析表格 → 下载', 'en': 'Upload MassHunter raw data → Auto-generate standard tables → Download'},
    'sidebar_header':       {'zh': '📋 实验参数',                                'en': '📋 Experiment Parameters'},
    'sample_type':          {'zh': '样本类型',                                   'en': 'Sample Type'},
    'sample_type_opts':     {'zh': ['尿液','灰尘','土壤','水','其他'],             'en': ['Urine','Dust','Soil','Water','Other']},
    'sample_vol':           {'zh': '取样量 (mL/g)',                              'en': 'Sample Volume (mL/g)'},
    'final_vol':            {'zh': '定容体积 (mL)',                              'en': 'Final Volume (mL)'},
    'extra_dil':            {'zh': '额外稀释倍数',                                'en': 'Extra Dilution Factor'},
    'is_correction':        {'zh': '数据是否经过内标（IS）校正？',                 'en': 'IS (Internal Standard) Correction Applied?'},
    'is_opts':              {'zh': ['✅ 是，有内标校正','❌ 否，无内标校正'],       'en': ['✅ Yes, IS Corrected','❌ No, Raw Concentration']},
    'is_help':              {'zh': '内标校正后，导出值已是原始样本浓度；无内标校正则需要手动换算',
                             'en': 'With IS correction, exported values are already original sample concentration; without IS correction, manual conversion is needed'},
    'csv_note':             {'zh': 'CSV 是单页、无公式数值报告；需要多工作表和可审计公式时请选择 XLSX。', 'en': 'CSV is a flat, formula-free report; choose XLSX for multi-sheet formulas.'},
    'cf_caption_yes':       {'zh': '💡 有内标校正，仪器已自动将进样瓶浓度换算为原始尿液浓度，换算因子 = 1',
                             'en': '💡 IS corrected: instrument has already converted vial concentration to original sample concentration. CF = 1'},
    'cf_caption_no':        {'zh': '💡 无内标校正，导出值是进样瓶浓度，换算因子用于将其换算为原始尿液浓度：定容体积 ÷ 取样量 × 稀释倍数',
                             'en': '💡 Not IS corrected: exported value is vial concentration. CF = Final Vol ÷ Sample Vol × Dilution to convert to original concentration'},
    'cf_label':             {'zh': '换算因子',                                   'en': 'Conversion Factor'},
    'cf_locked':            {'zh': '（已锁定）',                                  'en': ' (Locked)'},
    'cf_editable':          {'zh': '（可手动覆盖）',                              'en': ' (Editable)'},
    'spike_conc':           {'zh': '基质加标浓度 (ppb)',                          'en': 'Matrix Spike Conc (ppb)'},
    'ms_spike_header':      {'zh': '基质加标浓度设置（化合物 × MS列）',               'en': 'Matrix-spike concentrations (compound × MS column)'},
    'ms_spike_help':        {'zh': '系统根据上传文件的实际MS列数量自动增减输入列。逐格填写目标物和SS替代物在对应MS中的基质加标浓度（ppb）；原始表中的MS实测浓度由系统自动读取。SS回收率 = 对应MS实测浓度 ÷ 对应MS基质加标浓度 × 100%。IS不在此表填写浓度，也不计算回收率。',
                             'en': 'Input columns automatically follow the actual number of MS columns in the uploaded file. Enter the matrix-spike concentration for each target and SS in each corresponding MS; measured MS values are read from the source. SS recovery = measured MS concentration / corresponding MS matrix-spike concentration × 100%. IS has no concentration entry or recovery here.'},
    'ms_spike_example':     {'zh': '示例中的MS1、MS2、MS3仅用于说明；真实输入表会按原始文件的实际MS列数量动态生成，可能少于或多于3列。',
                             'en': 'MS1, MS2 and MS3 are examples only. The real input grid is generated dynamically and may contain fewer or more than three MS columns.'},
    'ss_is_example_header': {'zh': 'SS基质加标浓度与IS内标名称填写示例', 'en': 'SS matrix-spike concentration and IS name examples'},
    'ss_example_title': {'zh': 'SS替代物基质加标浓度示例', 'en': 'SS matrix-spike concentration example'},
    'ss_example_help': {
        'zh': '这些数值是SS自身在对应MS样品中的基质加标浓度，是SS回收率计算的分母，不是原始表中的MS实测浓度。实测浓度由系统自动读取。示例只展示3个MS；真实输入列数按上传文件动态生成。',
        'en': 'These are the SS matrix-spike concentrations used as recovery denominators, not the measured MS values in the source file. Measured values are read automatically. Three MS columns are shown only as an example; the real grid follows the uploaded file.',
    },
    'is_example_title': {'zh': 'IS内标名称输入示例', 'en': 'IS name input example'},
    'is_example_help': {
        'zh': 'IS只输入化合物名称，不输入加入浓度。系统将名称与上传文件中的化合物自动匹配并识别为IS；IS不计算回收率。是否经过IS校正仍由相应选项决定。',
        'en': 'Enter IS compound names only, without addition concentrations. The app matches them to uploaded compounds and marks them as IS. IS recovery is not calculated; the IS-corrected option remains separate.',
    },
    'ms_table_compound': {'zh': '化合物名称', 'en': 'Compound name'},
    'ms_table_role': {'zh': '类型/角色', 'en': 'Type / role'},
    'file_header':          {'zh': '📁 文件',                                    'en': '📁 File'},
    'upload_label':         {'zh': '上传原始数据（XLSX 或 CSV）',                  'en': 'Upload Raw Data (XLSX or CSV)'},
    'output_name':          {'zh': '输出文件名',                                  'en': 'Output Filename'},
    'output_default':       {'zh': '已处理数据.xlsx',                             'en': 'processed_data.xlsx'},
    'output_format':        {'zh': '输出格式',                                    'en': 'Output Format'},
    'roles_header':         {'zh': '化合物角色设置',                              'en': 'Compound Roles'},
    'roles_caption':        {'zh': '系统按名称预识别；请确认哪些为 IS、SS。未选为 IS/SS 的化合物将作为目标物。', 'en': 'Roles are auto-detected by name; confirm IS and SS. Unselected compounds remain targets.'},
    'is_select':            {'zh': 'IS 内标（可多选）',                            'en': 'IS internal standards'},
    'ss_select':            {'zh': 'SS 替代物（可多选）',                          'en': 'SS surrogates'},
    'classification_header': {'zh': '化合物智能分类（请核对并可直接修改）',           'en': 'Smart compound classification (review and edit)'},
    'classification_help': {'zh': '系统从名称自动建议类型、链长和角色；未知结构不猜测。请直接修改类别、链长和角色，所有输出表会使用此设置排序。', 'en': 'The system suggests type, chain length and role from names without guessing unknown structures. Edit these fields directly; all output sheets use this order.'},
    'classification_name': {'zh': '名称', 'en': 'Name'},
    'classification_type': {'zh': '类型', 'en': 'Type'},
    'classification_chain': {'zh': '链长', 'en': 'Chain length'},
    'classification_role': {'zh': '角色', 'en': 'Role'},
    'role_options': {'zh': ['目标物', '替代物 (SS)', '内标 (IS)'], 'en': ['Target', 'Surrogate (SS)', 'Internal standard (IS)']},
    'blank_zero_header':    {'zh': 'Blank/MDL 逐化合物设置',                      'en': 'Per-compound Blank/MDL settings'},
    'blank_workflow_header': {'zh': 'Blank/MDL 设置与计算', 'en': 'Blank/MDL settings and calculation'},
    'blank_workflow_before_upload': {
        'zh': '请先上传文件或加载 Demo。上传后，系统会自动识别每个化合物的 blank 列：blank 全为数值0时显示标曲浓度和 S/N 输入；blank 含非零值时自动计算均值、SD、动态 t 值和 MDL，并在“计算依据”中展示。',
        'en': 'Upload a file or load the Demo first. The app then evaluates blank columns for every compound: all-numeric-zero blanks require calibration concentration and S/N; non-zero blanks automatically show mean, SD, dynamic t and MDL in Calculation evidence.',
    },
    'ss_spike_grid':        {'zh': '已选 SS 的理论加标浓度（ppb）',                   'en': 'Theoretical spike concentration for selected SS (ppb)'},
    'is_spike_grid':        {'zh': '已选 IS 的加入浓度（ppb，仅记录）',               'en': 'Addition concentration for selected IS (ppb, record only)'},
    'custom_ss':            {'zh': '自定义 SS 替代物（可选）',                       'en': 'Custom SS surrogates (optional)'},
    'custom_ss_help':       {'zh': '先在此输入SS化合物名称。多个名称可使用：英文逗号“,”、中文逗号“，”、英文分号“;”、中文分号“；”、Tab或换行分隔。名称须与上传文件一致。上传后，必须在系统按实际MS列生成的表格中，逐格填写每个SS自身的基质加标浓度（ppb）；该浓度是SS回收率分母，原始MS实测浓度由系统自动读取，两者不是同一个数值。', 'en': 'First enter the SS compound names here. Separate names with an English comma, Chinese comma, English semicolon, Chinese semicolon, Tab, or a new line. Names must match the uploaded file. After upload, every SS matrix-spike concentration must be entered for every detected MS column. It is the recovery denominator; measured MS concentration is read from the source and is a different value.'},
    'custom_ss_placeholder': {'zh': 'd7-C12-BAC，d9-C10-ATMAC\nMy Surrogate',             'en': 'd7-C12-BAC, d9-C10-ATMAC\nMy Surrogate'},
    'custom_is':            {'zh': '自定义 IS 内标（可选）',                         'en': 'Custom IS internal standards (optional)'},
    'custom_is_help':       {'zh': '这里只输入IS化合物名称，不输入浓度。多个名称可使用：英文逗号“,”、中文逗号“，”、英文分号“;”、中文分号“；”、Tab或换行分隔。系统会与上传文件中的化合物自动匹配并识别为IS；IS不计算回收率。是否使用IS校正仍由上方选项决定。',
                             'en': 'Enter IS compound names only, without concentrations. Separate names with an English comma, Chinese comma, English semicolon, Chinese semicolon, Tab, or a new line. The app matches them to uploaded compounds and marks them as IS. IS recovery is not calculated; the IS-corrected option remains separate.'},
    'custom_is_placeholder': {'zh': 'IS-A；IS-B\nC13-Internal Standard',
                              'en': 'IS-A; IS-B\nC13-Internal Standard'},
    'blank_zero_help':      {'zh': '系统逐化合物判断：全部 blank 为数值0时填写标曲浓度和 S/N；blank 含非零值时自动使用均值、动态 t 值和 SD。', 'en': 'Each compound is evaluated independently: enter calibration concentration and S/N when every blank is numeric zero; non-zero blanks use mean, dynamic t and SD automatically.'},
    'calibration':          {'zh': '标曲浓度 (ppb)',                              'en': 'Calibration concentration (ppb)'},
    'sn':                   {'zh': 'S/N',                                         'en': 'S/N'},
    'mql_help':             {'zh': '默认 3.333333；请按实验室方法确认。',              'en': 'Default 3.333333; confirm with your laboratory method.'},
    'calculation_evidence': {'zh': '计算依据', 'en': 'Calculation evidence'},
    'evidence_compound': {'zh': '化合物', 'en': 'Compound'},
    'evidence_values': {'zh': '有效 blank 值', 'en': 'Valid blank values'},
    'evidence_valid': {'zh': '有效数 n', 'en': 'Valid n'},
    'evidence_nonzero': {'zh': '非零数', 'en': 'Non-zero n'},
    'evidence_status': {'zh': '系统判断', 'en': 'System decision'},
    'evidence_action': {'zh': '处理/操作', 'en': 'Treatment / action'},
    'evidence_mdl': {'zh': '瓶内MDL预览 (ppb)', 'en': 'Vial MDL preview (ppb)'},
    'evidence_blank_zero': {'zh': 'blank=0', 'en': 'blank=0'},
    'evidence_blank_nonzero': {'zh': 'blank≠0', 'en': 'blank≠0'},
    'evidence_missing': {'zh': '无有效blank', 'en': 'No valid blanks'},
    'evidence_insufficient': {'zh': '有效blank不足', 'en': 'Insufficient blanks'},
    'evidence_incomplete': {'zh': 'blank数据不完整', 'en': 'Incomplete blank data'},
    'evidence_enter_snr': {'zh': '填写标曲浓度和S/N', 'en': 'Enter calibration concentration and S/N'},
    'evidence_automatic': {'zh': '系统自动计算', 'en': 'Calculated automatically'},
    'evidence_unavailable': {'zh': '不能计算，请检查blank数据', 'en': 'Cannot calculate; check blank data'},
    'evidence_formula': {'zh': '代入公式', 'en': 'Substituted formula'},
    'evidence_df': {'zh': '自由度', 'en': 'Degrees of freedom'},
    'evidence_t': {'zh': '单侧99% t值', 'en': 'One-sided 99% t'},
    'evidence_mean': {'zh': 'blank平均值', 'en': 'Blank mean'},
    'evidence_sd': {'zh': 'blank标准差', 'en': 'Blank SD'},
    'evidence_reason_blank_zero': {'zh': 'blank=0时必须填写大于0的标曲浓度和S/N。', 'en': 'blank=0 requires a positive calibration concentration and S/N.'},
    'evidence_reason_missing': {'zh': '没有找到有效的blank测定结果。', 'en': 'No valid blank results were found.'},
    'evidence_reason_insufficient': {'zh': '至少需要2个有效blank结果才能计算标准差。', 'en': 'At least two valid blank results are required to calculate a standard deviation.'},
    'evidence_reason_incomplete': {'zh': 'blank中存在缺失单元格；只有每个blank均为数值0时才属于blank=0。', 'en': 'Blank cells are missing; blank=0 requires every blank cell to be numeric zero.'},
    'preview_stats':        {'zh': '描述性统计（在线数值预览）',                     'en': 'Descriptive statistics (numeric preview)'},
    'preview_stats_help':   {'zh': '这里显示已计算的数值；下载的 Excel 同时保留可审计公式。', 'en': 'Calculated values are shown here; downloaded Excel retains auditable formulas.'},
    'preview_final':        {'zh': '最终浓度（在线数值预览）',                       'en': 'Final concentrations (numeric preview)'},
    'tip':                  {'zh': '上传文件 → 调参数 → 点处理 → 下载结果',        'en': 'Upload → Adjust Params → Process → Download'},
    'demo_btn':             {'zh': '📥 加载 Demo 数据',                           'en': '📥 Load Demo Data'},
    'demo_help':            {'zh': '使用示例数据体验平台功能，无需上传自己的文件',    'en': 'Try the platform with example data, no upload needed'},
    'raw_preview':          {'zh': '📊 原始数据预览',                             'en': '📊 Raw Data Preview'},
    'total_cmpd':           {'zh': '化合物总数',                                  'en': 'Total Compounds'},
    'target_cmpd':          {'zh': '目标化合物',                                  'en': 'Target Compounds'},
    'is_cmpd':              {'zh': 'IS内标',                                     'en': 'IS Internal Std'},
    'ss_cmpd':              {'zh': 'SS替代物',                                    'en': 'SS Surrogates'},
    'view_classify':        {'zh': '查看化合物分类',                              'en': 'View Compound Classification'},
    'target_label':         {'zh': '目标化合物',                                  'en': 'Target Compounds'},
    'is_label':             {'zh': 'IS内标',                                     'en': 'IS Internal Standards'},
    'ss_label':             {'zh': 'SS替代物',                                    'en': 'SS Surrogates'},
    'blank_cols':           {'zh': 'BLANK列',                                    'en': 'BLANK Columns'},
    'ms_cols':              {'zh': 'MS列',                                       'en': 'MS Columns'},
    'sample_cols':          {'zh': '样品列',                                      'en': 'Sample Columns'},
    'preview_warn':         {'zh': '预览时请注意',                                 'en': 'Preview Warning'},
    'process_btn':          {'zh': '🚀 开始处理',                                 'en': '🚀 Start Processing'},
    'processing':           {'zh': '正在处理数据...',                              'en': 'Processing data...'},
    'success':              {'zh': '✅ 处理完成！',                                'en': '✅ Processing Complete!'},
    'result_preview':       {'zh': '📋 处理结果预览',                             'en': '📋 Result Preview'},
    'rows_cols':            {'zh': '行 × ',                                      'en': ' rows × '},
    'rows_cols_suffix':     {'zh': '列',                                         'en': ' columns'},
    'download_btn':         {'zh': '⬇️ 下载处理结果',                             'en': '⬇️ Download Result'},
    'error':                {'zh': '处理失败',                                    'en': 'Processing Failed'},
    'demo_loaded': {'zh': 'Demo 数据已加载：这是可直接处理的尿液 MassHunter 示例；请核对参数后点击“开始处理”。', 'en': 'Demo data loaded: this is a directly processable urine MassHunter example. Review settings, then click Start Processing.'},
    'format_ok': {'zh': '文件格式检查通过', 'en': 'Input format check passed'},
    'format_bad': {'zh': '文件格式检查未通过', 'en': 'Input format check failed'},
    'result_formula_note': {'zh': '下载的 Excel 保留可审计公式；在线预览显示按同一规则计算的数值。', 'en': 'The downloaded Excel keeps auditable formulas; the online preview shows numeric values calculated with the same rules.'},
    'custom_is_missing': {'zh': '未在上传文件中找到自定义 IS：', 'en': 'Custom IS was not found in the uploaded file: '},
    'custom_ss_missing': {'zh': '未在上传文件中找到自定义 SS：', 'en': 'Custom SS was not found in the uploaded file: '},
    'ss_spike_required': {'zh': 'SS基质加标浓度必须填写完整，不能使用系统猜测值。请填写：', 'en': 'Every SS matrix-spike concentration is required; the app will not guess values. Enter: '},
    'ss_spike_sidebar_header': {'zh': 'SS替代物基质加标浓度（必填）', 'en': 'SS matrix-spike concentrations (required)'},
    'ss_spike_sidebar_help': {'zh': '以下是实际计算输入，不是示例。系统根据上传文件识别到的MS列自动生成；请填写每个SS在每个MS中的自身基质加标浓度。', 'en': 'These are real calculation inputs, not examples. They follow the MS columns detected in the uploaded file; enter each SS compound’s own matrix-spike concentration for every MS.'},
    'ss_input_example_title': {'zh': 'SS名称和自身基质加标浓度填写示例', 'en': 'Example: SS names and their matrix-spike concentrations'},
    'ss_input_example_body': {
        'zh': '第一步：在上方名称框输入：\n`d7-C12-BAC，d9-C10-ATMAC`\n\n第二步：上传文件后，在左侧自动出现的输入框填写：\n- `d7-C12-BAC / MS1 = 4 ppb`\n- `d7-C12-BAC / MS2 = 8 ppb`\n- `d7-C12-BAC / MS3 = 12 ppb`\n- `d9-C10-ATMAC / MS1 = 2 ppb`\n- `d9-C10-ATMAC / MS2 = 2 ppb`\n- `d9-C10-ATMAC / MS3 = 4 ppb`\n\n示例数字仅用于说明填写方法，不会自动写入真实计算。实际有几个MS，左侧就生成几个浓度输入框。',
        'en': 'Step 1: enter the names above:\n`d7-C12-BAC, d9-C10-ATMAC`\n\nStep 2: after upload, fill the sidebar inputs:\n- `d7-C12-BAC / MS1 = 4 ppb`\n- `d7-C12-BAC / MS2 = 8 ppb`\n- `d7-C12-BAC / MS3 = 12 ppb`\n- `d9-C10-ATMAC / MS1 = 2 ppb`\n- `d9-C10-ATMAC / MS2 = 2 ppb`\n- `d9-C10-ATMAC / MS3 = 4 ppb`\n\nThese numbers only explain the input method and are never inserted into real calculations. The sidebar creates one input for every actual MS column.',
    },
    'role_overlap': {'zh': '同一化合物不能同时作为 IS 与 SS：', 'en': 'An analyte cannot be both IS and SS: '},
}

def t(key, lang='zh'):
    """Get translation for key in given language"""
    if key in T:
        return T[key].get(lang, T[key].get('zh', key))
    return key


def role_label(role, language):
    """Present internal roles in the language selected by the user."""
    labels = {
        'zh': {'Target': '目标物', 'SS': '替代物 (SS)', 'IS': '内标 (IS)'},
        'en': {'Target': 'Target', 'SS': 'Surrogate (SS)', 'IS': 'Internal standard (IS)'},
    }
    return labels[language].get(role, role)


def role_from_label(label):
    """Map an editable bilingual role label back to the processor role."""
    value = str(label or '').strip()
    if value in {'IS', 'Internal standard (IS)', '内标 (IS)', '内标'}:
        return 'IS'
    if value in {'SS', 'Surrogate (SS)', '替代物 (SS)', '替代物'}:
        return 'SS'
    return 'Target'

# ============================================================
# 页面设置
# ============================================================
st.set_page_config(
    page_title="LC-MS/MS 数据处理智能体 | LC-MS/MS Data Processor",
    page_icon="🔬",
    layout="wide",
)

# 语言切换
with st.sidebar:
    lang = st.radio("🌐 Language / 语言", ["中文", "English"], index=0, horizontal=True, label_visibility="collapsed")
    L = 'zh' if lang == '中文' else 'en'

st.title(f"🔬 {t('page_title', L)}")
st.markdown(t('page_subtitle', L))
st.caption(f'Version: {APP_VERSION}')

# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.header(t('sidebar_header', L))

    sample_type = st.selectbox(
        t('sample_type', L),
        t('sample_type_opts', L),
        index=0
    )
    # Map back to Chinese for internal use
    zh_sample_type = {'Urine':'尿液','Dust':'灰尘','Soil':'土壤','Water':'水','Other':'其他'}.get(sample_type, sample_type)
    if L == 'zh': zh_sample_type = sample_type
    else: sample_type = zh_sample_type

    output_unit = "ng/mL" if sample_type == "尿液" else ("ng/g" if sample_type in ("灰尘","土壤") else "ng/mL")

    col1, col2 = st.columns(2)
    with col1:
        sample_vol = st.number_input(t('sample_vol', L), value=2.0, step=0.1, format="%.1f")
    with col2:
        final_vol = st.number_input(t('final_vol', L), value=0.5, step=0.1, format="%.1f")

    extra_dil = st.number_input(t('extra_dil', L), value=1, step=1, min_value=1)

    use_is = st.radio(
        t('is_correction', L),
        options=t('is_opts', L),
        index=0,
        help=t('is_help', L)
    )
    is_corrected = use_is.startswith("✅")

    if is_corrected:
        auto_cf = 1.0
        st.caption(t('cf_caption_yes', L))
    else:
        auto_cf = round(final_vol / sample_vol * extra_dil, 6)
        st.caption(t('cf_caption_no', L))

    cf_locked_text = t('cf_locked', L) if is_corrected else t('cf_editable', L)
    conversion_factor = st.number_input(
        t('cf_label', L) + cf_locked_text,
        value=auto_cf,
        step=0.001,
        format="%.6f",
        disabled=is_corrected
    )
    spike_conc = st.number_input(t('spike_conc', L), value=10, step=1)

    st.subheader(t('custom_ss', L))
    st.caption(t('custom_ss_help', L))
    custom_ss_text = st.text_area(
        t('custom_ss', L),
        value='',
        placeholder=t('custom_ss_placeholder', L),
        key='custom_ss_text',
        label_visibility='collapsed',
    )
    st.markdown(f"**{t('ss_input_example_title', L)}**")
    st.info(t('ss_input_example_body', L))

    st.subheader(t('custom_is', L))
    st.caption(t('custom_is_help', L))
    custom_is_text = st.text_area(
        t('custom_is', L),
        value='',
        placeholder=t('custom_is_placeholder', L),
        key='custom_is_text',
        label_visibility='collapsed',
    )

    st.divider()
    st.header(t('file_header', L))
    uploaded_file = st.file_uploader(t('upload_label', L), type=["xlsx", "xls", "csv", "tsv"])

    output_name = st.text_input(t('output_name', L), t('output_default', L))
    output_format = st.selectbox(t('output_format', L), ['XLSX', 'CSV'])
    st.caption(t('csv_note', L))

    st.divider()
    st.caption(t('tip', L))

# ============================================================
# 主区域
# ============================================================

demo_path = os.path.join(os.path.dirname(__file__), 'demo_urine_qac_masshunter.xlsx')
with st.sidebar:
    use_demo = st.button(t('demo_btn', L), help=t('demo_help', L), width='stretch')

if use_demo:
    if os.path.exists(demo_path):
        with open(demo_path, 'rb') as demo_file:
            st.session_state.demo_bytes = demo_file.read()
        st.session_state.demo_active = True
    else:
        st.error('Demo file not found. Please upload your own file.' if L == 'en' else '未找到 Demo 文件，请上传自己的文件。')

file_bytes = st.session_state.get('demo_bytes') if st.session_state.get('demo_active') else None
if uploaded_file:
    file_bytes = uploaded_file.getvalue()
    st.session_state.demo_active = False
    st.session_state.pop('demo_bytes', None)

if st.session_state.get('demo_active') and file_bytes:
    st.info(t('demo_loaded', L))

st.subheader(t('blank_workflow_header', L))
if not file_bytes:
    st.info(t('blank_workflow_before_upload', L))

st.subheader(t('ss_is_example_header', L))
example_left, example_right = st.columns(2)
with example_left:
    st.markdown(f"**{t('ss_example_title', L)}**")
    st.caption(t('ss_example_help', L))
    ss_example_rows = (
        [
            {'SS替代物': 'd7-C12-BAC', 'MS1基质加标浓度（ppb）': 4, 'MS2基质加标浓度（ppb）': 8, 'MS3基质加标浓度（ppb）': 12},
            {'SS替代物': 'd9-C10-ATMAC', 'MS1基质加标浓度（ppb）': 2, 'MS2基质加标浓度（ppb）': 2, 'MS3基质加标浓度（ppb）': 4},
        ]
        if L == 'zh' else
        [
            {'SS surrogate': 'd7-C12-BAC', 'MS1 matrix spike (ppb)': 4, 'MS2 matrix spike (ppb)': 8, 'MS3 matrix spike (ppb)': 12},
            {'SS surrogate': 'd9-C10-ATMAC', 'MS1 matrix spike (ppb)': 2, 'MS2 matrix spike (ppb)': 2, 'MS3 matrix spike (ppb)': 4},
        ]
    )
    st.dataframe(pd.DataFrame(ss_example_rows), width='stretch', hide_index=True)
with example_right:
    st.markdown(f"**{t('is_example_title', L)}**")
    st.caption(t('is_example_help', L))
    is_example_rows = (
        [
            {'IS内标化合物名称': 'IS-A'},
            {'IS内标化合物名称': 'IS-B'},
        ]
        if L == 'zh' else
        [
            {'IS internal-standard compound name': 'IS-A'},
            {'IS internal-standard compound name': 'IS-B'},
        ]
    )
    st.dataframe(pd.DataFrame(is_example_rows), width='stretch', hide_index=True)

selected_is = []
selected_ss = []
is_spike_concentrations = {}
ss_concentrations = {}
matrix_spike_concentrations = {}
compound_metadata = {}
mdl_overrides = {}
mql_factor = 3.333333
layout_is_ready = False
mdl_evidence_ready = False
ss_spike_ready = True

if file_bytes:
    st.subheader(t('raw_preview', L))
    try:
        raw_data, blanks, mss, samps, target, is_c, ss_c, all_c = read_raw(file_bytes)

        layout_report = validate_input_layout(blanks, mss, samps, target, is_c, ss_c)
        layout_is_ready = layout_report['ready']
        if layout_report['ready']:
            st.success(f"{t('format_ok', L)}: {layout_report['summary']}")
        else:
            st.error(t('format_bad', L) + ': ' + '；'.join(layout_report['errors']))
        for message in layout_report['warnings']:
            st.warning(message)

        df_raw = read_preview_table(file_bytes)
        # Raw MassHunter header rows contain text while the data rows contain
        # numbers.  Displaying every cell as text avoids Arrow mixed-type
        # coercion warnings without changing the uploaded source data.
        st.dataframe(df_raw.head(8).fillna('').astype(str), width='stretch')

        col1, col2, col3, col4 = st.columns(4)
        col1.metric(t('total_cmpd', L), len(all_c))
        col2.metric(t('target_cmpd', L), len(target))
        col3.metric(t('is_cmpd', L), len(is_c))
        col4.metric(t('ss_cmpd', L), len(ss_c))

        col1, col2, col3 = st.columns(3)
        col1.metric(t('blank_cols', L), len(blanks))
        col2.metric(t('ms_cols', L), len(mss))
        col3.metric(t('sample_cols', L), len(samps))

        st.subheader(t('classification_header', L))
        st.caption(t('classification_help', L))
        classification_seed = []
        default_roles = resolve_roles(all_c, is_c, ss_c)
        for name in all_c:
            meta = compound_metadata_for(name)
            role = 'IS' if name in default_roles['is_compounds'] else ('SS' if name in default_roles['ss_compounds'] else 'Target')
            classification_seed.append({
                t('classification_name', L): name,
                t('classification_type', L): meta['type'],
                t('classification_chain', L): meta['chain_length'],
                t('classification_role', L): role_label(role, L),
            })
        classification_table = st.data_editor(
            pd.DataFrame(classification_seed), hide_index=True, width='stretch',
            disabled=[t('classification_name', L)],
            column_config={
                t('classification_role', L): st.column_config.SelectboxColumn(options=t('role_options', L)),
            }, key='compound_classification_editor',
        )
        for _, item in classification_table.iterrows():
            name = item[t('classification_name', L)]
            compound_metadata[name] = {
                'type': str(item[t('classification_type', L)] or 'Other'),
                'chain_length': str(item[t('classification_chain', L)] or 'NA'),
                'role': role_from_label(item[t('classification_role', L)]),
            }
        selected_is = [name for name, meta in compound_metadata.items() if meta['role'] == 'IS']
        selected_ss = [name for name, meta in compound_metadata.items() if meta['role'] == 'SS']
        custom_is = parse_compound_name_entries(custom_is_text)
        missing_custom_is = [name for name in custom_is if name not in all_c]
        if missing_custom_is:
            st.error(t('custom_is_missing', L) + ', '.join(missing_custom_is))
        valid_custom_is = [name for name in custom_is if name in all_c]
        selected_is = list(dict.fromkeys(selected_is + valid_custom_is))
        custom_ss = parse_compound_name_entries(custom_ss_text)
        missing_custom_ss = [name for name in custom_ss if name not in all_c]
        if missing_custom_ss:
            st.error(t('custom_ss_missing', L) + ', '.join(missing_custom_ss))
        valid_custom_ss = [name for name in custom_ss if name in all_c]
        selected_ss = list(dict.fromkeys(selected_ss + valid_custom_ss))
        overlap = set(selected_is) & set(selected_ss)
        if overlap:
            st.error(t('role_overlap', L) + str(sorted(overlap)))
        if mss:
            st.subheader(t('ms_spike_header', L))
            st.caption(t('ms_spike_help', L))
            st.info(t('ms_spike_example', L))
            roles_for_ms = resolve_roles(all_c, selected_is, selected_ss)
            ss_matrix_spike_concentrations = {}
            with st.sidebar:
                st.subheader(t('ss_spike_sidebar_header', L))
                st.caption(t('ss_spike_sidebar_help', L))
                for ss_index, ss_name in enumerate(roles_for_ms['ss_compounds']):
                    st.markdown(f'**{ss_name}**')
                    ss_matrix_spike_concentrations[ss_name] = {}
                    for ms_index, (_, _, header) in enumerate(mss, 1):
                        ss_value = st.number_input(
                            (f'{ss_name} / MS{ms_index} 基质加标浓度（ppb）' if L == 'zh'
                             else f'{ss_name} / MS{ms_index} matrix-spike concentration (ppb)'),
                            min_value=0.000001,
                            value=None,
                            step=0.1,
                            format='%.6f',
                            key=f'ss_matrix_spike_{ss_index}_{ms_index}',
                        )
                        ss_matrix_spike_concentrations[ss_name][header] = processor.safe_float(ss_value)
            ms_rows = []
            compound_column = t('ms_table_compound', L)
            role_column = t('ms_table_role', L)
            for name in roles_for_ms['target_compounds']:
                ms_rows.append({compound_column: name, role_column: role_label('Target', L), **{header: float(spike_conc) for _, _, header in mss}})
            ms_table = st.data_editor(
                pd.DataFrame(ms_rows),
                column_config={
                    header: st.column_config.NumberColumn(
                        (f'MS{index}基质加标浓度（ppb）' if L == 'zh' else f'MS{index} matrix-spike concentration (ppb)'),
                        min_value=0.000001, step=0.1, format='%.6f', required=True,
                    )
                    for index, (_, _, header) in enumerate(mss, 1)
                },
                disabled=[compound_column, role_column], hide_index=True, width='stretch',
                key='compound_matrix_spike_concentration_table',
            )
            matrix_spike_concentrations = {
                row[compound_column]: {header: processor.safe_float(row[header]) for _, _, header in mss}
                for _, row in ms_table.iterrows()
            }
            matrix_spike_concentrations.update(ss_matrix_spike_concentrations)
            missing_ss_spikes = missing_matrix_spike_entries(
                roles_for_ms['ss_compounds'],
                [header for _, _, header in mss],
                matrix_spike_concentrations,
            )
            ss_spike_ready = not missing_ss_spikes
            if missing_ss_spikes:
                st.error(
                    t('ss_spike_required', L)
                    + ', '.join(f'{compound} / {header}' for compound, header in missing_ss_spikes)
                )
        # The two-dimensional MS table is the source of truth.  Keep these
        # one-value maps only as backward-compatible fallbacks for old files.
        ss_concentrations = {}
        is_spike_concentrations = {}

        # This table reflects the final user-confirmed IS/SS selection, not
        # merely the name-pattern auto-detection shown immediately on import.
        with st.expander(t('view_classify', L)):
            roles_for_display = resolve_roles(all_c, selected_is, selected_ss)
            st.write(f"**{t('target_label', L)}**:", ", ".join(roles_for_display['target_compounds']) if roles_for_display['target_compounds'] else "-")
            st.write(f"**{t('is_label', L)}**:", ", ".join(roles_for_display['is_compounds']) if roles_for_display['is_compounds'] else "-")
            st.write(f"**{t('ss_label', L)}**:", ", ".join(roles_for_display['ss_compounds']) if roles_for_display['ss_compounds'] else "-")
            classification_rows = compound_classification_rows(all_c, selected_is, selected_ss, compound_metadata)
            st.dataframe(pd.DataFrame(classification_rows), width='stretch', hide_index=True)

        target_compounds = resolve_roles(all_c, selected_is, selected_ss)['target_compounds']
        st.markdown(f"### {t('blank_zero_header', L)}")
        st.caption(t('blank_zero_help', L))

        blank_values_by_compound = {
            name: [raw_data.get(name, {}).get(column_letter) for _, column_letter, _ in blanks]
            for name in target_compounds
        }
        for name in target_compounds:
            preliminary = build_blank_mdl_evidence(name, blank_values_by_compound[name], {})
            if preliminary['status'] == 'blank_zero':
                input_col1, input_col2 = st.columns(2)
                with input_col1:
                    calibration = st.number_input(
                        f'{name} {t("calibration", L)}', min_value=0.0, value=0.0,
                        step=0.1, key=f'mdl_cal_{name}'
                    )
                with input_col2:
                    sn = st.number_input(
                        f'{name} {t("sn", L)}', min_value=0.0, value=0.0,
                        step=1.0, key=f'mdl_sn_{name}'
                    )
                mdl_overrides[name] = {
                    'blank_zero': True,
                    'calibration_concentration': calibration,
                    'signal_to_noise': sn,
                }

        evidence_cfg = {'mdl_overrides': mdl_overrides}
        blank_evidence = {
            name: build_blank_mdl_evidence(name, blank_values_by_compound[name], evidence_cfg)
            for name in target_compounds
        }
        status_labels = {
            'blank_zero': t('evidence_blank_zero', L),
            'blank_nonzero': t('evidence_blank_nonzero', L),
            'missing': t('evidence_missing', L),
            'insufficient': t('evidence_insufficient', L),
            'incomplete': t('evidence_incomplete', L),
        }
        blank_evidence_rows = []
        for name, evidence in blank_evidence.items():
            action = (
                t('evidence_enter_snr', L) if evidence['status'] == 'blank_zero' and not evidence['ready']
                else t('evidence_automatic', L) if evidence['ready']
                else t('evidence_unavailable', L)
            )
            blank_evidence_rows.append({
                t('evidence_compound', L): name,
                t('evidence_values', L): ', '.join(f'{value:g}' for value in evidence['blank_values']) or '-',
                t('evidence_valid', L): evidence['valid_count'],
                t('evidence_nonzero', L): evidence['nonzero_count'],
                t('evidence_status', L): status_labels.get(evidence['status'], evidence['status']),
                t('evidence_action', L): action,
                t('evidence_mdl', L): evidence['mdl'],
            })
        st.dataframe(pd.DataFrame(blank_evidence_rows), width='stretch', hide_index=True)

        with st.expander(t('calculation_evidence', L)):
            reason_labels = {
                'blank_zero': t('evidence_reason_blank_zero', L),
                'missing': t('evidence_reason_missing', L),
                'insufficient': t('evidence_reason_insufficient', L),
                'incomplete': t('evidence_reason_incomplete', L),
            }
            for name, evidence in blank_evidence.items():
                st.markdown(f'**{name} — {status_labels.get(evidence["status"], evidence["status"])}**')
                st.write(f'{t("evidence_values", L)}: {evidence["blank_values"] or "-"}')
                st.write(f'{t("evidence_valid", L)}: {evidence["valid_count"]}; {t("evidence_nonzero", L)}: {evidence["nonzero_count"]}')
                if evidence['status'] == 'blank_nonzero':
                    st.write(
                        f'{t("evidence_df", L)}: {evidence["degrees_of_freedom"]}; '
                        f'{t("evidence_t", L)}: {evidence["t_value"]:.6g}; '
                        f'{t("evidence_mean", L)}: {evidence["mean"]:.6g}; '
                        f'{t("evidence_sd", L)}: {evidence["sd"]:.6g}'
                    )
                if evidence['ready']:
                    st.code(f'{t("evidence_formula", L)}: {evidence["formula"]} = {evidence["mdl"]:.6g} ppb')
                else:
                    st.warning(reason_labels.get(evidence['status'], evidence['reason']))
                st.divider()

        mdl_evidence_ready = bool(target_compounds) and all(item['ready'] for item in blank_evidence.values())


    except Exception as e:
        st.warning(f"{t('preview_warn', L)}: {e}")

# ============================================================
# 处理按钮
# ============================================================
st.divider()
process_btn = st.button(
    t('process_btn', L), type="primary",
    disabled=(file_bytes is None or not layout_is_ready or not mdl_evidence_ready or not ss_spike_ready),
    width='stretch',
)

if process_btn and file_bytes:
    with st.spinner(t('processing', L)):

        config = {
            'sample_type': sample_type,
            'sample_volume_ml': float(sample_vol),
            'final_volume_ml': float(final_vol),
            'extra_dilution': int(extra_dil),
            'conversion_factor': float(conversion_factor),
            'spike_conc_ppb': int(spike_conc),
            'matrix_spike_concentrations': matrix_spike_concentrations,
            'compound_metadata': compound_metadata,
            'is_compounds': selected_is,
            'is_spike_concentrations': is_spike_concentrations,
            'is_corrected': is_corrected,
            'ss_compounds': selected_ss,
            'ss_spike_concentrations': ss_concentrations,
            'mdl_overrides': mdl_overrides,
            'language': L,
            'masshunter_unit': 'ppb',
            'output_unit': output_unit,
            'blank_handling': 'ND',
            'input_file': '',
            'output_file': output_name,
            'output_format': output_format.lower(),
            'input_bytes': file_bytes,
        }

        try:
            output_bytes, filename = process(config=config, return_bytes=True)

            st.success(t('success', L))

            roles = resolve_roles(all_c, selected_is, selected_ss)
            preview_cfg = {**config, 'target_compounds': roles['target_compounds']}
            preview_rows = compute_preview_summary(raw_data, blanks, samps, preview_cfg)
            if preview_rows:
                st.subheader(t('preview_stats', L))
                st.caption(t('preview_stats_help', L))
                st.dataframe(pd.DataFrame(preview_rows), width='stretch')

            final_preview_rows = compute_preview_final_table(raw_data, blanks, samps, preview_cfg)
            if final_preview_rows:
                st.subheader(t('preview_final', L))
                st.dataframe(pd.DataFrame(final_preview_rows), width='stretch')

            st.subheader(t('result_preview', L))
            if output_format == 'CSV':
                # CSV is a flat, formula-free report; preview the actual exported values.
                # The summary and final-concentration sections have different widths,
                # so read rows directly and pad them before displaying the table.
                import csv
                csv_rows = list(csv.reader(io.StringIO(output_bytes.decode('utf-8-sig'))))
                width = max((len(row) for row in csv_rows), default=0)
                csv_preview = pd.DataFrame([row + [''] * (width - len(row)) for row in csv_rows])
                st.dataframe(csv_preview, width='stretch')
            else:
                # openpyxl cannot calculate formulas. Show the numerically evaluated
                # previews above instead of exposing formula strings in the result view.
                st.info(t('preview_stats_help', L))
                st.caption(t('result_formula_note', L))
                st.dataframe(pd.DataFrame(preview_rows), width='stretch')
                st.dataframe(pd.DataFrame(final_preview_rows), width='stretch')

            st.divider()
            st.download_button(
                label=t('download_btn', L),
                data=output_bytes,
                file_name=filename,
                mime=("text/csv" if output_format == 'CSV' else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                width='stretch',
                type="primary",
            )

        except Exception as e:
            st.error(f"{t('error', L)}: {e}")
            import traceback
            st.code(traceback.format_exc())
