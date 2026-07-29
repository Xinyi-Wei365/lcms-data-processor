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
from process_lcms_data import process, read_raw, classify_compounds

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
    'cf_caption_yes':       {'zh': '💡 有内标校正，仪器已自动将进样瓶浓度换算为原始尿液浓度，换算因子 = 1',
                             'en': '💡 IS corrected: instrument has already converted vial concentration to original sample concentration. CF = 1'},
    'cf_caption_no':        {'zh': '💡 无内标校正，导出值是进样瓶浓度，换算因子用于将其换算为原始尿液浓度：定容体积 ÷ 取样量 × 稀释倍数',
                             'en': '💡 Not IS corrected: exported value is vial concentration. CF = Final Vol ÷ Sample Vol × Dilution to convert to original concentration'},
    'cf_label':             {'zh': '换算因子',                                   'en': 'Conversion Factor'},
    'cf_locked':            {'zh': '（已锁定）',                                  'en': ' (Locked)'},
    'cf_editable':          {'zh': '（可手动覆盖）',                              'en': ' (Editable)'},
    'spike_conc':           {'zh': '基质加标浓度 (ppb)',                          'en': 'Matrix Spike Conc (ppb)'},
    'ss_spike_conc':        {'zh': 'SS 替代物加标浓度 (ppb)',                     'en': 'SS Surrogate Spike Conc (ppb)'},
    'ss_spike_help':        {'zh': '替代物自身的理论加标浓度，可能不同于基质加标浓度',
                             'en': 'Theoretical spike concentration of surrogates, may differ from matrix spike conc'},
    'file_header':          {'zh': '📁 文件',                                    'en': '📁 File'},
    'upload_label':         {'zh': '上传原始数据 Excel',                           'en': 'Upload Raw Data Excel'},
    'output_name':          {'zh': '输出文件名',                                  'en': 'Output Filename'},
    'output_default':       {'zh': '已处理数据.xlsx',                             'en': 'processed_data.xlsx'},
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
}

def t(key, lang='zh'):
    """Get translation for key in given language"""
    if key in T:
        return T[key].get(lang, T[key].get('zh', key))
    return key

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
    st.caption('SS 替代物理论加标浓度（分别设置）：')
    col_ss1, col_ss2 = st.columns(2)
    with col_ss1:
        ss_spike_d7 = st.number_input('d7-C12-BAC (ppb)', value=4, step=1)
    with col_ss2:
        ss_spike_d9 = st.number_input('d9-C10-ATMAC (ppb)', value=4, step=1)

    st.divider()
    st.header(t('file_header', L))
    uploaded_file = st.file_uploader(t('upload_label', L), type=["xlsx"])

    # Demo 数据按钮
    use_demo = st.button(t('demo_btn', L), help=t('demo_help', L), use_container_width=True)
    if use_demo:
        st.session_state.demo_active = True

    output_name = st.text_input(t('output_name', L), t('output_default', L))

    st.divider()
    st.caption(t('tip', L))

# ============================================================
# 主区域
# ============================================================

file_bytes = None
demo_loaded = st.session_state.get('demo_active', False)

# Demo 数据加载
if demo_loaded:
    demo_path = os.path.join(os.path.dirname(__file__), 'demo_urine_qac_masshunter.xlsx')
    if os.path.exists(demo_path):
        with open(demo_path, 'rb') as f:
            file_bytes = f.read()
        st.info("📥 " + ({'zh':'已加载 Demo 数据，可直接点击处理或替换为自己的文件','en':'Demo data loaded. Click Process or upload your own file'}[L]))
    else:
        st.warning("Demo file not found. Please upload your own data.")
        st.session_state.demo_active = False

if uploaded_file:
    file_bytes = uploaded_file.getvalue()
    st.session_state.demo_active = False

if file_bytes:
    st.subheader(t('raw_preview', L))
    try:
        raw_data, blanks, mss, samps, target, is_c, ss_c, all_c = read_raw(file_bytes)

        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)
        st.dataframe(df_raw.head(8), use_container_width=True)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric(t('total_cmpd', L), len(all_c))
        col2.metric(t('target_cmpd', L), len(target))
        col3.metric(t('is_cmpd', L), len(is_c))
        col4.metric(t('ss_cmpd', L), len(ss_c))

        with st.expander(t('view_classify', L)):
            st.write(f"**{t('target_label', L)}**:", ", ".join(target) if target else "-")
            st.write(f"**{t('is_label', L)}**:", ", ".join(is_c) if is_c else "-")
            st.write(f"**{t('ss_label', L)}**:", ", ".join(ss_c) if ss_c else "-")

        col1, col2, col3 = st.columns(3)
        col1.metric(t('blank_cols', L), len(blanks))
        col2.metric(t('ms_cols', L), len(mss))
        col3.metric(t('sample_cols', L), len(samps))

    except Exception as e:
        st.warning(f"{t('preview_warn', L)}: {e}")

# ============================================================
# 处理按钮
# ============================================================
st.divider()
process_btn = st.button(t('process_btn', L), type="primary", disabled=(file_bytes is None), use_container_width=True)

if process_btn and file_bytes:
    with st.spinner(t('processing', L)):

        config = {
            'sample_type': sample_type,
            'sample_volume_ml': float(sample_vol),
            'final_volume_ml': float(final_vol),
            'extra_dilution': int(extra_dil),
            'conversion_factor': float(conversion_factor),
            'spike_conc_ppb': int(spike_conc),
            'ss_spike_d7_ppb': int(ss_spike_d7),
            'ss_spike_d9_ppb': int(ss_spike_d9),
            'masshunter_unit': 'ppb',
            'output_unit': output_unit,
            'blank_handling': 'ND',
            'input_file': '',
            'output_file': output_name,
            'input_bytes': file_bytes,
        }

        try:
            output_bytes, filename = process(config=config, return_bytes=True)

            st.success(t('success', L))

            st.subheader(t('result_preview', L))
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(output_bytes), data_only=False)
            tabs = st.tabs([f"{i+1}. {n}" for i, n in enumerate(wb.sheetnames)])

            for tab, name in zip(tabs, wb.sheetnames):
                with tab:
                    ws = wb[name]
                    preview_data = []
                    for r in ws.iter_rows(min_row=1, max_row=min(12, ws.max_row), max_col=min(12, ws.max_column), values_only=True):
                        preview_data.append(list(r))
                    if preview_data:
                        df_preview = pd.DataFrame(preview_data)
                        if len(preview_data) > 1:
                            cols = []
                            seen = {}
                            for v in df_preview.iloc[0]:
                                s = str(v) if v is not None else ''
                                if s in seen:
                                    seen[s] += 1
                                    cols.append(f'{s}_{seen[s]}')
                                else:
                                    seen[s] = 0
                                    cols.append(s)
                            df_preview.columns = cols
                            df_preview = df_preview.iloc[1:]
                        st.dataframe(df_preview, use_container_width=True)
                    st.caption(f"{ws.max_row}{t('rows_cols', L)}{ws.max_column}{t('rows_cols_suffix', L)}")
            wb.close()

            st.divider()
            st.download_button(
                label=t('download_btn', L),
                data=output_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

        except Exception as e:
            st.error(f"{t('error', L)}: {e}")
            import traceback
            st.code(traceback.format_exc())
