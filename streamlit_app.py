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

st.set_page_config(
    page_title="LC-MS/MS 数据处理智能体",
    page_icon="🔬",
    layout="wide",
)

st.title("🔬 LC-MS/MS 数据处理智能体")
st.markdown("上传 MassHunter 原始数据 → 一键生成完整分析表格 → 下载")

# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.header("📋 实验参数")

    sample_type = st.selectbox("样本类型", ["尿液", "灰尘", "土壤", "水", "其他"], index=0)
    output_unit = "ng/mL" if sample_type == "尿液" else ("ng/g" if sample_type in ("灰尘","土壤") else "ng/mL")

    col1, col2 = st.columns(2)
    with col1:
        sample_vol = st.number_input("取样量 (mL/g)", value=2.0, step=0.1, format="%.1f")
    with col2:
        final_vol = st.number_input("定容体积 (mL)", value=0.5, step=0.1, format="%.1f")

    extra_dil = st.number_input("额外稀释倍数", value=1, step=1, min_value=1)

    # 内标校正开关
    use_is = st.radio(
        "数据是否经过内标（IS）校正？",
        options=["✅ 是，有内标校正", "❌ 否，无内标校正"],
        index=0,
        help="内标校正后，导出值已是原始样本浓度，换算因子=1；无内标校正则需要手动换算"
    )
    is_corrected = use_is.startswith("✅")

    if is_corrected:
        auto_cf = 1.0
        st.caption("💡 有内标校正，导出值已是原始样本浓度，换算因子 = 1")
    else:
        auto_cf = round(final_vol / sample_vol * extra_dil, 6)
        st.caption("💡 无内标校正，导出值是进样瓶浓度，换算因子 = 定容体积 ÷ 取样量 × 稀释倍数")

    conversion_factor = st.number_input(
        "换算因子" + ("（已锁定）" if is_corrected else "（可手动覆盖）"),
        value=auto_cf,
        step=0.001,
        format="%.6f",
        disabled=is_corrected
    )
    spike_conc = st.number_input("基质加标浓度 (ppb)", value=10, step=1)
    ss_spike_conc = st.number_input("SS 替代物加标浓度 (ppb)", value=10, step=1, help="替代物自身的理论加标浓度，可能不同于基质加标浓度")

    st.divider()
    st.header("📁 文件")
    uploaded_file = st.file_uploader("上传原始数据 Excel", type=["xlsx"])
    output_name = st.text_input("输出文件名", "已处理数据.xlsx")

    st.divider()
    st.caption("上传文件 → 调参数 → 点处理 → 下载结果")

# ============================================================
# 主区域
# ============================================================

# 上传后先保存 bytes 并预览
file_bytes = None
if uploaded_file:
    file_bytes = uploaded_file.getvalue()

    st.subheader("📊 原始数据预览")
    try:
        # 智能检测化合物
        raw_data, blanks, mss, samps, target, is_c, ss_c, all_c = read_raw(file_bytes)

        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)
        st.dataframe(df_raw.head(8), use_container_width=True)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("化合物总数", len(all_c))
        col2.metric("目标化合物", len(target))
        col3.metric("IS内标", len(is_c))
        col4.metric("SS替代物", len(ss_c))

        with st.expander("查看化合物分类"):
            st.write("**目标化合物**:", ", ".join(target) if target else "无")
            st.write("**IS内标**:", ", ".join(is_c) if is_c else "无")
            st.write("**SS替代物**:", ", ".join(ss_c) if ss_c else "无")

        col1, col2, col3 = st.columns(3)
        col1.metric("BLANK列", len(blanks))
        col2.metric("MS列", len(mss))
        col3.metric("样品列", len(samps))

    except Exception as e:
        st.warning(f"预览时请注意: {e}")

# ============================================================
# 处理按钮
# ============================================================
st.divider()
process_btn = st.button("🚀 开始处理", type="primary", disabled=(uploaded_file is None), use_container_width=True)

if process_btn and file_bytes:
    with st.spinner("正在处理数据..."):

        config = {
            'sample_type': sample_type,
            'sample_volume_ml': float(sample_vol),
            'final_volume_ml': float(final_vol),
            'extra_dilution': int(extra_dil),
            'conversion_factor': float(conversion_factor),
            'spike_conc_ppb': int(spike_conc),
            'ss_spike_conc_ppb': int(ss_spike_conc),
            'masshunter_unit': 'ppb',
            'output_unit': output_unit,
            'blank_handling': 'ND',
            'input_file': '',
            'output_file': output_name,
            'input_bytes': file_bytes,
        }

        try:
            output_bytes, filename = process(config=config, return_bytes=True)

            st.success(f"✅ 处理完成！")

            # 预览
            st.subheader("📋 处理结果预览")
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
                    st.caption(f"{ws.max_row}行 × {ws.max_column}列")
            wb.close()

            st.divider()
            st.download_button(
                label="⬇️ 下载处理结果",
                data=output_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

        except Exception as e:
            st.error(f"处理失败: {e}")
            import traceback
            st.code(traceback.format_exc())
