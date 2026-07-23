#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LC-MS/MS 数据处理智能体 — Streamlit 可视化界面
运行: streamlit run streamlit_app.py
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from process_lcms_data import process, TARGET_COMPS, ALL_COMPS

st.set_page_config(
    page_title="LC-MS/MS 数据处理智能体",
    page_icon="🔬",
    layout="wide",
)

st.title("🔬 LC-MS/MS 数据处理智能体")
st.markdown("从 MassHunter 导出数据 → 一键生成完整分析表格")

# ============================================================
# 侧边栏：参数配置
# ============================================================
with st.sidebar:
    st.header("📋 实验参数")

    sample_type = st.selectbox(
        "样本类型", ["尿液", "灰尘", "土壤", "水", "其他"],
        index=0,
        help="选择样本类型，影响输出单位（尿液→ng/mL，灰尘→ng/g）"
    )

    if sample_type == "尿液":
        output_unit = "ng/mL"
    elif sample_type in ("灰尘", "土壤"):
        output_unit = "ng/g"
    else:
        output_unit = "ng/mL"

    col1, col2 = st.columns(2)
    with col1:
        sample_vol = st.number_input("取样量 (mL 或 g)", value=2.0, step=0.1, format="%.1f")
    with col2:
        final_vol = st.number_input("定容体积 (mL)", value=0.5, step=0.1, format="%.1f")

    extra_dil = st.number_input("额外稀释倍数", value=1, step=1, min_value=1)

    # 换算因子自动计算 + 手动覆盖
    auto_cf = round(final_vol / sample_vol * extra_dil, 6)
    st.caption(f"💡 自动计算换算因子: {final_vol}/{sample_vol}×{extra_dil} = **{auto_cf}**")

    conversion_factor = st.number_input(
        "换算因子",
        value=1.0,
        step=0.001,
        format="%.6f",
        help="若 MassHunter 已是原始浓度 → 1；若是定容液浓度 → 自动计算值"
    )

    spike_conc = st.number_input("基质加标浓度 (ppb)", value=10, step=1)

    st.divider()
    st.header("📁 文件设置")

    uploaded_file = st.file_uploader(
        "上传原始数据 Excel",
        type=["xlsx"],
        help="选择 MassHunter 导出的原始数据文件"
    )

    output_name = st.text_input("输出文件名", "已处理数据.xlsx")

    st.divider()
    st.markdown("### 📖 使用说明")
    st.markdown("""
    1. 上传 MassHunter 导出的原始数据
    2. 调整实验参数
    3. 点击「开始处理」
    4. 预览结果后下载
    """)

# ============================================================
# 主区域
# ============================================================

# 上传文件后预览
if uploaded_file:
    st.subheader("📊 原始数据预览")

    # 保存临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name

    try:
        df_raw = pd.read_excel(tmp_path, sheet_name=0, header=None)
        st.dataframe(df_raw.head(10), use_container_width=True)
        st.caption(f"共 {df_raw.shape[0]} 行 × {df_raw.shape[1]} 列")

        # 检测化合物
        compounds_found = []
        for i in range(2, min(df_raw.shape[0], 50)):
            v = df_raw.iloc[i, 1] if df_raw.shape[1] > 1 else None
            if pd.notna(v) and str(v).strip():
                compounds_found.append(str(v).strip())

        st.info(f"🔍 检测到 **{len(compounds_found)}** 个化合物")
        with st.expander("查看化合物列表"):
            for c in compounds_found:
                matched = "✅" if c in ALL_COMPS else "⚠️"
                st.text(f"  {matched} {c}")
    except Exception as e:
        st.error(f"读取文件失败: {e}")
    finally:
        os.unlink(tmp_path)

# 处理按钮
st.divider()
col_btn, col_status = st.columns([1, 3])
with col_btn:
    process_btn = st.button(
        "🚀 开始处理",
        type="primary",
        disabled=(uploaded_file is None),
        use_container_width=True,
    )

if process_btn and uploaded_file:
    with st.spinner("正在处理数据..."):
        # 保存上传文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(uploaded_file.getvalue())
            input_path = tmp.name

        # 构建配置
        config = {
            'sample_type': sample_type,
            'sample_volume_ml': float(sample_vol),
            'final_volume_ml': float(final_vol),
            'extra_dilution': int(extra_dil),
            'conversion_factor': float(conversion_factor),
            'spike_conc_ppb': int(spike_conc),
            'masshunter_unit': 'ppb',
            'output_unit': output_unit,
            'blank_handling': 'ND',
            'input_file': input_path,
            'output_file': output_name,
        }

        try:
            # 处理
            output_bytes, filename = process(config=config, return_bytes=True)

            # 清理临时文件
            os.unlink(input_path)

            st.success(f"✅ 处理完成！生成 {len(TARGET_COMPS)} 个目标化合物的数据")

            # 预览结果
            st.subheader("📋 处理结果预览")

            # 读取生成的文件进行预览
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_out:
                tmp_out.write(output_bytes)
                tmp_out_path = tmp_out.name

            import openpyxl
            wb = openpyxl.load_workbook(tmp_out_path, data_only=False)
            tabs = st.tabs([f"{i+1}. {n}" for i, n in enumerate(wb.sheetnames)])

            for i, (tab, name) in enumerate(zip(tabs, wb.sheetnames)):
                with tab:
                    ws = wb[name]
                    # 读取前15行、前15列作为预览
                    preview_data = []
                    for row in ws.iter_rows(min_row=1, max_row=min(15, ws.max_row),
                                             max_col=min(15, ws.max_column), values_only=True):
                        preview_data.append(list(row))

                    if preview_data:
                        df_preview = pd.DataFrame(preview_data)
                        # 第一行作为列名
                        if len(preview_data) > 1:
                            df_preview.columns = df_preview.iloc[0]
                            df_preview = df_preview.iloc[1:]
                        st.dataframe(df_preview, use_container_width=True)
                    st.caption(f"工作表: {ws.max_row}行 × {ws.max_column}列")

            wb.close()
            os.unlink(tmp_out_path)

            # 下载按钮
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
            # 清理
            if os.path.exists(input_path):
                os.unlink(input_path)
else:
    # 显示参数摘要
    st.info("👈 请在左侧上传原始数据文件，然后点击「开始处理」")

    with st.expander("📌 当前参数摘要"):
        col_a, col_b = st.columns(2)
        with col_a:
            st.metric("样本类型", sample_type)
            st.metric("取样量", f"{sample_vol} mL")
            st.metric("定容体积", f"{final_vol} mL")
        with col_b:
            st.metric("换算因子", conversion_factor)
            st.metric("加标浓度", f"{spike_conc} ppb")
            st.metric("输出单位", output_unit)

    st.markdown("---")
    st.markdown("### 📐 表格结构说明")
    st.markdown(f"""
    | 工作表 | 说明 | 化合物数 |
    |--------|------|---------|
    | Matrix spike | 基质加标浓度 + 回收率 + 统计 | {len(ALL_COMPS)} (含IS/SS) |
    | Blanks_MDL | 空白检出限 + MDL | {len(ALL_COMPS)} |
    | Conc. in bottle | 瓶内实测浓度 | {len(TARGET_COMPS)} |
    | Final conc | 最终计算浓度 + 统计 | {len(TARGET_COMPS)} |
    | 统计计算结果 | 辅助统计表 | {len(TARGET_COMPS)} |
    | 计算说明 | 公式说明 | - |
    """)
