# LC-MS/MS 数据处理智能体

从 MassHunter 导出的原始数据 → 一键生成完整分析表格。

## 功能

- 上传原始 LC-QTOF 质谱数据（.xlsx）
- 自动识别化合物、BLANK 列、MS 列、样品列
- 生成 6 个工作表：基质加标浓度、空白检出限、瓶内实测浓度、最终计算浓度、统计辅助表、计算说明
- 所有统计列带 Excel 公式，换算因子可调
- 下载处理结果

## 本地运行

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## 在线使用

访问 Streamlit Cloud 部署链接（见仓库 About 或部署后生成的 URL）。

界面只保留一个“数据是否经过内标（IS）校正？”选项：已校正时换算因子为 1；未校正时仅按取样体积、定容体积和稀释倍数进行换算，不会凭浓度猜测 IS 响应比。
