import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outputPath = "demo_urine_qac_masshunter.xlsx";
const workbook = Workbook.create();
const data = workbook.worksheets.add("MassHunter Demo");
const notes = workbook.worksheets.add("Demo Guide");

const headers = [
  "Compound Name", "Transition", "BLANK1", "BLANK2", "BLANK3", "BLANK4",
  "MS1", "MS2", "F1", "F2", "F3", "F4", "F5",
];
const rows = [
  ["C8-BAC", "248.2>91.0", 0.08, 0.10, 0.09, 0.11, 9.70, 10.20, 0.05, 0.16, 0.80, 1.30, null],
  ["C10-BAC", "276.3>91.1", 0.04, 0.05, 0.06, 0.05, 10.20, 9.80, 0, 0.07, 0.30, 0.46, 0.55],
  ["C8-DDAC", "270.3>158.2", 0.06, 0.05, 0.07, 0.06, 9.90, 10.10, 0.03, 0.11, 0.40, 0.77, 1.05],
  ["C8-ATMAC", "172.2>71.1", 0.02, 0.03, 0.02, 0.03, 10.10, 10.00, 0.01, 0.04, 0.18, 0.25, 0.32],
  ["C8-PFAS", "499.0>80.0", 0.01, 0.02, 0.01, 0.02, 10.00, 9.90, 0.01, 0.03, 0.12, 0.25, 0.34],
  ["d7-C12-BAC", "311.3>98.1", 0.18, 0.20, 0.19, 0.21, 3.85, 4.12, 3.90, 3.95, 4.04, 3.98, 4.08],
  ["d9-C10-ATMAC", "209.3>71.1", 0.21, 0.20, 0.22, 0.21, 4.04, 3.96, 3.98, 4.01, 4.03, 3.99, 4.00],
  ["C13-C12-BAC IS", "317.3>98.1", 0.25, 0.24, 0.26, 0.25, 4.00, 4.00, 4.01, 3.99, 4.02, 4.00, 4.01],
];

data.getRange(`A1:M${rows.length + 1}`).values = [headers, ...rows];
data.getRange("A1:M1").format = {
  fill: "#1F4E78", font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center",
};
data.getRange(`A2:B${rows.length + 1}`).format = { horizontalAlignment: "left" };
data.getRange(`C2:M${rows.length + 1}`).format = { numberFormat: "0.000000" };
data.getRange("A1:M1").format.wrapText = true;
data.getRange("A:M").format.autofitColumns();
data.getRange("A:A").format.columnWidth = 26;
data.getRange("B:B").format.columnWidth = 16;
data.getRange("C:M").format.columnWidth = 12;
data.freezePanes.freezeRows(1);
data.showGridLines = false;

const guide = [
  ["Demo urine MassHunter data / 尿液 MassHunter 示例数据"],
  ["Item", "What this demo demonstrates / 本 Demo 展示内容"],
  ["Input layout", "Four blanks, two matrix spikes (MS1/MS2), and five samples (F1-F5). / 4 个空白、2 个基质加标、5 个样品。"],
  ["Targets", "C8-BAC, C10-BAC, C8-DDAC (displayed as C8-DADMAC after processing), C8-ATMAC, and C8-PFAS. / 目标物包含 QAC 与非 QAC。"],
  ["SS", "d7-C12-BAC and d9-C10-ATMAC are automatically suggested as SS. Both use 4 ppb in MS1/MS2 so the Demo runs with the default settings. / SS 使用自身加标浓度计算回收率。"],
  ["IS", "C13-C12-BAC IS is automatically suggested as IS; it is record-only and has no recovery. / IS 仅记录加入浓度，不计算回收率。"],
  ["MDL", "All blank values are numeric non-zero. Enter same-level low-spike replicate results if the full MAX(t×SD(spike), mean(blank)+t×SD(blank)) MDL method is required. / 所有空白非零。"],
  ["Direct use", "Click Load Demo Data and then Start Processing. To practise different MS1/MS2/MS3 spike amounts, edit the compound × MS grid before processing. / 点击加载 Demo 数据后即可开始处理；也可编辑化合物 × MS 表格练习不同加标浓度。"],
];
notes.getRange(`A1:B${guide.length}`).values = guide;
notes.getRange("A1:B1").merge();
notes.getRange("A1").format = { fill: "#1F4E78", font: { bold: true, color: "#FFFFFF", size: 14 }, horizontalAlignment: "center" };
notes.getRange("A2:B2").format = { fill: "#D9EAF7", font: { bold: true, color: "#1F4E78" } };
notes.getRange(`A3:A${guide.length}`).format = { fill: "#F2F2F2", font: { bold: true } };
notes.getRange(`A1:B${guide.length}`).format.wrapText = true;
notes.getRange("A:A").format.columnWidth = 18;
notes.getRange("B:B").format.columnWidth = 100;
notes.getRange(`A1:B${guide.length}`).format.autofitRows();
notes.showGridLines = false;

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
