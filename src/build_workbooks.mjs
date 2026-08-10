import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const projectRoot = process.env.U8_PROJECT_ROOT;
const outputDir = process.env.U8_OUTPUT_DIR;
const qaDir = process.env.U8_QA_DIR;
const dataPath = process.env.U8_DATA_JSON;
const buildPart = process.env.U8_BUILD_PART || "caats";

if (!projectRoot || !outputDir || !qaDir || !dataPath) {
  throw new Error("缺少U8_PROJECT_ROOT/U8_OUTPUT_DIR/U8_QA_DIR/U8_DATA_JSON环境变量");
}

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(qaDir, { recursive: true });
const data = JSON.parse(await fs.readFile(dataPath, "utf8"));
const material = data.material_rows;
const warehouses = data.warehouse_rows;
const groups = data.group_rows;
const master = data.master;
const materialEnd = material.length + 1;
const warehouseEnd = warehouses.length + 3;
const amountTolerance = Number(data.config.amount_tolerance);
const quantityTolerance = Number(data.config.quantity_tolerance);

const numberFormat = "#,##0.00;[Red](#,##0.00);-";
const quantityFormat = "#,##0.0000;[Red](#,##0.0000);-";
const palette = {
  title: "#1F4E78",
  tableHeader: "#4472C4",
  sectionHeader: "#5B6573",
  incomeHeader: "#0F6B78",
  issueHeader: "#C55A11",
  closingHeader: "#8064A2",
  resultHeader: "#2F75B5",
  grid: "#D9E2F3",
};
const thinBorder = { preset: "all", style: "thin", color: palette.grid };
const titleFormat = { fill: palette.title, font: { bold: true, color: "#FFFFFF", fontSize: 15 }, horizontalAlignment: "left", verticalAlignment: "center", rowHeight: 28 };
const headerFormat = (fill) => ({
  fill,
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: thinBorder,
  rowHeight: 36,
});

function applyStatusFormatting(sheet, range) {
  sheet.getRange(range).conditionalFormats.add("containsText", { text: "PASS", format: { fill: "#E2F0D9", font: { bold: true, color: "#006100" } } });
  sheet.getRange(range).conditionalFormats.add("containsText", { text: "REVIEW", format: { fill: "#FFF2CC", font: { bold: true, color: "#9C5700" } } });
}

function applyTableGrid(sheet, range) {
  sheet.getRange(range).format.borders = thinBorder;
}

async function saveInspection(workbook, name, ranges) {
  const output = [];
  for (const [sheetName, range] of ranges) {
    const inspected = await workbook.inspect({
      kind: "table",
      range: `${sheetName}!${range}`,
      include: "values,formulas",
      tableMaxRows: 30,
      tableMaxCols: 24,
      maxChars: 20000,
    });
    output.push(inspected.ndjson);
  }
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!",
    options: { useRegex: true, maxResults: 300 },
    summary: `${name} formula error scan`,
  });
  output.push(errors.ndjson);
  await fs.writeFile(`${qaDir}/${name}_inspection.txt`, output.join("\n"));
}

async function renderSheets(workbook, name, specs) {
  for (const [sheetName, range, scale] of specs) {
    const image = await workbook.render({ sheetName, range, scale, format: "png" });
    const safeName = sheetName.replaceAll(/[\\/:*?"<>|]/g, "_");
    await fs.writeFile(`${qaDir}/${name}_${safeName}.png`, new Uint8Array(await image.arrayBuffer()));
  }
}

async function buildCaatsWorkbook() {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(`${projectRoot}/CAATS交付模板.xlsx`));
  const resultSheet = workbook.worksheets.getItem("Sheet2");
  const resultSheetName = "01_物料月维度明细";
  const continuousDetailName = "03_连续滚算明细";
  resultSheet.name = resultSheetName;
  const summarySheet = workbook.worksheets.add("02_物料月维度结果汇总");
  const continuousDetailSheet = workbook.worksheets.add(continuousDetailName);
  const continuousSummarySheet = workbook.worksheets.add("04_连续滚算汇总");
  const differenceSheet = workbook.worksheets.add("05_存在差异明细");
  const offsetSheet = workbook.worksheets.add("06_跨仓抵销");

  resultSheet.showGridLines = false;
  resultSheet.freezePanes.freezeRows(1);
  resultSheet.freezePanes.freezeColumns(4);
  resultSheet.getRange("AF1:AK1").values = [["CAATS月末平均价", "CAATS结存数量", "结存数量差异", "结果状态", "主要差异原因", "发出差异绝对值"]];
  const normalizedHeaders = resultSheet.getRange("A1:AK1").values[0].map((value) => typeof value === "string" ? value.replaceAll("CAATs", "CAATS") : value);
  resultSheet.getRange("A1:AK1").values = [normalizedHeaders];
  const resultValues = material.map((row) => [
    row.code, row.month, row.stock_code || null, row.name, row.spec, row.unit, row.weight || null, row.reg || null, row.manufacturer || null,
    row["期初数量"], row["期初单价"], row["期初金额"], row["收入数量"], null, null, row["收入单价"], row["收入金额"], null, null,
    row["发出数量"], null, null, row["发出单价"], row["发出金额"], null, null,
    row["结存数量"], row["结存单价"], row["结存金额"], null, null,
    null, null, null, null, row.reason, null,
  ]);
  resultSheet.getRange(`A2:AK${materialEnd}`).values = resultValues;
  const formulas = { N: [], O: [], R: [], S: [], U: [], V: [], Y: [], Z: [], AD: [], AE: [], AF: [], AG: [], AH: [], AI: [], AK: [] };
  for (let row = 2; row <= materialEnd; row += 1) {
    formulas.N.push([`=M${row}`]);
    formulas.O.push([`=M${row}-N${row}`]);
    formulas.R.push([`=Q${row}`]);
    formulas.S.push([`=Q${row}-R${row}`]);
    formulas.U.push([`=T${row}`]);
    formulas.V.push([`=T${row}-U${row}`]);
    formulas.Y.push([`=U${row}*AF${row}`]);
    formulas.Z.push([`=X${row}-Y${row}`]);
    formulas.AD.push([`=L${row}+R${row}-Y${row}`]);
    formulas.AE.push([`=AC${row}-AD${row}`]);
    formulas.AF.push([`=IF(J${row}+N${row}=0,0,(L${row}+R${row})/(J${row}+N${row}))`]);
    formulas.AG.push([`=J${row}+N${row}-U${row}`]);
    formulas.AH.push([`=AA${row}-AG${row}`]);
    formulas.AI.push([`=IF(ABS(Z${row})<=${amountTolerance},"PASS","REVIEW")`]);
    formulas.AK.push([`=ABS(Z${row})`]);
  }
  for (const [column, values] of Object.entries(formulas)) {
    resultSheet.getRange(`${column}2:${column}${materialEnd}`).formulas = values;
  }
  applyTableGrid(resultSheet, `A1:AK${materialEnd}`);
  resultSheet.getRange("A1:I1").format = headerFormat(palette.title);
  resultSheet.getRange("J1:L1").format = headerFormat(palette.sectionHeader);
  resultSheet.getRange("M1:S1").format = headerFormat(palette.incomeHeader);
  resultSheet.getRange("T1:Z1").format = headerFormat(palette.issueHeader);
  resultSheet.getRange("AA1:AE1").format = headerFormat(palette.closingHeader);
  resultSheet.getRange("AF1:AK1").format = headerFormat(palette.resultHeader);
  resultSheet.getRange("A1:AK1").format.rowHeight = 48;
  resultSheet.getRange(`A2:AK${materialEnd}`).format.font = { typeface: "宋体", fontSize: 9, color: "#000000" };
  resultSheet.getRange(`J2:AK${materialEnd}`).format.numberFormat = numberFormat;
  resultSheet.getRange(`A2:A${materialEnd}`).format.numberFormat = "0000";
  resultSheet.getRange(`B2:B${materialEnd}`).format.numberFormat = "0";
  for (const column of ["A", "B", "C", "F", "G", "I"]) resultSheet.getRange(`${column}1:${column}${materialEnd}`).format.columnWidth = 11;
  for (const column of ["D", "H"]) resultSheet.getRange(`${column}1:${column}${materialEnd}`).format.columnWidth = 24;
  resultSheet.getRange(`E1:E${materialEnd}`).format.columnWidth = 18;
  for (const column of ["J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AK"]) resultSheet.getRange(`${column}1:${column}${materialEnd}`).format.columnWidth = 14;
  resultSheet.getRange(`AJ1:AJ${materialEnd}`).format.columnWidth = 36;
  resultSheet.getRange(`AJ2:AJ${materialEnd}`).format.wrapText = true;
  applyStatusFormatting(resultSheet, `AI2:AI${materialEnd}`);

  continuousDetailSheet.showGridLines = false;
  continuousDetailSheet.freezePanes.freezeRows(1);
  continuousDetailSheet.freezePanes.freezeColumns(4);
  continuousDetailSheet.getRange("A1:AK1").values = [normalizedHeaders];
  continuousDetailSheet.getRange("J1:L1").values = [["连续期初数量", "连续期初单价", "连续期初金额"]];
  continuousDetailSheet.getRange("AD1:AD1").values = [["连续CAATS结存金额"]];
  continuousDetailSheet.getRange("AF1:AK1").values = [["连续月末平均价", "连续CAATS结存数量", "结存数量差异", "连续结果状态", "连续滚算说明", "连续发出差异绝对值"]];
  continuousDetailSheet.getRange("AL1:AO1").values = [["U8期初数量(核对)", "U8期初金额(核对)", "连续起算方式", "连续结存差异绝对值"]];
  const continuousMaterial = [...material].sort((left, right) => {
    const codeOrder = String(left.code).localeCompare(String(right.code), "zh-CN", { numeric: true });
    return codeOrder || String(left.month).localeCompare(String(right.month));
  });
  const continuousDetailValues = continuousMaterial.map((row) => [
    row.code, row.month, row.stock_code || null, row.name, row.spec, row.unit, row.weight || null, row.reg || null, row.manufacturer || null,
    null, null, null, row["收入数量"], null, null, row["收入单价"], row["收入金额"], null, null,
    row["发出数量"], null, null, row["发出单价"], row["发出金额"], null, null,
    row["结存数量"], row["结存单价"], row["结存金额"], null, null,
    null, null, null, null, null, null,
    row["期初数量"], row["期初金额"], null, null,
  ]);
  continuousDetailSheet.getRange(`A2:AO${materialEnd}`).values = continuousDetailValues;
  const continuousDetailFormulas = { J: [], K: [], L: [], N: [], O: [], R: [], S: [], U: [], V: [], Y: [], Z: [], AD: [], AE: [], AF: [], AG: [], AH: [], AI: [], AJ: [], AK: [], AN: [], AO: [] };
  for (let row = 2; row <= materialEnd; row += 1) {
    const canRoll = row === 2 ? null : `A${row}=A${row - 1}`;
    continuousDetailFormulas.J.push([row === 2 ? `=AL${row}` : `=IF(${canRoll},AG${row - 1},AL${row})`]);
    continuousDetailFormulas.K.push([`=IF(J${row}=0,0,L${row}/J${row})`]);
    continuousDetailFormulas.L.push([row === 2 ? `=AM${row}` : `=IF(${canRoll},AD${row - 1},AM${row})`]);
    continuousDetailFormulas.N.push([`=M${row}`]);
    continuousDetailFormulas.O.push([`=M${row}-N${row}`]);
    continuousDetailFormulas.R.push([`=Q${row}`]);
    continuousDetailFormulas.S.push([`=Q${row}-R${row}`]);
    continuousDetailFormulas.U.push([`=T${row}`]);
    continuousDetailFormulas.V.push([`=T${row}-U${row}`]);
    continuousDetailFormulas.Y.push([`=U${row}*AF${row}`]);
    continuousDetailFormulas.Z.push([`=X${row}-Y${row}`]);
    continuousDetailFormulas.AD.push([`=L${row}+R${row}-Y${row}`]);
    continuousDetailFormulas.AE.push([`=AC${row}-AD${row}`]);
    continuousDetailFormulas.AF.push([`=IF(J${row}+N${row}=0,0,(L${row}+R${row})/(J${row}+N${row}))`]);
    continuousDetailFormulas.AG.push([`=J${row}+N${row}-U${row}`]);
    continuousDetailFormulas.AH.push([`=AA${row}-AG${row}`]);
    continuousDetailFormulas.AI.push([`=IF(ABS(Z${row})<=${amountTolerance},"PASS","REVIEW")`]);
    continuousDetailFormulas.AJ.push([`=IF(AN${row}="使用U8期初","该物料首次出现，使用U8期初","承接最近一次可追溯连续CAATS期末")`]);
    continuousDetailFormulas.AK.push([`=ABS(Z${row})`]);
    continuousDetailFormulas.AN.push([row === 2 ? `="使用U8期初"` : `=IF(${canRoll},"承接最近可追溯连续期末","使用U8期初")`]);
    continuousDetailFormulas.AO.push([`=ABS(AE${row})`]);
  }
  for (const [column, values] of Object.entries(continuousDetailFormulas)) {
    continuousDetailSheet.getRange(`${column}2:${column}${materialEnd}`).formulas = values;
  }
  applyTableGrid(continuousDetailSheet, `A1:AO${materialEnd}`);
  continuousDetailSheet.getRange("A1:I1").format = headerFormat(palette.title);
  continuousDetailSheet.getRange("J1:L1").format = headerFormat(palette.sectionHeader);
  continuousDetailSheet.getRange("M1:S1").format = headerFormat(palette.incomeHeader);
  continuousDetailSheet.getRange("T1:Z1").format = headerFormat(palette.issueHeader);
  continuousDetailSheet.getRange("AA1:AE1").format = headerFormat(palette.closingHeader);
  continuousDetailSheet.getRange("AF1:AK1").format = headerFormat(palette.resultHeader);
  continuousDetailSheet.getRange("AL1:AN1").format = headerFormat(palette.sectionHeader);
  continuousDetailSheet.getRange("AO1:AO1").format = headerFormat(palette.resultHeader);
  continuousDetailSheet.getRange("A1:AO1").format.rowHeight = 48;
  continuousDetailSheet.getRange(`J2:AM${materialEnd}`).format.numberFormat = numberFormat;
  continuousDetailSheet.getRange(`AO2:AO${materialEnd}`).format.numberFormat = numberFormat;
  continuousDetailSheet.getRange(`A2:A${materialEnd}`).format.numberFormat = "0000";
  continuousDetailSheet.getRange(`B2:B${materialEnd}`).format.numberFormat = "0";
  for (const column of ["A", "B", "C", "F", "G", "I"]) continuousDetailSheet.getRange(`${column}1:${column}${materialEnd}`).format.columnWidth = 11;
  for (const column of ["D", "H"]) continuousDetailSheet.getRange(`${column}1:${column}${materialEnd}`).format.columnWidth = 24;
  continuousDetailSheet.getRange(`E1:E${materialEnd}`).format.columnWidth = 18;
  for (const column of ["J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AK"]) continuousDetailSheet.getRange(`${column}1:${column}${materialEnd}`).format.columnWidth = 14;
  continuousDetailSheet.getRange(`AJ1:AJ${materialEnd}`).format.columnWidth = 36;
  continuousDetailSheet.getRange(`AL1:AM${materialEnd}`).format.columnWidth = 18;
  continuousDetailSheet.getRange(`AN1:AN${materialEnd}`).format.columnWidth = 22;
  continuousDetailSheet.getRange(`AO1:AO${materialEnd}`).format.columnWidth = 18;
  continuousDetailSheet.getRange(`AJ2:AN${materialEnd}`).format.wrapText = true;
  applyStatusFormatting(continuousDetailSheet, `AI2:AI${materialEnd}`);

  summarySheet.showGridLines = false;
  summarySheet.getRange("A1:J1").merge();
  summarySheet.getRange("A1").values = [["U8存货发出计价——物料月维度结果汇总"]];
  summarySheet.getRange("A1:J1").format = titleFormat;
  summarySheet.getRange("A3:B10").values = [
    ["指标", "结果"],
    ["物料月份行数", null],
    ["PASS行数", null],
    ["REVIEW行数", null],
    ["发出金额差异净额", null],
    ["发出金额差异绝对值", null],
    ["总体正负抵销金额", data.metrics.global_netting],
    ["同物料月跨仓抵销金额", data.offset_summary.total_offset],
  ];
  summarySheet.getRange("A3:B3").format = headerFormat(palette.sectionHeader);
  summarySheet.getRange("B4:B9").formulas = [
    [`=COUNTA('${resultSheetName}'!A2:A${materialEnd})`],
    [`=COUNTIF('${resultSheetName}'!AI2:AI${materialEnd},"PASS")`],
    [`=COUNTIF('${resultSheetName}'!AI2:AI${materialEnd},"REVIEW")`],
    [`=SUM('${resultSheetName}'!Z2:Z${materialEnd})`],
    [`=SUM('${resultSheetName}'!AK2:AK${materialEnd})`],
    [`=MIN(SUMIF('${resultSheetName}'!Z2:Z${materialEnd},">0",'${resultSheetName}'!Z2:Z${materialEnd}),-SUMIF('${resultSheetName}'!Z2:Z${materialEnd},"<0",'${resultSheetName}'!Z2:Z${materialEnd}))`],
  ];
  summarySheet.getRange("D3:G3").values = [["结论项目", "结果/口径", "状态", "审阅提示"]];
  summarySheet.getRange("D3:G3").format = headerFormat(palette.tableHeader);
  summarySheet.getRange("D4:G9").values = [
    ["月度独立口径", "物料+月份", "已执行", "每月使用客户U8期初，不累计上月CAATS差异"],
    ["连续滚算", `6月期末差异净额${data.continuous_monthly.at(-1).continuous_end_difference.toFixed(2)}元`, "敏感性分析", "未模拟自动调整单及最终取价调整"],
    ["跨仓抵销", `${data.offset_summary.both_sign_keys}个物料月份存在仓库正负贡献`, "已识别", "仓库贡献相对物料统一价，不代表仓库错账"],
    ["客户实际配置", "普通单据按仓；特殊单据按仓库组", "配置一致", "物料统一价仅为审计模拟"],
    ["差异方向", "U8发出金额-CAATS发出金额", "统一", "净额存在正负抵销，调整应看明细"],
    ["支持证据", "另见ITA核对钩稽数据工作簿", "职责分离", "CAATS表不保存流水及配置底表"],
  ];
  summarySheet.getRange("A12:J12").values = [["年月", "物料月份数", "REVIEW数", "PASS数", "U8发出金额", "CAATS发出金额", "发出差异净额", "发出差异绝对值", "REVIEW占比", "说明"]];
  summarySheet.getRange("A12:J12").format = headerFormat(palette.tableHeader);
  for (let index = 0; index < data.monthly.length; index += 1) {
    const row = 13 + index;
    const month = data.monthly[index].month;
    summarySheet.getRange(`A${row}`).values = [[month]];
    summarySheet.getRange(`B${row}:I${row}`).formulas = [[
      `=COUNTIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row})`,
      `=COUNTIFS('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AI$2:$AI$${materialEnd},"REVIEW")`,
      `=COUNTIFS('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AI$2:$AI$${materialEnd},"PASS")`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$X$2:$X$${materialEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$Y$2:$Y$${materialEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$Z$2:$Z$${materialEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AK$2:$AK$${materialEnd})`,
      `=IF(B${row}=0,0,C${row}/B${row})`,
    ]];
    summarySheet.getRange(`J${row}`).values = [["月度独立测试；差异方向为U8-CAATS"]];
  }
  summarySheet.getRange("B4:B10").format.numberFormat = numberFormat;
  summarySheet.getRange("E13:H18").format.numberFormat = numberFormat;
  summarySheet.getRange("I13:I18").format.numberFormat = "0.0%";
  summarySheet.getRange("A3:B10").format.borders = thinBorder;
  summarySheet.getRange("D3:G9").format.borders = thinBorder;
  summarySheet.getRange("A12:J18").format.borders = thinBorder;
  summarySheet.getRange("A1:A22").format.columnWidth = 22;
  summarySheet.getRange("B1:C22").format.columnWidth = 16;
  summarySheet.getRange("D1:D22").format.columnWidth = 28;
  summarySheet.getRange("E1:F22").format.columnWidth = 25;
  summarySheet.getRange("G1:G22").format.columnWidth = 44;
  summarySheet.getRange("H1:J22").format.columnWidth = 18;
  summarySheet.getRange("D4:G9").format.wrapText = true;

  differenceSheet.showGridLines = false;
  differenceSheet.freezePanes.freezeRows(1);
  differenceSheet.freezePanes.freezeColumns(2);
  const differenceHeaders = ["年月", "存货编码", "存货名称", "规格", "U8发出金额", "CAATS发出金额", "发出金额差异", "结果状态", "主要差异原因"];
  differenceSheet.getRange("A1:I1").values = [differenceHeaders];
  differenceSheet.getRange("A1:I1").format = headerFormat(palette.issueHeader);
  const differenceValues = data.difference_rows.map((row) => [row.month, row.code, row.name, row.spec, row["发出金额"], row.caats_issue, row.issue_diff, row.status, row.reason]);
  differenceSheet.getRange(`A2:I${differenceValues.length + 1}`).values = differenceValues;
  differenceSheet.getRange(`E2:G${differenceValues.length + 1}`).format.numberFormat = numberFormat;
  differenceSheet.getRange(`B2:B${differenceValues.length + 1}`).format.numberFormat = "0000";
  applyTableGrid(differenceSheet, `A1:I${differenceValues.length + 1}`);
  differenceSheet.getRange("A1:I1").format = headerFormat(palette.issueHeader);
  for (const column of ["A", "B", "D", "H"]) differenceSheet.getRange(`${column}1:${column}${differenceValues.length + 1}`).format.columnWidth = 15;
  differenceSheet.getRange(`C1:C${differenceValues.length + 1}`).format.columnWidth = 28;
  for (const column of ["E", "F", "G"]) differenceSheet.getRange(`${column}1:${column}${differenceValues.length + 1}`).format.columnWidth = 18;
  differenceSheet.getRange(`I1:I${differenceValues.length + 1}`).format.columnWidth = 40;
  applyStatusFormatting(differenceSheet, `H2:H${differenceValues.length + 1}`);

  continuousSummarySheet.showGridLines = false;
  continuousSummarySheet.getRange("A1:L1").merge();
  continuousSummarySheet.getRange("A1").values = [["连续滚算汇总（敏感性分析）"]];
  continuousSummarySheet.getRange("A1:L1").format = titleFormat;
  continuousSummarySheet.getRange("A3:B7").values = [
    ["指标", "结果"],
    ["月间期初期末不一致数", data.transition_checks.reduce((sum, row) => sum + row.quantity_mismatch + row.amount_mismatch, 0)],
    ["6月连续期末差异净额", data.continuous_monthly.at(-1).continuous_end_difference],
    ["6月连续期末差异绝对值", data.continuous_monthly.at(-1).continuous_end_abs],
    ["计算限制", "未模拟自动调整单、成本卷积调整及最终取价调整"],
  ];
  continuousSummarySheet.getRange("A3:B3").format = headerFormat(palette.sectionHeader);
  continuousSummarySheet.getRange("A3:B7").format.borders = thinBorder;
  continuousSummarySheet.getRange("B5:B6").format.numberFormat = numberFormat;
  continuousSummarySheet.getRange("D3:L3").values = [["解释", "首次出现", "后续期初", "当月收入", "当月发出数量", "当月差异", "期末差异", "是否调整金额", "用途"]];
  continuousSummarySheet.getRange("D3:L3").format = headerFormat(palette.tableHeader);
  continuousSummarySheet.getRange("D4:L4").values = [["连续滚算规则", "该物料首次出现时使用U8期初", "使用最近一次可追溯连续CAATS期末", "使用U8月表", "使用U8月表", "U8-连续CAATS", "U8期末-连续CAATS期末", "否", "观察差异传导，不作为直接调整"]];
  continuousSummarySheet.getRange("D3:L4").format.borders = thinBorder;
  continuousSummarySheet.getRange("D4:L4").format.wrapText = true;
  const continuousHeaders = ["年月", "物料月份数", "U8发出金额", "月度独立CAATS发出", "月度独立差异", "月度独立绝对差异", "连续CAATS发出", "连续当月差异", "连续当月绝对差异", "连续期末差异净额", "连续期末差异绝对值", "上月差异传导影响"];
  continuousSummarySheet.getRange("A10:L10").values = [continuousHeaders];
  continuousSummarySheet.getRange("A10:L10").format = headerFormat(palette.tableHeader);
  for (let index = 0; index < data.monthly.length; index += 1) {
    const row = 11 + index;
    const month = data.monthly[index].month;
    continuousSummarySheet.getRange(`A${row}`).values = [[month]];
    continuousSummarySheet.getRange(`B${row}:L${row}`).formulas = [[
      `=COUNTIF('${continuousDetailName}'!$B$2:$B$${materialEnd},A${row})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${materialEnd},A${row},'${continuousDetailName}'!$X$2:$X$${materialEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$Y$2:$Y$${materialEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$Z$2:$Z$${materialEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AK$2:$AK$${materialEnd})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${materialEnd},A${row},'${continuousDetailName}'!$Y$2:$Y$${materialEnd})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${materialEnd},A${row},'${continuousDetailName}'!$Z$2:$Z$${materialEnd})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${materialEnd},A${row},'${continuousDetailName}'!$AK$2:$AK$${materialEnd})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${materialEnd},A${row},'${continuousDetailName}'!$AE$2:$AE$${materialEnd})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${materialEnd},A${row},'${continuousDetailName}'!$AO$2:$AO$${materialEnd})`,
      `=H${row}-E${row}`,
    ]];
  }
  const continuousTotalRow = data.monthly.length + 11;
  continuousSummarySheet.getRange(`A${continuousTotalRow}:L${continuousTotalRow}`).values = [["合计/期末", null, null, null, null, null, null, null, null, null, null, null]];
  continuousSummarySheet.getRange(`B${continuousTotalRow}:I${continuousTotalRow}`).formulas = [[
    `=SUM(B11:B${continuousTotalRow - 1})`, `=SUM(C11:C${continuousTotalRow - 1})`, `=SUM(D11:D${continuousTotalRow - 1})`, `=SUM(E11:E${continuousTotalRow - 1})`, `=SUM(F11:F${continuousTotalRow - 1})`, `=SUM(G11:G${continuousTotalRow - 1})`, `=SUM(H11:H${continuousTotalRow - 1})`, `=SUM(I11:I${continuousTotalRow - 1})`,
  ]];
  continuousSummarySheet.getRange(`J${continuousTotalRow}:K${continuousTotalRow}`).formulas = [[`=J${continuousTotalRow - 1}`, `=K${continuousTotalRow - 1}`]];
  continuousSummarySheet.getRange(`L${continuousTotalRow}`).formulas = [[`=SUM(L11:L${continuousTotalRow - 1})`]];
  continuousSummarySheet.getRange("B5:B6").formulas = [[`=J${continuousTotalRow - 1}`], [`=K${continuousTotalRow - 1}`]];
  continuousSummarySheet.getRange(`A10:L${continuousTotalRow}`).format.borders = thinBorder;
  continuousSummarySheet.getRange(`C11:L${continuousTotalRow}`).format.numberFormat = numberFormat;
  continuousSummarySheet.getRange(`A${continuousTotalRow}:L${continuousTotalRow}`).format.font = { bold: true };
  continuousSummarySheet.getRange("A1:A22").format.columnWidth = 20;
  continuousSummarySheet.getRange("B1:B22").format.columnWidth = 18;
  continuousSummarySheet.getRange("C1:L22").format.columnWidth = 20;

  offsetSheet.showGridLines = false;
  offsetSheet.freezePanes.freezeRows(9);
  offsetSheet.freezePanes.freezeColumns(4);
  offsetSheet.getRange("A1:Q1").merge();
  offsetSheet.getRange("A1").values = [["跨仓抵销明细（相对物料统一CAATS平均价）"]];
  offsetSheet.getRange("A1:Q1").format = titleFormat;
  offsetSheet.getRange("A3:B7").values = [
    ["指标", "结果"],
    ["多仓发出物料月份", data.offset_summary.multi_issue_keys],
    ["数量金额干净桥接组合", data.offset_summary.clean_multi_keys],
    ["存在正负仓库贡献组合", data.offset_summary.both_sign_keys],
    ["同物料月份跨仓抵销金额", data.offset_summary.total_offset],
  ];
  offsetSheet.getRange("A3:B3").format = headerFormat(palette.sectionHeader);
  offsetSheet.getRange("A3:B7").format.borders = thinBorder;
  offsetSheet.getRange("B7").format.numberFormat = numberFormat;
  offsetSheet.getRange("D3:G6").values = [
    ["计算项目", "公式/条件", "解释", "结论边界"],
    ["仓库贡献", "仓库U8发出金额-仓库发出数量×物料统一平均价", "正负表示相对统一价的方向", "不是仓库错账"],
    ["纳入条件", "同物料月至少2个出库仓且数量、金额桥接通过", "排除流水金额缺口干扰", "调整单缺失组合不纳入"],
    ["抵销金额", "MIN(正贡献合计,负贡献绝对值合计)", "展示总体物料结果掩盖的仓库差异", "不能直接记账"],
  ];
  offsetSheet.getRange("D3:G3").format = headerFormat(palette.tableHeader);
  offsetSheet.getRange("D3:G6").format.borders = thinBorder;
  offsetSheet.getRange("D4:G6").format.wrapText = true;
  const offsetHeaders = ["年月", "存货编码", "存货名称", "规格", "物料统一平均价", "物料发出差异", "仓库编码", "仓库名称", "仓库核算组", "仓库发出数量", "仓库发出金额", "统一价预期金额", "仓库贡献", "同物料月正贡献", "同物料月负贡献绝对值", "同物料月抵销金额", "同物料月净额"];
  offsetSheet.getRange("A9:Q9").values = [offsetHeaders];
  offsetSheet.getRange("A9:Q9").format = headerFormat(palette.issueHeader);
  const offsetValues = data.offset_rows.map((row) => [row.month, row.code, row.name, row.spec, row.pooled_avg, row.material_issue_difference, row.warehouse_code, row.warehouse, row.group, row.warehouse_issue_quantity, row.warehouse_issue_amount, row.pooled_expected_amount, row.warehouse_contribution, row.key_positive, row.key_negative_abs, row.key_offset, row.key_net]);
  if (offsetValues.length) {
    offsetSheet.getRange(`A10:Q${offsetValues.length + 9}`).values = offsetValues;
    offsetSheet.getRange(`E10:Q${offsetValues.length + 9}`).format.numberFormat = numberFormat;
    offsetSheet.getRange(`B10:B${offsetValues.length + 9}`).format.numberFormat = "0000";
    offsetSheet.getRange(`G10:G${offsetValues.length + 9}`).format.numberFormat = "00";
    offsetSheet.getRange(`I10:I${offsetValues.length + 9}`).format.numberFormat = "00";
    applyTableGrid(offsetSheet, `A9:Q${offsetValues.length + 9}`);
    offsetSheet.getRange("A9:Q9").format = headerFormat(palette.issueHeader);
  }
  for (const column of ["A", "B", "D", "G", "I"]) offsetSheet.getRange(`${column}1:${column}${Math.max(20, offsetValues.length + 9)}`).format.columnWidth = 15;
  offsetSheet.getRange(`C1:C${Math.max(20, offsetValues.length + 9)}`).format.columnWidth = 28;
  offsetSheet.getRange(`H1:H${Math.max(20, offsetValues.length + 9)}`).format.columnWidth = 36;
  for (const column of ["E", "F", "J", "K", "L", "M", "N", "O", "P", "Q"]) offsetSheet.getRange(`${column}1:${column}${Math.max(20, offsetValues.length + 9)}`).format.columnWidth = 18;

  await saveInspection(workbook, "caats", [["02_物料月维度结果汇总", "A1:J18"], ["03_连续滚算明细", "A1:AO20"], ["04_连续滚算汇总", "A1:L17"], ["05_存在差异明细", "A1:I20"]]);
  await renderSheets(workbook, "caats", [
    ["01_物料月维度明细", "A1:AK24", 0.7],
    ["02_物料月维度结果汇总", "A1:J18", 1.0],
    ["03_连续滚算明细", "A1:AO24", 0.66],
    ["04_连续滚算汇总", "A1:L17", 0.95],
    ["05_存在差异明细", "A1:I24", 0.95],
    ["06_跨仓抵销", "A1:Q22", 0.8],
  ]);
  const output = await SpreadsheetFile.exportXlsx(workbook);
  const path = `${outputDir}/U8存货发出计价_CAATS审计结果表.xlsx`;
  await output.save(path);
  return path;
}

async function buildItaWorkbook() {
  const workbook = Workbook.create();
  const summarySheet = workbook.worksheets.add("01_核对摘要");
  const inputSheet = workbook.worksheets.add("02_输入清单");
  const transitionSheet = workbook.worksheets.add("03_月间衔接");
  const bridgeSheet = workbook.worksheets.add("04_流水桥接");
  const warehouseSheet = workbook.worksheets.add("05_仓库物料月");
  const materialBridgeSheet = workbook.worksheets.add("06_物料月桥接");
  const groupSheet = workbook.worksheets.add("07_仓库组特殊单据");
  const masterSheet = workbook.worksheets.add("08_仓库档案");
  const scopeSheet = workbook.worksheets.add("09_口径说明");
  const checkSheet = workbook.worksheets.add("10_检查");

  summarySheet.showGridLines = false;
  summarySheet.getRange("A1:H1").merge();
  summarySheet.getRange("A1").values = [["U8存货发出计价——ITA核对钩稽摘要"]];
  summarySheet.getRange("A1:H1").format = titleFormat;
  summarySheet.getRange("A3:B11").values = [
    ["指标", "结果"],
    ["输入文件数", data.source_manifest.length],
    ["物料月份行数", material.length],
    ["仓库物料月份行数", warehouses.length],
    ["仓库组特殊单据组合", groups.length],
    ["仓库档案数", master.length],
    ["已记账成本流水", data.metrics.posted_cost_rows],
    ["未记账流水", data.metrics.unposted_rows],
    ["月间期初期末不一致", data.transition_checks.reduce((sum, row) => sum + row.quantity_mismatch + row.amount_mismatch, 0)],
  ];
  summarySheet.getRange("A3:B3").format = headerFormat(palette.sectionHeader);
  summarySheet.getRange("A3:B11").format.borders = thinBorder;
  summarySheet.getRange("D3:H3").values = [["钩稽层级", "来源", "核对对象", "状态", "说明"]];
  summarySheet.getRange("D3:H3").format = headerFormat(palette.tableHeader);
  summarySheet.getRange("D4:H9").values = [
    ["月间衔接", "各月收发存", "上月期末=下月期初", data.transition_checks.every((row) => row.status === "PASS") ? "PASS" : "REVIEW", "数量和金额分别检查"],
    ["流水范围", "出入库流水+仓库档案", "记账人非空且记入成本=是", data.metrics.missing_master.length ? "REVIEW" : "PASS", "四类借调拨移动不按名称整类剔除"],
    ["物料月桥接", "流水vs收发存", "收入、发出数量", "PASS", "金额缺口单独披露"],
    ["仓库维度", "流水+仓库档案", "仓库+物料+月份", "待补资料", "缺分仓期初数量和金额"],
    ["特殊单据", "其他出入库单", "仓库组+物料+月份", "待补资料", "需成本卷积结果"],
    ["CAATS结果", "另表", "差异、连续滚算、跨仓抵销", "已分离", "本表保存支持性证据"],
  ];
  summarySheet.getRange("D3:H9").format.borders = thinBorder;
  summarySheet.getRange("A1:A14").format.columnWidth = 28;
  summarySheet.getRange("B1:B14").format.columnWidth = 25;
  summarySheet.getRange("D1:H14").format.columnWidth = 31;
  summarySheet.getRange("D4:H9").format.wrapText = true;

  inputSheet.showGridLines = false;
  inputSheet.freezePanes.freezeRows(1);
  inputSheet.getRange("A1:F1").values = [["角色", "文件名", "完整路径", "文件大小(Byte)", "SHA-256", "用途"]];
  inputSheet.getRange("A1:F1").format = headerFormat(palette.tableHeader);
  const inputValues = data.source_manifest.map((row) => [row.role, row.name, row.path, row.size, row.sha256, row.name.includes("流水") ? "出入库发生明细" : row.name.includes("收发存") ? "物料月期初/收发/期末" : row.name.includes("仓库") ? "仓库配置" : "CAATS展示模板"]);
  inputSheet.getRange(`A2:F${inputValues.length + 1}`).values = inputValues;
  inputSheet.getRange(`A1:F${inputValues.length + 1}`).format.borders = thinBorder;
  inputSheet.getRange("A1:A20").format.columnWidth = 16;
  inputSheet.getRange("B1:B20").format.columnWidth = 40;
  inputSheet.getRange("C1:C20").format.columnWidth = 72;
  inputSheet.getRange("D1:D20").format.columnWidth = 20;
  inputSheet.getRange("E1:E20").format.columnWidth = 68;
  inputSheet.getRange("F1:F20").format.columnWidth = 30;

  transitionSheet.showGridLines = false;
  transitionSheet.getRange("A1:J1").values = [["上月", "本月", "比较物料数", "数量不一致数", "金额不一致数", "数量差异绝对值", "金额差异绝对值", "数量容差", "金额容差", "状态"]];
  transitionSheet.getRange("A1:J1").format = headerFormat(palette.tableHeader);
  const transitionValues = data.transition_checks.map((row) => [row.previous_month, row.month, row.compared, row.quantity_mismatch, row.amount_mismatch, row.quantity_abs, row.amount_abs, quantityTolerance, amountTolerance, row.status]);
  transitionSheet.getRange(`A2:J${transitionValues.length + 1}`).values = transitionValues;
  transitionSheet.getRange(`F2:I${transitionValues.length + 1}`).format.numberFormat = numberFormat;
  transitionSheet.getRange(`A1:J${transitionValues.length + 1}`).format.borders = thinBorder;
  transitionSheet.getRange("A1:J12").format.columnWidth = 19;
  applyStatusFormatting(transitionSheet, `J2:J${transitionValues.length + 1}`);

  bridgeSheet.showGridLines = false;
  bridgeSheet.freezePanes.freezeRows(1);
  bridgeSheet.freezePanes.freezeColumns(3);
  const bridgeHeaders = ["年月", "存货编码", "存货名称", "U8收入数量", "流水已记账收入数量", "收入数量差异", "U8收入金额", "流水已记账收入金额", "收入金额缺口", "U8发出数量", "流水已记账发出数量", "发出数量差异", "U8发出金额", "流水已记账发出金额", "发出金额缺口", "调拨入库数量", "调拨入库金额", "调拨出库数量", "调拨出库金额", "桥接状态", "说明"];
  bridgeSheet.getRange("A1:U1").values = [bridgeHeaders];
  bridgeSheet.getRange("A1:U1").format = headerFormat(palette.tableHeader);
  bridgeSheet.getRange("A1:U1").format.rowHeight = 48;
  const bridgeValues = material.map((row) => [
    row.month, row.code, row.name, row["收入数量"], row.ledger_iq, row["收入数量"] - row.ledger_iq,
    row["收入金额"], row.ledger_ia, row.income_amount_gap, row["发出数量"], row.ledger_oq, row["发出数量"] - row.ledger_oq,
    row["发出金额"], row.ledger_oa, row.issue_amount_gap, row.transfer_iq, row.transfer_ia, row.transfer_oq, row.transfer_oa,
    Math.abs(row["收入数量"] - row.ledger_iq) <= quantityTolerance && Math.abs(row["发出数量"] - row.ledger_oq) <= quantityTolerance ? "PASS" : "REVIEW",
    Math.abs(row.income_amount_gap) > amountTolerance || Math.abs(row.issue_amount_gap) > amountTolerance ? "数量已对平；金额缺口需调整单/最终取价明细" : "数量金额均与流水一致",
  ]);
  bridgeSheet.getRange(`A2:U${bridgeValues.length + 1}`).values = bridgeValues;
  bridgeSheet.getRange(`D2:S${bridgeValues.length + 1}`).format.numberFormat = numberFormat;
  bridgeSheet.getRange(`B2:B${bridgeValues.length + 1}`).format.numberFormat = "0000";
  applyTableGrid(bridgeSheet, `A1:U${bridgeValues.length + 1}`);
  bridgeSheet.getRange("A1:U1").format = headerFormat(palette.tableHeader);
  for (const column of ["A", "B", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"]) bridgeSheet.getRange(`${column}1:${column}${bridgeValues.length + 1}`).format.columnWidth = 15;
  bridgeSheet.getRange(`C1:C${bridgeValues.length + 1}`).format.columnWidth = 28;
  bridgeSheet.getRange(`U1:U${bridgeValues.length + 1}`).format.columnWidth = 44;
  applyStatusFormatting(bridgeSheet, `T2:T${bridgeValues.length + 1}`);

  warehouseSheet.showGridLines = false;
  warehouseSheet.freezePanes.freezeRows(3);
  warehouseSheet.freezePanes.freezeColumns(4);
  warehouseSheet.getRange("A1:AX1").merge();
  warehouseSheet.getRange("A1").values = [["仓库+物料+月份核对表（ITA全量）"]];
  warehouseSheet.getRange("A1:AX1").format = titleFormat;
  const warehouseHeaders = ["年月", "仓库编码", "仓库名称", "仓库核算组", "计价方式", "记入成本", "停用日期", "存货编码", "存货名称", "规格", "单位", "已记账成本行数", "未记账行数", "已记账非成本行数", "已记账入库数量", "已记账入库金额", "正常单据入库数量", "正常单据入库金额", "特殊单据入库数量", "特殊单据入库金额", "已记账出库数量", "已记账出库金额", "观察出库单价", "正常单据出库数量", "正常单据出库金额", "特殊单据出库数量", "特殊单据出库金额", "调拨入库数量", "调拨入库金额", "调拨出库数量", "调拨出库金额", "涉及收发类别", "涉及单据类型", "U8物料月收入数量", "分仓汇总收入数量", "收入数量桥接", "U8物料月发出数量", "分仓汇总发出数量", "发出数量桥接", "分仓期初数量(输入)", "分仓期初金额(输入)", "分仓重算平均价", "正常单据重算出库金额", "特殊单据已记账金额(占位)", "重算发出金额", "发出金额差异", "分仓期末数量", "分仓期末金额", "状态", "说明"];
  warehouseSheet.getRange("A3:AX3").values = [warehouseHeaders];
  warehouseSheet.getRange("A3:AX3").format = headerFormat(palette.tableHeader);
  warehouseSheet.getRange("A3:AX3").format.rowHeight = 60;
  const warehouseValues = warehouses.map((row) => [
    row.month, row.warehouse_code, row.warehouse, row.group, row.method, row.costed, row.stopped || null, row.code, row.name, row.spec, row.unit,
    row.posted_rows, row.unposted_rows, row.posted_no_cost_rows, row.iq, row.ia, row.normal_iq, row.normal_ia, row.special_iq, row.special_ia,
    row.oq, row.oa, row.out_unit, row.normal_oq, row.normal_oa, row.special_oq, row.special_oa, row.transfer_iq, row.transfer_ia, row.transfer_oq, row.transfer_oa,
    row.categories, row.doc_types, row.u8_mat_iq, row.mat_sum_iq, null, row.u8_mat_oq, row.mat_sum_oq, null, null, null, null, null, row.special_oa, null, null, null, null, null, null,
  ]);
  warehouseSheet.getRange(`A4:AX${warehouseEnd}`).values = warehouseValues;
  const warehouseFormulas = { AJ: [], AM: [], AP: [], AQ: [], AS: [], AT: [], AU: [], AV: [], AW: [], AX: [] };
  for (let row = 4; row <= warehouseEnd; row += 1) {
    warehouseFormulas.AJ.push([`=AH${row}-AI${row}`]);
    warehouseFormulas.AM.push([`=AK${row}-AL${row}`]);
    warehouseFormulas.AP.push([`=IF(OR(AN${row}="",AO${row}="",AN${row}+O${row}=0),"",(AO${row}+P${row})/(AN${row}+O${row}))`]);
    warehouseFormulas.AQ.push([`=IF(AP${row}="","",X${row}*AP${row})`]);
    warehouseFormulas.AS.push([`=IF(AQ${row}="","",AQ${row}+AR${row})`]);
    warehouseFormulas.AT.push([`=IF(AS${row}="","",V${row}-AS${row})`]);
    warehouseFormulas.AU.push([`=IF(AN${row}="","",AN${row}+O${row}-U${row})`]);
    warehouseFormulas.AV.push([`=IF(AO${row}="","",AO${row}+P${row}-AS${row})`]);
    warehouseFormulas.AW.push([`=IF(F${row}="否","非成本仓",IF(AN${row}="","数量已桥接/待补期初",IF(ABS(AT${row})<=${amountTolerance},"PASS","REVIEW")))`]);
    warehouseFormulas.AX.push([`=IF(F${row}="否","不进入成本池",IF(AN${row}="","正常单据按仓、特殊单据按组；补充分仓期初后复算","特殊单据金额仍使用已记账金额占位，需仓库组成本结果替换"))`]);
  }
  for (const [column, values] of Object.entries(warehouseFormulas)) warehouseSheet.getRange(`${column}4:${column}${warehouseEnd}`).formulas = values;
  applyTableGrid(warehouseSheet, `A3:AX${warehouseEnd}`);
  warehouseSheet.getRange("A3:AX3").format = headerFormat(palette.tableHeader);
  warehouseSheet.getRange(`O4:AV${warehouseEnd}`).format.numberFormat = numberFormat;
  warehouseSheet.getRange(`B4:B${warehouseEnd}`).format.numberFormat = "00";
  warehouseSheet.getRange(`D4:D${warehouseEnd}`).format.numberFormat = "00";
  warehouseSheet.getRange(`H4:H${warehouseEnd}`).format.numberFormat = "0000";
  warehouseSheet.getRange(`AN4:AO${warehouseEnd}`).format = { fill: "#FFF2CC", font: { color: "#0000FF" } };
  for (const column of ["A", "B", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "AJ", "AM", "AW"]) warehouseSheet.getRange(`${column}1:${column}${warehouseEnd}`).format.columnWidth = 15;
  warehouseSheet.getRange(`C1:C${warehouseEnd}`).format.columnWidth = 34;
  warehouseSheet.getRange(`I1:I${warehouseEnd}`).format.columnWidth = 28;
  for (const column of ["O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AH", "AI", "AK", "AL", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV"]) warehouseSheet.getRange(`${column}1:${column}${warehouseEnd}`).format.columnWidth = 16;
  warehouseSheet.getRange(`AF1:AG${warehouseEnd}`).format.columnWidth = 34;
  warehouseSheet.getRange(`AX1:AX${warehouseEnd}`).format.columnWidth = 48;
  applyStatusFormatting(warehouseSheet, `AW4:AW${warehouseEnd}`);

  materialBridgeSheet.showGridLines = false;
  materialBridgeSheet.freezePanes.freezeRows(1);
  const materialBridgeHeaders = ["年月", "存货编码", "存货名称", "U8收入数量", "分仓已记账收入数量", "收入数量差异", "U8收入金额", "分仓已记账收入金额", "收入金额缺口", "U8发出数量", "分仓已记账发出数量", "发出数量差异", "U8发出金额", "分仓已记账发出金额", "发出金额缺口", "涉及成本仓数", "涉及出库仓数", "数量状态", "金额状态"];
  materialBridgeSheet.getRange("A1:S1").values = [materialBridgeHeaders];
  materialBridgeSheet.getRange("A1:S1").format = headerFormat(palette.tableHeader);
  const materialBridgeValues = material.map((row) => [row.month, row.code, row.name, row["收入数量"], row.ledger_iq, row["收入数量"] - row.ledger_iq, row["收入金额"], row.ledger_ia, row.income_amount_gap, row["发出数量"], row.ledger_oq, row["发出数量"] - row.ledger_oq, row["发出金额"], row.ledger_oa, row.issue_amount_gap, row.warehouse_count, row.issue_warehouse_count, Math.abs(row["收入数量"] - row.ledger_iq) <= quantityTolerance && Math.abs(row["发出数量"] - row.ledger_oq) <= quantityTolerance ? "PASS" : "REVIEW", Math.abs(row.income_amount_gap) <= amountTolerance && Math.abs(row.issue_amount_gap) <= amountTolerance ? "PASS" : "待调整单"]);
  materialBridgeSheet.getRange(`A2:S${materialBridgeValues.length + 1}`).values = materialBridgeValues;
  materialBridgeSheet.getRange(`D2:Q${materialBridgeValues.length + 1}`).format.numberFormat = numberFormat;
  materialBridgeSheet.getRange(`B2:B${materialBridgeValues.length + 1}`).format.numberFormat = "0000";
  applyTableGrid(materialBridgeSheet, `A1:S${materialBridgeValues.length + 1}`);
  materialBridgeSheet.getRange("A1:S1").format = headerFormat(palette.tableHeader);
  for (const column of ["A", "B", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S"]) materialBridgeSheet.getRange(`${column}1:${column}${materialBridgeValues.length + 1}`).format.columnWidth = 16;
  materialBridgeSheet.getRange(`C1:C${materialBridgeValues.length + 1}`).format.columnWidth = 28;
  applyStatusFormatting(materialBridgeSheet, `R2:R${materialBridgeValues.length + 1}`);

  groupSheet.showGridLines = false;
  groupSheet.freezePanes.freezeRows(1);
  const groupHeaders = ["年月", "仓库核算组", "存货编码", "存货名称", "规格", "单位", "特殊单据行数", "入库数量", "入库金额", "入库观察单价", "出库数量", "出库金额", "出库观察单价", "涉及收发类别", "配置说明"];
  groupSheet.getRange("A1:O1").values = [groupHeaders];
  groupSheet.getRange("A1:O1").format = headerFormat(palette.tableHeader);
  const groupValues = groups.map((row) => [row.month, row.group, row.code, row.name, row.spec, row.unit, row.rows, row.iq, row.ia, row.in_unit, row.oq, row.oa, row.out_unit, row.categories, "特殊单据按仓库核算组统一计价；当前为已记账观察结果"]);
  groupSheet.getRange(`A2:O${groupValues.length + 1}`).values = groupValues;
  groupSheet.getRange(`G2:M${groupValues.length + 1}`).format.numberFormat = numberFormat;
  groupSheet.getRange(`B2:B${groupValues.length + 1}`).format.numberFormat = "00";
  groupSheet.getRange(`C2:C${groupValues.length + 1}`).format.numberFormat = "0000";
  applyTableGrid(groupSheet, `A1:O${groupValues.length + 1}`);
  groupSheet.getRange("A1:O1").format = headerFormat(palette.tableHeader);
  for (const column of ["A", "B", "C", "E", "F", "G", "H", "I", "J", "K", "L", "M"]) groupSheet.getRange(`${column}1:${column}${groupValues.length + 1}`).format.columnWidth = 16;
  groupSheet.getRange(`D1:D${groupValues.length + 1}`).format.columnWidth = 28;
  groupSheet.getRange(`N1:N${groupValues.length + 1}`).format.columnWidth = 34;
  groupSheet.getRange(`O1:O${groupValues.length + 1}`).format.columnWidth = 48;

  masterSheet.showGridLines = false;
  masterSheet.freezePanes.freezeRows(1);
  const masterHeaders = ["仓库编码", "仓库名称", "部门名称", "计价方式", "仓库核算组", "记入成本", "纳入可用量计算", "停用日期", "仓库属性", "备注"];
  masterSheet.getRange("A1:J1").values = [masterHeaders];
  masterSheet.getRange("A1:J1").format = headerFormat(palette.tableHeader);
  const masterValues = master.map((row) => [row["仓库编码"], row["仓库名称"], row["部门名称"] || null, row["计价方式"], row["仓库核算组"] || null, row["记入成本"], row["纳入可用量计算"], row["停用日期"] || null, row["仓库属性"] || null, row["备注"] || null]);
  masterSheet.getRange(`A2:J${masterValues.length + 1}`).values = masterValues;
  masterSheet.getRange(`A2:A${masterValues.length + 1}`).format.numberFormat = "00";
  masterSheet.getRange(`E2:E${masterValues.length + 1}`).format.numberFormat = "00";
  applyTableGrid(masterSheet, `A1:J${masterValues.length + 1}`);
  masterSheet.getRange("A1:J1").format = headerFormat(palette.tableHeader);
  masterSheet.getRange("A1:A100").format.columnWidth = 14;
  masterSheet.getRange("B1:B100").format.columnWidth = 38;
  masterSheet.getRange("C1:J100").format.columnWidth = 20;

  scopeSheet.showGridLines = false;
  scopeSheet.getRange("A1:D1").merge();
  scopeSheet.getRange("A1").values = [["ITA核对钩稽及CAATS结果口径与限制"]];
  scopeSheet.getRange("A1:D1").format = titleFormat;
  scopeSheet.getRange("A3:D3").values = [["项目", "处理", "依据", "限制/后续"]];
  scopeSheet.getRange("A3:D3").format = headerFormat(palette.sectionHeader);
  const itaScopeRows = [
    ["纳入成本流水", "记账人非空且仓库记入成本=是", "流水记账状态+仓库档案", "未记账记录单独披露"],
    ["四类移动", "借用归还、借出借用、调拨出入库不按名称整类剔除", "以记账和成本属性判断", "特殊单据仍需配置清单"],
    ["月间衔接", "上月期末与下月期初逐物料比较", "客户月度收发存", "当前数量、金额均已衔接"],
    ["月度独立重算", "物料+月份；每月使用U8期初、收入和发出数量", "01_物料月维度明细", "不累计上月CAATS差异"],
    ["连续滚算", "该物料首次出现时使用U8期初；后续承接最近一次可追溯连续CAATS期末", "03_连续滚算明细", "仅观察差异传导，不作为直接调整"],
    ["连续滚算缺月", "同一物料中间缺月时不重置，继续承接最近一次存在数值的连续CAATS期末", "连续起算方式字段", "未模拟自动调整单及最终取价调整"],
    ["差异方向与容差", "U8金额-CAATS金额；正数表示U8金额较高", `金额容差${amountTolerance}元`, "浮点尾差按10位小数规范化后判断"],
    ["正常单据", "按仓库汇总", "仓库档案全月平均法", "缺分仓期初，暂不能独立复算"],
    ["特殊单据", "按仓库核算组汇总", "成本卷积配置", "需成本卷积结果"],
    ["跨仓抵销", "仅展示数量金额干净桥接组合", "仓库贡献相对物料统一价", "不是仓库错账或调整建议"],
    ["金额缺口", "物料月单独披露", "流水金额与U8汇总比较", "取得调整单/最终取价后更新"],
    ["CAATS/ITA边界", "CAATS仅保留六张结果表；口径及底层证据保存在ITA", "职责分离", "审阅结果时需同时索引本Sheet及ITA明细"],
  ];
  scopeSheet.getRange(`A4:D${itaScopeRows.length + 3}`).values = itaScopeRows;
  scopeSheet.getRange(`A3:D${itaScopeRows.length + 3}`).format.borders = thinBorder;
  scopeSheet.getRange("A1:A20").format.columnWidth = 25;
  scopeSheet.getRange("B1:D20").format.columnWidth = 52;
  scopeSheet.getRange("A4:D20").format.wrapText = true;

  checkSheet.showGridLines = false;
  checkSheet.getRange("A1:G1").merge();
  checkSheet.getRange("A1").values = [["ITA完整性与钩稽检查"]];
  checkSheet.getRange("A1:G1").format = titleFormat;
  checkSheet.getRange("A3:G3").values = [["检查", "实际值", "期望值", "差异", "容差", "状态", "说明"]];
  checkSheet.getRange("A3:G3").format = headerFormat(palette.sectionHeader);
  const incomeQuantityGapAbs = material.reduce((sum, row) => sum + Math.abs(row["收入数量"] - row.ledger_iq), 0);
  const issueQuantityGapAbs = material.reduce((sum, row) => sum + Math.abs(row["发出数量"] - row.ledger_oq), 0);
  const transitionMismatch = data.transition_checks.reduce((sum, row) => sum + row.quantity_mismatch + row.amount_mismatch, 0);
  const checkValues = [
    ["物料月份行数", material.length, data.metrics.material_month_rows, null, 0, null, "收发存去除合计行"],
    ["仓库物料月份行数", warehouses.length, data.metrics.warehouse_material_month_rows, null, 0, null, "流水发生额组合"],
    ["流水仓库映射缺失数", data.metrics.missing_master.length, 0, null, 0, null, "应全部映射仓库档案"],
    ["物料月收入数量桥接绝对差异", incomeQuantityGapAbs, 0, null, quantityTolerance, null, "逐物料月份绝对差异"],
    ["物料月发出数量桥接绝对差异", issueQuantityGapAbs, 0, null, quantityTolerance, null, "逐物料月份绝对差异"],
    ["月间期初期末不一致数", transitionMismatch, 0, null, 0, null, "数量+金额不一致项"],
    ["仓库档案非全月平均数", master.filter((row) => row["计价方式"] !== "全月平均法").length, 0, null, 0, null, "客户档案"],
    ["跨仓抵销明细行数", data.offset_rows.length, data.offset_summary.offset_rows, null, 0, null, "仅干净桥接组合"],
    ["缺分仓期初的可复算行数", warehouses.length, 0, null, 0, null, "已知限制，不代表数据错误"],
  ];
  checkSheet.getRange("A4:G12").values = checkValues;
  for (let row = 4; row <= 12; row += 1) {
    checkSheet.getRange(`D${row}`).formulas = [[`=B${row}-C${row}`]];
    checkSheet.getRange(`F${row}`).formulas = [[`=IF(ABS(D${row})<=E${row},"PASS","REVIEW")`]];
  }
  checkSheet.getRange("A3:G12").format.borders = thinBorder;
  checkSheet.getRange("B4:E12").format.numberFormat = numberFormat;
  checkSheet.getRange("A1:A15").format.columnWidth = 40;
  checkSheet.getRange("B1:F15").format.columnWidth = 16;
  checkSheet.getRange("G1:G15").format.columnWidth = 44;
  applyStatusFormatting(checkSheet, "F4:F12");

  await saveInspection(workbook, "ita", [["01_核对摘要", "A1:H11"], ["03_月间衔接", "A1:J6"], ["10_检查", "A1:G12"]]);
  await renderSheets(workbook, "ita", [
    ["01_核对摘要", "A1:H11", 1.05],
    ["02_输入清单", `A1:F${inputValues.length + 1}`, 0.9],
    ["03_月间衔接", "A1:J6", 1.0],
    ["04_流水桥接", "A1:U22", 0.72],
    ["05_仓库物料月", "A1:AX22", 0.63],
    ["06_物料月桥接", "A1:S22", 0.75],
    ["07_仓库组特殊单据", "A1:O22", 0.85],
    ["08_仓库档案", "A1:J22", 0.9],
    ["09_口径说明", "A1:D16", 1.0],
    ["10_检查", "A1:G12", 1.1],
  ]);
  const output = await SpreadsheetFile.exportXlsx(workbook);
  const path = `${outputDir}/U8存货发出计价_ITA核对钩稽数据.xlsx`;
  await output.save(path);
  return path;
}

let path;
if (buildPart === "caats") path = await buildCaatsWorkbook();
else if (buildPart === "ita") path = await buildItaWorkbook();
else throw new Error(`未知U8_BUILD_PART：${buildPart}`);

console.log(JSON.stringify({ buildPart, path, metrics: data.metrics, offsetSummary: data.offset_summary }, null, 2));
