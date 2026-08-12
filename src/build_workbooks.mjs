import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const projectRoot = process.env.U8_PROJECT_ROOT;
const outputDir = process.env.U8_OUTPUT_DIR;
const qaDir = process.env.U8_QA_DIR;
const dataPath = process.env.U8_DATA_JSON;
const buildPart = process.env.U8_BUILD_PART || "caats";
const skipQa = process.env.U8_SKIP_QA === "1";

if (!projectRoot || !outputDir || !qaDir || !dataPath) {
  throw new Error("缺少U8_PROJECT_ROOT/U8_OUTPUT_DIR/U8_QA_DIR/U8_DATA_JSON环境变量");
}

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(qaDir, { recursive: true });
const data = JSON.parse(await fs.readFile(dataPath, "utf8"));
const material = data.material_rows;
const continuousMaterial = data.continuous_material_rows;
const warehouses = data.warehouse_rows;
const groups = data.group_rows;
const master = data.master;
const materialEnd = material.length + 1;
const continuousEnd = continuousMaterial.length + 1;
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

const excludedHeaders = ["年月", "日期", "单据类型", "单据号", "仓库编码", "仓库名称", "仓库核算组", "存货编码", "存货名称", "规格型号", "收发类别", "记账人", "剔除数量", "剔除单价", "剔除金额", "来源文件", "来源行号"];
const incomeDifferenceHeaders = [
  "年月", "存货编码", "存货名称", "规格型号", "计量单位",
  "用友U8月度收入数量", "剔除收入数量", "用友U8月度收入数量（剔除后）", "CAATS计入收入数量", "收入数量差异（U8剔除后-CAATS）",
  "用友U8月度收入单价", "用友U8月度收入金额", "剔除收入金额", "用友U8月度收入单价（剔除后）", "用友U8月度收入金额（剔除后）",
  "CAATS计入收入单价", "CAATS计入收入金额", "收入金额差异（U8剔除后-CAATS）", "差异状态", "月度汇总来源文件", "来源行号", "差异说明",
];

function populateExcludedDetailSheet(sheet, rows, headerFill) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(4);
  sheet.getRange("A1:Q1").values = [excludedHeaders];
  sheet.getRange("A1:Q1").format = headerFormat(headerFill);
  const values = rows.map((row) => [row.month, row.date, row.document_type, row.document_number, row.warehouse_code, row.warehouse, row.warehouse_group, row.code, row.name, row.spec, row.category, row.posted_by, row.quantity, row.unit_price, row.amount, row.source_file, row.source_row]);
  if (values.length) {
    sheet.getRange(`A2:Q${values.length + 1}`).values = values;
    sheet.getRange(`M2:O${values.length + 1}`).format.numberFormat = numberFormat;
    sheet.getRange(`E2:E${values.length + 1}`).format.numberFormat = "00";
    sheet.getRange(`H2:H${values.length + 1}`).format.numberFormat = "0000";
    applyTableGrid(sheet, `A1:Q${values.length + 1}`);
    sheet.getRange("A1:Q1").format = headerFormat(headerFill);
  }
  for (const column of ["A", "B", "C", "D", "E", "G", "H", "J", "K", "L", "Q"]) sheet.getRange(`${column}1:${column}${Math.max(20, values.length + 1)}`).format.columnWidth = 15;
  sheet.getRange(`F1:F${Math.max(20, values.length + 1)}`).format.columnWidth = 34;
  sheet.getRange(`I1:I${Math.max(20, values.length + 1)}`).format.columnWidth = 28;
  for (const column of ["M", "N", "O"]) sheet.getRange(`${column}1:${column}${Math.max(20, values.length + 1)}`).format.columnWidth = 18;
  sheet.getRange(`P1:P${Math.max(20, values.length + 1)}`).format.columnWidth = 40;
}

function populateIncomeDifferenceSheet(sheet, rows) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(5);
  sheet.getRange("A1:V1").values = [incomeDifferenceHeaders];
  sheet.getRange("A1:V1").format = headerFormat(palette.incomeHeader);
  sheet.getRange("A1:V1").format.rowHeight = 64;
  const values = rows.map((row) => [
    row.month, row.code, row.name, row.spec, row.unit,
    row["收入数量"], row.excluded_posted_iq, null, row.filtered_ledger_iq, null,
    row["收入单价"], row["收入金额"], row.excluded_posted_ia, row.u8_filtered_income_price, null,
    null, row.filtered_ledger_ia, null, null, row.source_file, row.source_row,
    "月度汇总与筛选后流水的金额桥接差；需结合U8成本调整明细进一步解释",
  ]);
  if (values.length) {
    const end = values.length + 1;
    sheet.getRange(`A2:V${end}`).values = values;
    const formulas = { H: [], J: [], O: [], P: [], R: [], S: [] };
    for (let row = 2; row <= end; row += 1) {
      formulas.H.push([`=F${row}-G${row}`]);
      formulas.J.push([`=H${row}-I${row}`]);
      formulas.O.push([`=L${row}-M${row}`]);
      formulas.P.push([`=IF(I${row}=0,0,Q${row}/I${row})`]);
      formulas.R.push([`=O${row}-Q${row}`]);
      formulas.S.push([`=IF(ABS(ROUND(R${row},10))>${amountTolerance},"REVIEW","容差内")`]);
    }
    for (const [column, formulasForColumn] of Object.entries(formulas)) {
      sheet.getRange(`${column}2:${column}${end}`).formulas = formulasForColumn;
    }
    sheet.getRange(`F2:R${end}`).format.numberFormat = numberFormat;
    sheet.getRange(`A2:A${end}`).format.numberFormat = "0";
    sheet.getRange(`B2:B${end}`).format.numberFormat = "0000";
    sheet.getRange(`U2:U${end}`).format.numberFormat = "0";
    applyTableGrid(sheet, `A1:V${end}`);
    sheet.getRange("A1:V1").format = headerFormat(palette.incomeHeader);
    sheet.getRange("A1:V1").format.rowHeight = 64;
    applyStatusFormatting(sheet, `S2:S${end}`);
  }
  for (const column of ["A", "B", "D", "E", "T", "U"]) sheet.getRange(`${column}1:${column}${Math.max(20, values.length + 1)}`).format.columnWidth = 15;
  sheet.getRange(`C1:C${Math.max(20, values.length + 1)}`).format.columnWidth = 30;
  for (const column of ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S"]) sheet.getRange(`${column}1:${column}${Math.max(20, values.length + 1)}`).format.columnWidth = 18;
  sheet.getRange(`T1:T${Math.max(20, values.length + 1)}`).format.columnWidth = 38;
  sheet.getRange(`V1:V${Math.max(20, values.length + 1)}`).format.columnWidth = 48;
  sheet.getRange(`V2:V${Math.max(20, values.length + 1)}`).format.wrapText = true;
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
  const independentDifferenceName = "05_月度独立差异明细（对应01）";
  const continuousDifferenceName = "06_连续滚算差异明细（对应03）";
  resultSheet.name = resultSheetName;
  const summarySheet = workbook.worksheets.add("02_物料月维度结果汇总");
  const continuousDetailSheet = workbook.worksheets.add(continuousDetailName);
  const continuousSummarySheet = workbook.worksheets.add("04_连续滚算汇总");
  const differenceSheet = workbook.worksheets.add(independentDifferenceName);
  const continuousDifferenceSheet = workbook.worksheets.add(continuousDifferenceName);
  const offsetSheet = workbook.worksheets.add("07_跨仓抵销");

  resultSheet.showGridLines = true;
  resultSheet.freezePanes.freezeRows(1);
  resultSheet.freezePanes.freezeColumns(4);
  const identityHeaders = resultSheet.getRange("A1:I1").values[0].map((value) => typeof value === "string" ? value.replaceAll("CAATs", "CAATS") : value);
  const resultHeaders = [
    ...identityHeaders,
    "客户U8期初数量", "客户U8期初单价", "客户U8期初金额",
    "用友U8月度收入数量", "剔除收入数量", "用友U8月度收入数量（剔除后）", "CAATS计入收入数量", "收入数量差异（U8剔除后-CAATS）",
    "用友U8月度收入单价", "用友U8月度收入金额", "剔除收入金额", "用友U8月度收入单价（剔除后）", "用友U8月度收入金额（剔除后）", "CAATS计入收入金额", "收入金额差异（U8剔除后-CAATS）",
    "用友U8月度发出数量", "剔除发出数量", "用友U8月度发出数量（剔除后）", "CAATS计入发出数量", "发出数量差异（U8剔除后-CAATS）",
    "用友U8月度发出单价", "用友U8月度发出金额", "剔除发出金额", "用友U8月度发出单价（剔除后）", "用友U8月度发出金额（剔除后）", "CAATS发出金额", "发出金额差异（U8剔除后-CAATS）",
    "用友U8月度结存数量", "用友U8月度结存单价", "用友U8月度结存金额", "用友U8月度结存数量（剔除后）", "用友U8月度结存单价（剔除后）", "用友U8月度结存金额（剔除后）", "CAATS结存金额", "结存金额差异（U8剔除后-CAATS）",
    "CAATS月末平均价", "CAATS结存数量", "结存数量差异（U8剔除后-CAATS）", "结果状态", "主要差异原因",
  ];
  resultSheet.getRange("A1:AW1").values = [resultHeaders];
  const resultValues = material.map((row) => [
    row.code, row.month, row.stock_code || null, row.name, row.spec, row.unit, row.weight || null, row.reg || null, row.manufacturer || null,
    row["期初数量"], row["期初单价"], row["期初金额"],
    row["收入数量"], row.excluded_posted_iq, null, row.filtered_ledger_iq, null, row["收入单价"], row["收入金额"], row.excluded_posted_ia, row.u8_filtered_income_price, null, row.filtered_ledger_ia, null,
    row["发出数量"], row.excluded_posted_oq, null, row.filtered_ledger_oq, null, row["发出单价"], row["发出金额"], row.excluded_posted_oa, row.u8_filtered_issue_price, null, null, null,
    row["结存数量"], row["结存单价"], row["结存金额"], row.u8_filtered_end_quantity, row.u8_filtered_end_price, row.u8_filtered_end_amount, null, null,
    null, null, null, null, row.reason,
  ]);
  resultSheet.getRange(`A2:AW${materialEnd}`).values = resultValues;
  const formulas = { O: [], Q: [], V: [], X: [], AA: [], AC: [], AH: [], AI: [], AJ: [], AQ: [], AR: [], AS: [], AT: [], AU: [], AV: [] };
  for (let row = 2; row <= materialEnd; row += 1) {
    formulas.O.push([`=M${row}-N${row}`]);
    formulas.Q.push([`=O${row}-P${row}`]);
    formulas.V.push([`=S${row}-T${row}`]);
    formulas.X.push([`=V${row}-W${row}`]);
    formulas.AA.push([`=Y${row}-Z${row}`]);
    formulas.AC.push([`=AA${row}-AB${row}`]);
    formulas.AH.push([`=AE${row}-AF${row}`]);
    formulas.AI.push([`=AB${row}*AS${row}`]);
    formulas.AJ.push([`=AH${row}-AI${row}`]);
    formulas.AQ.push([`=L${row}+W${row}-AI${row}`]);
    formulas.AR.push([`=AP${row}-AQ${row}`]);
    formulas.AS.push([`=IF(J${row}+P${row}=0,0,(L${row}+W${row})/(J${row}+P${row}))`]);
    formulas.AT.push([`=J${row}+P${row}-AB${row}`]);
    formulas.AU.push([`=AN${row}-AT${row}`]);
    formulas.AV.push([`=IF(ABS(ROUND(AJ${row},10))<=${amountTolerance},"PASS","REVIEW")`]);
  }
  for (const [column, values] of Object.entries(formulas)) {
    resultSheet.getRange(`${column}2:${column}${materialEnd}`).formulas = values;
  }
  resultSheet.getRange("A1:I1").format = headerFormat(palette.title);
  resultSheet.getRange("J1:L1").format = headerFormat(palette.sectionHeader);
  resultSheet.getRange("M1:X1").format = headerFormat(palette.incomeHeader);
  resultSheet.getRange("Y1:AJ1").format = headerFormat(palette.issueHeader);
  resultSheet.getRange("AK1:AR1").format = headerFormat(palette.closingHeader);
  resultSheet.getRange("AS1:AW1").format = headerFormat(palette.resultHeader);
  resultSheet.getRange("A1:AW1").format.rowHeight = 64;
  resultSheet.getRange("A1:AW1").format.wrapText = true;
  resultSheet.getRange(`J2:AW${materialEnd}`).format.numberFormat = numberFormat;
  resultSheet.getRange(`A2:A${materialEnd}`).format.numberFormat = "0000";
  resultSheet.getRange(`B2:B${materialEnd}`).format.numberFormat = "0";
  for (const column of ["A", "B", "C", "F", "G", "I"]) resultSheet.getRange(`${column}1:${column}${materialEnd}`).format.columnWidth = 11;
  for (const column of ["D", "H"]) resultSheet.getRange(`${column}1:${column}${materialEnd}`).format.columnWidth = 24;
  resultSheet.getRange(`E1:E${materialEnd}`).format.columnWidth = 18;
  for (const column of ["J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV"]) resultSheet.getRange(`${column}1:${column}${materialEnd}`).format.columnWidth = 15;
  resultSheet.getRange(`AW1:AW${materialEnd}`).format.columnWidth = 36;
  resultSheet.getRange(`AW2:AW${materialEnd}`).format.wrapText = true;
  applyStatusFormatting(resultSheet, `AV2:AV${materialEnd}`);

  continuousDetailSheet.showGridLines = true;
  continuousDetailSheet.freezePanes.freezeRows(1);
  continuousDetailSheet.freezePanes.freezeColumns(4);
  const continuousHeaders = [...resultHeaders, "记录类型", "U8期初数量(核对)", "U8期初金额(核对)", "连续起算方式", "期初承接结存差异（U8-连续CAATS）", "金额滚转钩稽差异", "金额滚转钩稽状态"];
  continuousDetailSheet.getRange("A1:BD1").values = [continuousHeaders];
  continuousDetailSheet.getRange("J1:L1").values = [["连续期初数量", "连续期初单价", "连续期初金额"]];
  continuousDetailSheet.getRange("AQ1:AW1").values = [["连续CAATS结存金额", "结存金额差异（U8剔除后-连续CAATS）", "连续CAATS月末平均价", "连续CAATS结存数量", "结存数量差异（U8剔除后-连续CAATS）", "连续结果状态", "连续滚算说明"]];
  continuousMaterial.sort((left, right) => {
    const codeOrder = String(left.code).localeCompare(String(right.code), "zh-CN", { numeric: true });
    return codeOrder || String(left.month).localeCompare(String(right.month));
  });
  const continuousDetailValues = continuousMaterial.map((row) => [
    row.code, row.month, row.stock_code || null, row.name, row.spec, row.unit, row.weight || null, row.reg || null, row.manufacturer || null,
    null, null, null,
    row["收入数量"], row.excluded_posted_iq, null, row.filtered_ledger_iq, null, row["收入单价"], row["收入金额"], row.excluded_posted_ia, row.u8_filtered_income_price, null, row.filtered_ledger_ia, null,
    row["发出数量"], row.excluded_posted_oq, null, row.filtered_ledger_oq, null, row["发出单价"], row["发出金额"], row.excluded_posted_oa, row.u8_filtered_issue_price, null, null, null,
    row["结存数量"], row["结存单价"], row["结存金额"], row.u8_filtered_end_quantity, row.u8_filtered_end_price, row.u8_filtered_end_amount, null, null,
    null, null, null, null, null,
    row.record_type, row["期初数量"], row["期初金额"], null, null, null, null,
  ]);
  continuousDetailSheet.getRange(`A2:BD${continuousEnd}`).values = continuousDetailValues;
  const continuousDetailFormulas = { J: [], K: [], L: [], O: [], Q: [], V: [], X: [], AA: [], AC: [], AH: [], AI: [], AJ: [], AQ: [], AR: [], AS: [], AT: [], AU: [], AV: [], AW: [], BA: [], BB: [], BC: [], BD: [] };
  for (let row = 2; row <= continuousEnd; row += 1) {
    const canRoll = row === 2 ? null : `A${row}=A${row - 1}`;
    continuousDetailFormulas.J.push([row === 2 ? `=AY${row}` : `=IF(${canRoll},AT${row - 1},AY${row})`]);
    continuousDetailFormulas.K.push([`=IF(J${row}=0,0,L${row}/J${row})`]);
    continuousDetailFormulas.L.push([row === 2 ? `=AZ${row}` : `=IF(${canRoll},AQ${row - 1},AZ${row})`]);
    continuousDetailFormulas.O.push([`=M${row}-N${row}`]);
    continuousDetailFormulas.Q.push([`=O${row}-P${row}`]);
    continuousDetailFormulas.V.push([`=S${row}-T${row}`]);
    continuousDetailFormulas.X.push([`=V${row}-W${row}`]);
    continuousDetailFormulas.AA.push([`=Y${row}-Z${row}`]);
    continuousDetailFormulas.AC.push([`=AA${row}-AB${row}`]);
    continuousDetailFormulas.AH.push([`=AE${row}-AF${row}`]);
    continuousDetailFormulas.AI.push([`=AB${row}*AS${row}`]);
    continuousDetailFormulas.AJ.push([`=AH${row}-AI${row}`]);
    continuousDetailFormulas.AQ.push([`=L${row}+W${row}-AI${row}`]);
    continuousDetailFormulas.AR.push([`=AP${row}-AQ${row}`]);
    continuousDetailFormulas.AS.push([`=IF(J${row}+P${row}=0,0,(L${row}+W${row})/(J${row}+P${row}))`]);
    continuousDetailFormulas.AT.push([`=J${row}+P${row}-AB${row}`]);
    continuousDetailFormulas.AU.push([`=AN${row}-AT${row}`]);
    continuousDetailFormulas.AV.push([`=IF(ABS(ROUND(AJ${row},10))<=${amountTolerance},"PASS","REVIEW")`]);
    continuousDetailFormulas.AW.push([`=IF(BA${row}="使用U8期初","该物料首次出现，使用U8期初","承接最近一次可追溯连续CAATS期末")`]);
    continuousDetailFormulas.BA.push([row === 2 ? `="使用U8期初"` : `=IF(${canRoll},"承接最近可追溯连续期末","使用U8期初")`]);
    continuousDetailFormulas.BB.push([`=AZ${row}-L${row}`]);
    continuousDetailFormulas.BC.push([`=BB${row}+X${row}-AJ${row}-AR${row}`]);
    continuousDetailFormulas.BD.push([`=IF(ABS(ROUND(BC${row},10))<=${amountTolerance},"PASS","REVIEW")`]);
  }
  for (const [column, values] of Object.entries(continuousDetailFormulas)) {
    continuousDetailSheet.getRange(`${column}2:${column}${continuousEnd}`).formulas = values;
  }
  continuousDetailSheet.getRange("A1:I1").format = headerFormat(palette.title);
  continuousDetailSheet.getRange("J1:L1").format = headerFormat(palette.sectionHeader);
  continuousDetailSheet.getRange("M1:X1").format = headerFormat(palette.incomeHeader);
  continuousDetailSheet.getRange("Y1:AJ1").format = headerFormat(palette.issueHeader);
  continuousDetailSheet.getRange("AK1:AR1").format = headerFormat(palette.closingHeader);
  continuousDetailSheet.getRange("AS1:AW1").format = headerFormat(palette.resultHeader);
  continuousDetailSheet.getRange("AX1:BB1").format = headerFormat(palette.sectionHeader);
  continuousDetailSheet.getRange("BC1:BD1").format = headerFormat(palette.resultHeader);
  continuousDetailSheet.getRange("A1:BD1").format.rowHeight = 64;
  continuousDetailSheet.getRange("A1:BD1").format.wrapText = true;
  continuousDetailSheet.getRange(`J2:AW${continuousEnd}`).format.numberFormat = numberFormat;
  continuousDetailSheet.getRange(`AY2:AZ${continuousEnd}`).format.numberFormat = numberFormat;
  continuousDetailSheet.getRange(`BB2:BC${continuousEnd}`).format.numberFormat = numberFormat;
  continuousDetailSheet.getRange(`A2:A${continuousEnd}`).format.numberFormat = "0000";
  continuousDetailSheet.getRange(`B2:B${continuousEnd}`).format.numberFormat = "0";
  for (const column of ["A", "B", "C", "F", "G", "I"]) continuousDetailSheet.getRange(`${column}1:${column}${continuousEnd}`).format.columnWidth = 11;
  for (const column of ["D", "H"]) continuousDetailSheet.getRange(`${column}1:${column}${continuousEnd}`).format.columnWidth = 24;
  continuousDetailSheet.getRange(`E1:E${continuousEnd}`).format.columnWidth = 18;
  for (const column of ["J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AY", "AZ", "BB", "BC"]) continuousDetailSheet.getRange(`${column}1:${column}${continuousEnd}`).format.columnWidth = 15;
  continuousDetailSheet.getRange(`AW1:AW${continuousEnd}`).format.columnWidth = 36;
  continuousDetailSheet.getRange(`AX1:AX${continuousEnd}`).format.columnWidth = 18;
  continuousDetailSheet.getRange(`BA1:BA${continuousEnd}`).format.columnWidth = 24;
  continuousDetailSheet.getRange(`BD1:BD${continuousEnd}`).format.columnWidth = 16;
  continuousDetailSheet.getRange(`AW2:BA${continuousEnd}`).format.wrapText = true;
  applyStatusFormatting(continuousDetailSheet, `AV2:AV${continuousEnd}`);
  applyStatusFormatting(continuousDetailSheet, `BD2:BD${continuousEnd}`);

  summarySheet.showGridLines = false;
  summarySheet.getRange("A1:J1").merge();
  summarySheet.getRange("A1").values = [["U8存货发出计价——物料月维度结果汇总（CAATS双重筛选）"]];
  summarySheet.getRange("A1:J1").format = titleFormat;
  summarySheet.getRange("A3:B9").values = [
    ["指标", "结果"],
    ["物料月份行数", null],
    ["PASS行数", null],
    ["REVIEW行数", null],
    ["发出金额差异净额", null],
    ["总体正负抵销金额", data.metrics.global_netting],
    ["同物料月跨仓抵销金额", data.offset_summary.total_offset],
  ];
  summarySheet.getRange("A3:B3").format = headerFormat(palette.sectionHeader);
  summarySheet.getRange("B4:B8").formulas = [
    [`=COUNTA('${resultSheetName}'!A2:A${materialEnd})`],
    [`=COUNTIF('${resultSheetName}'!AV2:AV${materialEnd},"PASS")`],
    [`=COUNTIF('${resultSheetName}'!AV2:AV${materialEnd},"REVIEW")`],
    [`=SUM('${resultSheetName}'!AJ2:AJ${materialEnd})`],
    [`=MIN(SUMIF('${resultSheetName}'!AJ2:AJ${materialEnd},">0",'${resultSheetName}'!AJ2:AJ${materialEnd}),-SUMIF('${resultSheetName}'!AJ2:AJ${materialEnd},"<0",'${resultSheetName}'!AJ2:AJ${materialEnd}))`],
  ];
  summarySheet.getRange("D3:G3").values = [["结论项目", "结果/口径", "状态", "审阅提示"]];
  summarySheet.getRange("D3:G3").format = headerFormat(palette.tableHeader);
  summarySheet.getRange("D4:G10").values = [
    ["月度独立口径", "物料+月份；先从U8月表剔除四类移动，再与CAATS筛选取数及重算结果比较", "已执行", "表内依次展示U8原值、U8剔除后、CAATS及差异"],
    ["连续滚算", `6月期末差异净额${data.continuous_monthly.at(-1).continuous_end_difference.toFixed(2)}元`, "敏感性分析", "未模拟自动调整单及最终取价调整"],
    ["跨仓抵销", `${data.offset_summary.both_sign_keys}个物料月份存在仓库正负贡献`, "已识别", "仓库贡献相对物料统一价，不代表仓库错账"],
    ["客户实际配置", "普通单据按仓；特殊单据按仓库组", "配置一致", "物料统一价仅为审计模拟"],
    ["差异方向", "用友U8月度发出金额（剔除后）-CAATS发出金额", "统一", "范围先统一，再评价重计价差异"],
    ["支持证据", "另见ITA核对钩稽数据工作簿", "职责分离", "CAATS表不保存流水及配置底表"],
    ["差异拆分", `U8剔除后与CAATS计入范围差异${data.metrics.scope_issue_difference.toFixed(2)}元；重计价差异${data.metrics.repricing_issue_difference.toFixed(2)}元`, "已拆分", `两项合计为U8剔除后-CAATS总差异${data.metrics.signed_issue_diff.toFixed(2)}元`],
  ];
  summarySheet.getRange("A12:I12").values = [["年月", "物料月份数", "REVIEW数", "PASS数", "用友U8月度发出金额（剔除后）", "CAATS发出金额", "发出差异净额", "REVIEW占比", "说明"]];
  summarySheet.getRange("A12:I12").format = headerFormat(palette.tableHeader);
  for (let index = 0; index < data.monthly.length; index += 1) {
    const row = 13 + index;
    const month = data.monthly[index].month;
    summarySheet.getRange(`A${row}`).values = [[month]];
    summarySheet.getRange(`B${row}:H${row}`).formulas = [[
      `=COUNTIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row})`,
      `=COUNTIFS('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AV$2:$AV$${materialEnd},"REVIEW")`,
      `=COUNTIFS('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AV$2:$AV$${materialEnd},"PASS")`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AH$2:$AH$${materialEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AI$2:$AI$${materialEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AJ$2:$AJ$${materialEnd})`,
      `=IF(B${row}=0,0,C${row}/B${row})`,
    ]];
    summarySheet.getRange(`I${row}`).values = [["差异方向：U8月表剔除四类移动后-CAATS"]];
  }
  summarySheet.getRange("B4:B9").format.numberFormat = numberFormat;
  summarySheet.getRange("E13:G18").format.numberFormat = numberFormat;
  summarySheet.getRange("H13:H18").format.numberFormat = "0.0%";
  summarySheet.getRange("A3:B9").format.borders = thinBorder;
  summarySheet.getRange("D3:G10").format.borders = thinBorder;
  summarySheet.getRange("A12:I18").format.borders = thinBorder;
  summarySheet.getRange("A1:A22").format.columnWidth = 22;
  summarySheet.getRange("B1:C22").format.columnWidth = 16;
  summarySheet.getRange("D1:D22").format.columnWidth = 28;
  summarySheet.getRange("E1:F22").format.columnWidth = 25;
  summarySheet.getRange("G1:G22").format.columnWidth = 44;
  summarySheet.getRange("H1:I22").format.columnWidth = 18;
  summarySheet.getRange("D4:G10").format.wrapText = true;

  differenceSheet.showGridLines = false;
  differenceSheet.freezePanes.freezeRows(1);
  differenceSheet.freezePanes.freezeColumns(2);
  const differenceHeaders = ["年月", "存货编码", "存货名称", "规格", "用友U8月度发出金额（剔除后）", "CAATS发出金额", "发出金额差异（U8剔除后-CAATS）", "结果状态", "主要差异原因"];
  differenceSheet.getRange("A1:I1").values = [differenceHeaders];
  differenceSheet.getRange("A1:I1").format = headerFormat(palette.issueHeader);
  const differenceValues = data.difference_rows.map((row) => [row.month, row.code, row.name, row.spec, row.u8_filtered_issue_amount, row.caats_issue, row.issue_diff, row.status, row.reason]);
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

  continuousDifferenceSheet.showGridLines = false;
  continuousDifferenceSheet.freezePanes.freezeRows(1);
  continuousDifferenceSheet.freezePanes.freezeColumns(2);
  const continuousDifferenceHeaders = [
    "年月", "存货编码", "存货名称", "规格", "记录类型",
    "期初承接结存差异（U8-连续CAATS）", "本期收入金额差异（U8剔除后-CAATS）",
    "本期发出金额差异（U8剔除后-连续CAATS）", "本期期末结存差异（U8剔除后-连续CAATS）",
    "金额滚转钩稽差异", "滚转钩稽状态", "连续结果状态", "连续起算方式",
  ];
  continuousDifferenceSheet.getRange("A1:M1").values = [continuousDifferenceHeaders];
  continuousDifferenceSheet.getRange("A1:M1").format = headerFormat(palette.issueHeader);
  continuousDifferenceSheet.getRange("A1:M1").format.rowHeight = 54;
  const continuousDifferenceValues = data.continuous_difference_rows.map((row) => [
    row.month, row.code, row.name, row.spec, row.record_type,
    row.continuous_begin_difference, row.income_amount_difference, row.continuous_issue_difference, row.continuous_end_difference,
    row.rollforward_check, row.rollforward_status, row.status, row.start_type,
  ]);
  if (continuousDifferenceValues.length) {
    continuousDifferenceSheet.getRange(`A2:M${continuousDifferenceValues.length + 1}`).values = continuousDifferenceValues;
    continuousDifferenceSheet.getRange(`F2:J${continuousDifferenceValues.length + 1}`).format.numberFormat = numberFormat;
    continuousDifferenceSheet.getRange(`B2:B${continuousDifferenceValues.length + 1}`).format.numberFormat = "0000";
    applyTableGrid(continuousDifferenceSheet, `A1:M${continuousDifferenceValues.length + 1}`);
    continuousDifferenceSheet.getRange("A1:M1").format = headerFormat(palette.issueHeader);
    continuousDifferenceSheet.getRange("A1:M1").format.rowHeight = 54;
    applyStatusFormatting(continuousDifferenceSheet, `K2:L${continuousDifferenceValues.length + 1}`);
  }
  for (const column of ["A", "B", "D", "E", "K", "L"]) continuousDifferenceSheet.getRange(`${column}1:${column}${continuousDifferenceValues.length + 1}`).format.columnWidth = 15;
  continuousDifferenceSheet.getRange(`C1:C${continuousDifferenceValues.length + 1}`).format.columnWidth = 28;
  for (const column of ["F", "G", "H", "I", "J"]) continuousDifferenceSheet.getRange(`${column}1:${column}${continuousDifferenceValues.length + 1}`).format.columnWidth = 20;
  continuousDifferenceSheet.getRange(`M1:M${continuousDifferenceValues.length + 1}`).format.columnWidth = 30;

  continuousSummarySheet.showGridLines = false;
  continuousSummarySheet.getRange("A1:I1").merge();
  continuousSummarySheet.getRange("A1").values = [["连续滚算汇总（敏感性分析）"]];
  continuousSummarySheet.getRange("A1:I1").format = titleFormat;
  continuousSummarySheet.getRange("A3:B11").values = [
    ["指标", "结果"],
    ["完整连续面板行数", data.metrics.continuous_panel_rows],
    ["补齐无收发月份行数", data.metrics.continuous_synthetic_rows],
    ["Final物料数", data.metrics.final_material_rows],
    ["Final补齐记录数", data.metrics.final_synthetic_rows],
    ["Final用友U8结存金额（剔除后）", data.metrics.final_u8_end_amount],
    ["Final连续CAATS结存金额", data.metrics.final_caats_end_amount],
    ["Final结存金额差异", data.metrics.final_end_difference],
    ["金额滚转钩稽异常数", data.metrics.rollforward_review_rows],
  ];
  continuousSummarySheet.getRange("A3:B3").format = headerFormat(palette.sectionHeader);
  continuousSummarySheet.getRange("A3:B11").format.borders = thinBorder;
  continuousSummarySheet.getRange("B8:B10").format.numberFormat = numberFormat;
  continuousSummarySheet.getRange("D3:I3").values = [["金额滚转公式", "期初承接差异", "本期收入差异", "本期发出差异", "本期期末差异", "钩稽结论"]];
  continuousSummarySheet.getRange("D3:I3").format = headerFormat(palette.tableHeader);
  continuousSummarySheet.getRange("D4:I4").values = [["期初+收入-发出-期末=0", "本月U8期初-连续CAATS期初", "U8剔除后收入-CAATS收入", "U8剔除后发出-连续CAATS发出", "U8剔除后期末-连续CAATS期末", "发出与期末不得脱离期初及收入单独相加"]];
  continuousSummarySheet.getRange("D3:I4").format.borders = thinBorder;
  continuousSummarySheet.getRange("D4:I4").format.wrapText = true;
  continuousSummarySheet.getRange("D6:I6").values = [["结果项目", "属性", "汇总方式", "是否相加", "用途", "说明"]];
  continuousSummarySheet.getRange("D6:I6").format = headerFormat(palette.tableHeader);
  continuousSummarySheet.getRange("D7:I8").values = [
    ["连续发出金额差异", "期间流量", "1-6月逐月求和", "否", "评价期间发出计价", "需结合期初、收入和期末桥接"],
    ["Final结存金额差异", "期末时点", "每个物料仅取6月Final", "否", "评价期末存货", "无收发月份承接上期，不跨月累计"],
  ];
  continuousSummarySheet.getRange("D6:I8").format.borders = thinBorder;
  continuousSummarySheet.getRange("D7:I8").format.wrapText = true;
  const continuousSummaryHeaders = ["年月", "连续面板物料数", "用友U8月度发出金额（剔除后）", "月度独立CAATS发出", "月度独立发出差异", "连续CAATS发出", "连续发出金额差异", "连续期末结存差异", "上期差异传导影响"];
  continuousSummarySheet.getRange("A13:I13").values = [continuousSummaryHeaders];
  continuousSummarySheet.getRange("A13:I13").format = headerFormat(palette.tableHeader);
  for (let index = 0; index < data.monthly.length; index += 1) {
    const row = 14 + index;
    const month = data.monthly[index].month;
    continuousSummarySheet.getRange(`A${row}`).values = [[month]];
    continuousSummarySheet.getRange(`B${row}:I${row}`).formulas = [[
      `=COUNTIF('${continuousDetailName}'!$B$2:$B$${continuousEnd},A${row})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${continuousEnd},A${row},'${continuousDetailName}'!$AH$2:$AH$${continuousEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AI$2:$AI$${materialEnd})`,
      `=SUMIF('${resultSheetName}'!$B$2:$B$${materialEnd},A${row},'${resultSheetName}'!$AJ$2:$AJ$${materialEnd})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${continuousEnd},A${row},'${continuousDetailName}'!$AI$2:$AI$${continuousEnd})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${continuousEnd},A${row},'${continuousDetailName}'!$AJ$2:$AJ$${continuousEnd})`,
      `=SUMIF('${continuousDetailName}'!$B$2:$B$${continuousEnd},A${row},'${continuousDetailName}'!$AR$2:$AR$${continuousEnd})`,
      `=G${row}-E${row}`,
    ]];
  }
  const continuousTotalRow = data.monthly.length + 14;
  continuousSummarySheet.getRange(`A${continuousTotalRow}:I${continuousTotalRow}`).values = [["合计/Final", null, null, null, null, null, null, null, null]];
  continuousSummarySheet.getRange(`B${continuousTotalRow}:I${continuousTotalRow}`).formulas = [[
    `=B${continuousTotalRow - 1}`, `=SUM(C14:C${continuousTotalRow - 1})`, `=SUM(D14:D${continuousTotalRow - 1})`, `=SUM(E14:E${continuousTotalRow - 1})`, `=SUM(F14:F${continuousTotalRow - 1})`, `=SUM(G14:G${continuousTotalRow - 1})`, `=H${continuousTotalRow - 1}`, `=SUM(I14:I${continuousTotalRow - 1})`,
  ]];
  continuousSummarySheet.getRange(`A13:I${continuousTotalRow}`).format.borders = thinBorder;
  continuousSummarySheet.getRange(`C14:I${continuousTotalRow}`).format.numberFormat = numberFormat;
  continuousSummarySheet.getRange(`A${continuousTotalRow}:I${continuousTotalRow}`).format.font = { bold: true };
  continuousSummarySheet.getRange("A1:A22").format.columnWidth = 24;
  continuousSummarySheet.getRange("B1:B22").format.columnWidth = 18;
  continuousSummarySheet.getRange("C1:I22").format.columnWidth = 20;

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
    ["抵销金额", "MIN(正贡献合计,-负贡献合计)", "展示总体物料结果掩盖的仓库差异", "不能直接记账"],
  ];
  offsetSheet.getRange("D3:G3").format = headerFormat(palette.tableHeader);
  offsetSheet.getRange("D3:G6").format.borders = thinBorder;
  offsetSheet.getRange("D4:G6").format.wrapText = true;
  const offsetHeaders = ["年月", "存货编码", "存货名称", "规格", "物料统一平均价", "物料发出差异", "仓库编码", "仓库名称", "仓库核算组", "仓库发出数量", "仓库发出金额", "统一价预期金额", "仓库贡献", "同物料月正贡献", "同物料月负贡献（负数）", "同物料月抵销金额", "同物料月净额"];
  offsetSheet.getRange("A9:Q9").values = [offsetHeaders];
  offsetSheet.getRange("A9:Q9").format = headerFormat(palette.issueHeader);
  const offsetValues = data.offset_rows.map((row) => [row.month, row.code, row.name, row.spec, row.pooled_avg, row.material_issue_difference, row.warehouse_code, row.warehouse, row.group, row.warehouse_issue_quantity, row.warehouse_issue_amount, row.pooled_expected_amount, row.warehouse_contribution, row.key_positive, -row.key_negative_abs, row.key_offset, row.key_net]);
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

  if (!skipQa) {
    await saveInspection(workbook, "caats", [["01_物料月维度明细", "A1:AW20"], ["02_物料月维度结果汇总", "A1:I18"], ["03_连续滚算明细", "A1:BD20"], ["04_连续滚算汇总", "A1:I20"], [independentDifferenceName, "A1:I20"], [continuousDifferenceName, "A1:M20"]]);
    await renderSheets(workbook, "caats", [
      ["01_物料月维度明细", "A1:AW24", 0.6],
      ["02_物料月维度结果汇总", "A1:I18", 1.0],
      ["03_连续滚算明细", "A1:BD24", 0.56],
      ["04_连续滚算汇总", "A1:I20", 0.95],
      [independentDifferenceName, "A1:I24", 0.95],
      [continuousDifferenceName, "A1:M24", 0.8],
      ["07_跨仓抵销", "A1:Q22", 0.8],
    ]);
  }
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
    ["CAATS本期收发", "已记账成本流水", "剔除调拨及借用四类移动", "已执行", `${data.metrics.excluded_income_detail_rows + data.metrics.excluded_issue_detail_rows}条已记账成本收发明细形成01及03剔除桥`],
    ["本期收发钩稽", "剔除后口径vs原月表", "不要求一致", "不适用", "差额即剔除移动及原流水口径影响"],
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
  const bridgeHeaders = ["年月", "存货编码", "存货名称", "U8月表收入数量", "四类移动入库数量", "CAATS计入收入数量", "U8月表收入金额", "四类移动入库金额", "CAATS计入收入金额", "U8月表发出数量", "四类移动出库数量", "CAATS计入发出数量", "U8月表发出金额", "四类移动出库金额", "CAATS筛选后流水发出金额(观察)", "流水已记账收入数量", "流水已记账发出数量", "剔除移动行数", "计算口径", "本期收发钩稽", "说明"];
  bridgeSheet.getRange("A1:U1").values = [bridgeHeaders];
  bridgeSheet.getRange("A1:U1").format = headerFormat(palette.tableHeader);
  bridgeSheet.getRange("A1:U1").format.rowHeight = 48;
  const bridgeValues = material.map((row) => [
    row.month, row.code, row.name, row["收入数量"], row.excluded_iq, row.filtered_ledger_iq,
    row["收入金额"], row.excluded_ia, row.filtered_ledger_ia, row["发出数量"], row.excluded_oq, row.filtered_ledger_oq,
    row["发出金额"], row.excluded_oa, row.filtered_ledger_oa, row.ledger_iq, row.ledger_oq, row.excluded_rows,
    "记账人非空且排除四类移动", "展示差异", "U8月表原值与CAATS筛选值分别保留，不要求两者相等",
  ]);
  bridgeSheet.getRange(`A2:U${bridgeValues.length + 1}`).values = bridgeValues;
  bridgeSheet.getRange(`D2:S${bridgeValues.length + 1}`).format.numberFormat = numberFormat;
  bridgeSheet.getRange(`B2:B${bridgeValues.length + 1}`).format.numberFormat = "0000";
  applyTableGrid(bridgeSheet, `A1:U${bridgeValues.length + 1}`);
  bridgeSheet.getRange("A1:U1").format = headerFormat(palette.tableHeader);
  for (const column of ["A", "B", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"]) bridgeSheet.getRange(`${column}1:${column}${bridgeValues.length + 1}`).format.columnWidth = 15;
  bridgeSheet.getRange(`C1:C${bridgeValues.length + 1}`).format.columnWidth = 28;
  bridgeSheet.getRange(`U1:U${bridgeValues.length + 1}`).format.columnWidth = 44;

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
  const materialBridgeHeaders = ["年月", "存货编码", "存货名称", "U8月表收入数量", "U8月表收入数量（剔除后）", "CAATS计入收入数量", "收入数量差异（U8剔除后-CAATS）", "U8月表收入金额", "U8月表收入金额（剔除后）", "CAATS计入收入金额", "收入金额差异（U8剔除后-CAATS）", "U8月表发出数量", "U8月表发出数量（剔除后）", "CAATS计入发出数量", "发出数量差异（U8剔除后-CAATS）", "U8月表发出金额", "U8月表发出金额（剔除后）", "CAATS筛选后流水发出金额(观察)", "范围差异（U8剔除后-筛选后）", "CAATS重算发出金额", "重计价差异（筛选后-CAATS）", "总差异（U8剔除后-CAATS）", "涉及成本仓数", "涉及出库仓数", "CAATS筛选条件"];
  materialBridgeSheet.getRange("A1:Y1").values = [materialBridgeHeaders];
  materialBridgeSheet.getRange("A1:Y1").format = headerFormat(palette.tableHeader);
  materialBridgeSheet.getRange("A1:Y1").format.rowHeight = 64;
  materialBridgeSheet.getRange("A1:Y1").format.wrapText = true;
  const materialBridgeValues = material.map((row) => [row.month, row.code, row.name, row["收入数量"], row.u8_filtered_income_quantity, row.filtered_ledger_iq, row.u8_filtered_income_quantity - row.filtered_ledger_iq, row["收入金额"], row.u8_filtered_income_amount, row.filtered_ledger_ia, row.u8_filtered_income_amount - row.filtered_ledger_ia, row["发出数量"], row.u8_filtered_issue_quantity, row.filtered_ledger_oq, row.u8_filtered_issue_quantity - row.filtered_ledger_oq, row["发出金额"], row.u8_filtered_issue_amount, row.filtered_ledger_oa, row.u8_filtered_issue_amount - row.filtered_ledger_oa, row.caats_issue, row.filtered_ledger_oa - row.caats_issue, row.issue_diff, row.warehouse_count, row.issue_warehouse_count, "记账人非空且排除四类移动"]);
  materialBridgeSheet.getRange(`A2:Y${materialBridgeValues.length + 1}`).values = materialBridgeValues;
  materialBridgeSheet.getRange(`D2:X${materialBridgeValues.length + 1}`).format.numberFormat = numberFormat;
  materialBridgeSheet.getRange(`B2:B${materialBridgeValues.length + 1}`).format.numberFormat = "0000";
  applyTableGrid(materialBridgeSheet, `A1:Y${materialBridgeValues.length + 1}`);
  materialBridgeSheet.getRange("A1:Y1").format = headerFormat(palette.tableHeader);
  for (const column of ["A", "B", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"]) materialBridgeSheet.getRange(`${column}1:${column}${materialBridgeValues.length + 1}`).format.columnWidth = 16;
  materialBridgeSheet.getRange(`C1:C${materialBridgeValues.length + 1}`).format.columnWidth = 28;

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
    ["四类移动", "调拨出库、调拨入库、借出借用出库、借用归还入库全部排除", "客户指定剔除口径", `${data.metrics.excluded_movement_rows}条原始分类记录；其中${data.metrics.excluded_income_detail_rows + data.metrics.excluded_issue_detail_rows}条已记账成本收发明细形成01及03显式剔除桥`],
    ["CAATS筛选方法", "流水同时满足记账人非空、仓库记入成本且移动方式不属于四类排除项；U8月表先建立四类移动剔除后口径", "04_流水桥接、06_物料月桥接", "CAATS结果统一与U8剔除后口径比较"],
    ["月间衔接", "上月期末与下月期初逐物料比较", "客户月度收发存", "当前数量、金额均已衔接"],
    ["月度独立重算", "物料+月份；每月使用U8期初，CAATS本期收发取双重筛选后的流水", "01_物料月维度明细", "不累计上月CAATS差异"],
    ["连续滚算", "首次出现使用U8期初；从首次出现月至Final建立完整月份面板；后续承接上月连续CAATS期末", "03_连续滚算明细", "BB为期初承接差异，BC为四段式滚转钩稽差异，BD为状态"],
    ["连续滚算缺月", "无收发月份补零收发记录，期初期末承接上月，不重置追溯链", "记录类型及连续起算方式字段", "仅在确认CAATS亦无有效收发时补齐"],
    ["连续滚算钩稽", "期初承接差异+本期收入差异-本期发出差异-本期期末差异=0", "03_连续滚算明细、04_连续滚算汇总", `全量异常${data.metrics.rollforward_review_rows}项；不设置倒挤数`],
    ["Final结存汇总", "同一物料仅取最终期间结存差异，不跨月累计", "04_连续滚算汇总", `${data.metrics.final_material_rows}个Final物料；未模拟自动调整单及最终取价调整`],
    ["差异方向与容差", "U8月表剔除后金额-CAATS金额；正数表示U8剔除后金额较高", `金额容差${amountTolerance}元`, "浮点尾差按10位小数规范化后判断"],
    ["差异拆分", "总差异=U8剔除后与CAATS计入范围差异+重计价差异", "06_物料月桥接", "先统一四类移动口径，再解释剩余流水桥接及计价影响"],
    ["正常单据", "按仓库汇总", "仓库档案全月平均法", "缺分仓期初，暂不能独立复算"],
    ["特殊单据", "按仓库核算组汇总", "成本卷积配置", "需成本卷积结果"],
    ["跨仓抵销", "仅展示数量金额干净桥接组合", "仓库贡献相对物料统一价", "不是仓库错账或调整建议"],
    ["本期收发钩稽", "逐项展示U8原值、U8剔除后、CAATS计入值及差异", "客户指定口径", "CAATS结果仅与U8剔除后口径比较"],
    ["剔除明细", "已记账、记入成本且属于四类移动的流水按收入/发出分别释放", "四类移动剔除明细工作簿", "按物料+月份汇总后分别等于01及03的剔除数量、金额"],
    ["CAATS/ITA边界", "CAATS保留结果表，ITA保留核对证据，四类移动流水另册", "职责分离", "05对应01，06对应03；剔除流水见配套明细工作簿"],
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
  const transitionMismatch = data.transition_checks.reduce((sum, row) => sum + row.quantity_mismatch + row.amount_mismatch, 0);
  const excludedBreakdownRows = data.excluded_movement_breakdown.reduce((sum, row) => sum + row.rows, 0);
  const checkValues = [
    ["物料月份行数", material.length, data.metrics.material_month_rows, null, 0, null, "收发存去除合计行"],
    ["仓库物料月份行数", warehouses.length, data.metrics.warehouse_material_month_rows, null, 0, null, "流水发生额组合"],
    ["流水仓库映射缺失数", data.metrics.missing_master.length, 0, null, 0, null, "应全部映射仓库档案"],
    ["剔除移动方式配置数", data.metrics.excluded_movement_categories.length, 4, null, 0, null, "应为调拨出入库及借用出入库四类"],
    ["剔除移动流水行数钩稽", excludedBreakdownRows, data.metrics.excluded_movement_rows, null, 0, null, "四类分项行数合计应等于剔除总行数"],
    ["月间期初期末不一致数", transitionMismatch, 0, null, 0, null, "数量+金额不一致项"],
    ["仓库档案非全月平均数", master.filter((row) => row["计价方式"] !== "全月平均法").length, 0, null, 0, null, "客户档案"],
    ["跨仓抵销明细行数", data.offset_rows.length, data.offset_summary.offset_rows, null, 0, null, "仅干净桥接组合"],
    ["缺分仓期初的可复算行数", warehouses.length, 0, null, 0, null, "已知限制，不代表数据错误"],
    ["连续完整月份面板行数", continuousMaterial.length, data.metrics.continuous_panel_rows, null, 0, null, "从各物料首次出现月至Final"],
    ["Final物料覆盖数", data.metrics.final_material_rows, new Set(continuousMaterial.map((row) => row.code)).size, null, 0, null, "每个物料Final仅保留一条"],
    ["金额滚转钩稽异常数", data.metrics.rollforward_review_rows, 0, null, 0, null, "期初+收入-发出-期末"],
    ["收入剔除明细数量钩稽", data.metrics.excluded_income_detail_quantity, data.metrics.excluded_income_quantity_bridge, null, quantityTolerance, null, "剔除明细工作簿01汇总应等于01及03剔除收入数量"],
    ["收入剔除明细金额钩稽", data.metrics.excluded_income_detail_amount, data.metrics.excluded_income_amount_bridge, null, amountTolerance, null, "剔除明细工作簿01汇总应等于01及03剔除收入金额"],
    ["发出剔除明细数量钩稽", data.metrics.excluded_issue_detail_quantity, data.metrics.excluded_issue_quantity_bridge, null, quantityTolerance, null, "剔除明细工作簿02汇总应等于01及03剔除发出数量"],
    ["发出剔除明细金额钩稽", data.metrics.excluded_issue_detail_amount, data.metrics.excluded_issue_bridge, null, amountTolerance, null, "剔除明细工作簿02汇总应等于01及03剔除发出金额"],
  ];
  const checkEnd = checkValues.length + 3;
  checkSheet.getRange(`A4:G${checkEnd}`).values = checkValues;
  for (let row = 4; row <= checkEnd; row += 1) {
    checkSheet.getRange(`D${row}`).formulas = [[`=B${row}-C${row}`]];
    checkSheet.getRange(`F${row}`).formulas = [[`=IF(ABS(D${row})<=E${row},"PASS","REVIEW")`]];
  }
  checkSheet.getRange(`A3:G${checkEnd}`).format.borders = thinBorder;
  checkSheet.getRange(`B4:E${checkEnd}`).format.numberFormat = numberFormat;
  checkSheet.getRange(`A1:A${checkEnd + 2}`).format.columnWidth = 40;
  checkSheet.getRange(`B1:F${checkEnd + 2}`).format.columnWidth = 16;
  checkSheet.getRange(`G1:G${checkEnd + 2}`).format.columnWidth = 44;
  applyStatusFormatting(checkSheet, `F4:F${checkEnd}`);

  if (!skipQa) {
    await saveInspection(workbook, "ita", [["01_核对摘要", "A1:H11"], ["03_月间衔接", "A1:J6"], ["10_检查", `A1:G${checkEnd}`]]);
    await renderSheets(workbook, "ita", [
      ["01_核对摘要", "A1:H11", 1.05],
      ["02_输入清单", `A1:F${inputValues.length + 1}`, 0.9],
      ["03_月间衔接", "A1:J6", 1.0],
      ["04_流水桥接", "A1:U22", 0.72],
      ["05_仓库物料月", "A1:AX22", 0.63],
      ["06_物料月桥接", "A1:Y22", 0.64],
      ["07_仓库组特殊单据", "A1:O22", 0.85],
      ["08_仓库档案", "A1:J22", 0.9],
      ["09_口径说明", `A1:D${itaScopeRows.length + 3}`, 1.0],
      ["10_检查", `A1:G${checkEnd}`, 1.1],
    ]);
  }
  const output = await SpreadsheetFile.exportXlsx(workbook);
  const path = `${outputDir}/U8存货发出计价_ITA核对钩稽数据.xlsx`;
  await output.save(path);
  return path;
}

async function buildExcludedWorkbook() {
  const workbook = Workbook.create();
  const incomeSheet = workbook.worksheets.add("01_收入剔除明细");
  const issueSheet = workbook.worksheets.add("02_发出剔除明细");
  const incomeDifferenceSheet = workbook.worksheets.add("03_收入金额差异明细");
  populateExcludedDetailSheet(incomeSheet, data.excluded_income_details, palette.incomeHeader);
  populateExcludedDetailSheet(issueSheet, data.excluded_issue_details, palette.issueHeader);
  const incomeDifferenceRows = material.filter((row) => Math.abs(row.u8_filtered_income_amount - row.filtered_ledger_ia) >= 0.005);
  populateIncomeDifferenceSheet(incomeDifferenceSheet, incomeDifferenceRows);
  if (!skipQa) {
    await saveInspection(workbook, "excluded", [["01_收入剔除明细", "A1:Q20"], ["02_发出剔除明细", "A1:Q20"], ["03_收入金额差异明细", "A1:V24"]]);
    await renderSheets(workbook, "excluded", [["01_收入剔除明细", "A1:Q24", 0.8], ["02_发出剔除明细", "A1:Q24", 0.8], ["03_收入金额差异明细", "A1:V24", 0.7]]);
  }
  const output = await SpreadsheetFile.exportXlsx(workbook);
  const path = `${outputDir}/U8存货发出计价_四类移动剔除明细.xlsx`;
  await output.save(path);
  return path;
}

try {
  let path;
  if (buildPart === "caats") path = await buildCaatsWorkbook();
  else if (buildPart === "ita") path = await buildItaWorkbook();
  else if (buildPart === "excluded") path = await buildExcludedWorkbook();
  else throw new Error(`未知U8_BUILD_PART：${buildPart}`);
  console.log(JSON.stringify({ buildPart, path, metrics: data.metrics, offsetSummary: data.offset_summary }, null, 2));
} catch (error) {
  console.error(JSON.stringify({
    buildPart,
    errorName: error?.name || "Error",
    errorMessage: error?.message || String(error),
    stackTail: String(error?.stack || "").split("\n").slice(-6),
  }, null, 2));
  process.exitCode = 1;
}
