import fs from "node:fs/promises";
import crypto from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const execFileAsync = promisify(execFile);
const outputDir = process.env.U8_OUTPUT_DIR;
const qaDir = process.env.U8_QA_DIR;
if (!outputDir || !qaDir) throw new Error("缺少U8_OUTPUT_DIR/U8_QA_DIR");

const files = {
  caats: `${outputDir}/U8存货发出计价_CAATS审计结果表.xlsx`,
  ita: `${outputDir}/U8存货发出计价_ITA核对钩稽数据.xlsx`,
};

const expectedSheets = {
  caats: ["01_物料月维度明细", "02_物料月维度结果汇总", "03_连续滚算明细", "04_连续滚算汇总", "05_存在差异明细", "06_跨仓抵销"],
  ita: ["01_核对摘要", "02_输入清单", "03_月间衔接", "04_流水桥接", "05_仓库物料月", "06_物料月桥接", "07_仓库组特殊单据", "08_仓库档案", "09_口径说明", "10_检查"],
};

async function sha256(path) {
  const buffer = await fs.readFile(path);
  return crypto.createHash("sha256").update(buffer).digest("hex");
}

async function verifyWorkbook(kind, path) {
  const stat = await fs.stat(path);
  if (stat.size < 10000) throw new Error(`${kind}工作簿文件过小：${stat.size}`);
  const { stdout } = await execFileAsync("unzip", ["-t", path]);
  if (!stdout.includes("No errors detected")) throw new Error(`${kind}工作簿ZIP检查失败`);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
  const sheetInspection = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 12000 });
  const actualSheets = sheetInspection.ndjson.split("\n").filter(Boolean).map((line) => JSON.parse(line).name);
  if (JSON.stringify(actualSheets) !== JSON.stringify(expectedSheets[kind])) {
    throw new Error(`${kind}工作表不符合预期。实际=${JSON.stringify(actualSheets)}`);
  }
  if (kind === "caats" && actualSheets.some((name) => ["04_流水桥接", "08_仓库档案", "09_口径说明", "10_检查", "05_仓库物料月"].includes(name))) {
    throw new Error("CAATS结果表混入ITA支持性Sheet");
  }
  if (kind === "caats") {
    const independentRows = workbook.worksheets.getItem("01_物料月维度明细").getUsedRange(true).values.length - 1;
    const continuousRows = workbook.worksheets.getItem("03_连续滚算明细").getUsedRange(true).values.length - 1;
    const differenceRows = workbook.worksheets.getItem("05_存在差异明细").getUsedRange(true).values.length - 1;
    const reviewRows = Number(workbook.worksheets.getItem("02_物料月维度结果汇总").getRange("B6").values[0][0]);
    const differenceHeaders = workbook.worksheets.getItem("05_存在差异明细").getRange("A1:I1").values[0];
    const independentSheet = workbook.worksheets.getItem("01_物料月维度明细");
    const independentHeaders = independentSheet.getRange("A1:AT1").values[0];
    const continuousSheet = workbook.worksheets.getItem("03_连续滚算明细");
    const continuousHeaders = continuousSheet.getRange("A1:AY1").values[0];
    if (independentRows !== continuousRows) throw new Error(`连续滚算明细行数不一致：独立=${independentRows}，连续=${continuousRows}`);
    if (differenceRows !== reviewRows) throw new Error(`存在差异明细与REVIEW数量不一致：明细=${differenceRows}，REVIEW=${reviewRows}`);
    if (differenceHeaders.includes("绝对差异")) throw new Error("存在差异明细仍包含绝对差异字段");
    if (independentHeaders.length !== 46) throw new Error(`01明细应至AT列，实际列数=${independentHeaders.length}`);
    if (continuousHeaders.length !== 51) throw new Error(`03明细应至AY列，实际列数=${continuousHeaders.length}`);
    if (independentHeaders[12] !== "用友U8月度收入数量" || independentHeaders[13] !== "用友U8月度收入数量（剔除后）" || independentHeaders[14] !== "CAATS计入收入数量") throw new Error("01明细收入字段未形成U8原值-剔除后-CAATS桥接");
    if (independentHeaders[27] !== "用友U8月度发出金额" || independentHeaders[29] !== "用友U8月度发出金额（剔除后）" || independentHeaders[30] !== "CAATS发出金额") throw new Error("01明细发出字段未形成U8原值-剔除后-CAATS桥接");
    if (independentHeaders[34] !== "用友U8月度结存金额" || independentHeaders[37] !== "用友U8月度结存金额（剔除后）" || independentHeaders[39] !== "结存金额差异（U8剔除后-CAATS）") throw new Error("01明细结存字段未按U8剔除后-CAATS展示");
    if (continuousHeaders[50] !== "期初金额差异(U8-连续CAATS)") throw new Error(`03明细AY列标题不正确：${continuousHeaders[50]}`);
    if (continuousHeaders[39] !== "结存金额差异（U8剔除后-连续CAATS）") throw new Error("03明细结存差异方向不正确");
    const independentSample = independentSheet.getRange("M2:AN101").values;
    for (const row of independentSample) {
      if (Math.abs((Number(row[1] || 0) - Number(row[2] || 0)) - Number(row[3] || 0)) > 0.000001) throw new Error("01明细收入数量未按U8剔除后-CAATS钩稽");
      if (Math.abs((Number(row[7] || 0) - Number(row[8] || 0)) - Number(row[9] || 0)) > 0.000001) throw new Error("01明细收入金额未按U8剔除后-CAATS钩稽");
      if (Math.abs((Number(row[11] || 0) - Number(row[12] || 0)) - Number(row[13] || 0)) > 0.000001) throw new Error("01明细发出数量未按U8剔除后-CAATS钩稽");
      if (Math.abs((Number(row[17] || 0) - Number(row[18] || 0)) - Number(row[19] || 0)) > 0.000001) throw new Error("01明细发出金额未按U8剔除后-CAATS钩稽");
      if (Math.abs((Number(row[25] || 0) - Number(row[26] || 0)) - Number(row[27] || 0)) > 0.000001) throw new Error("01明细结存金额未按U8剔除后-CAATS钩稽");
    }
    const continuousBegin = continuousSheet.getRange("L2:L101").values;
    const continuousAudit = continuousSheet.getRange("AV2:AY101").values;
    for (let index = 0; index < continuousBegin.length; index += 1) {
      const caatsBegin = Number(continuousBegin[index][0] || 0);
      const u8Begin = Number(continuousAudit[index][0] || 0);
      const beginDifference = Number(continuousAudit[index][3] || 0);
      if (Math.abs((u8Begin - caatsBegin) - beginDifference) > 0.000001) throw new Error("03明细AY期初金额差异未按U8-连续CAATS钩稽");
    }
    const continuousEndingSample = continuousSheet.getRange("M2:AN101").values;
    for (const row of continuousEndingSample) {
      if (Math.abs((Number(row[1] || 0) - Number(row[2] || 0)) - Number(row[3] || 0)) > 0.000001) throw new Error("03明细收入数量未按U8剔除后-CAATS钩稽");
      if (Math.abs((Number(row[7] || 0) - Number(row[8] || 0)) - Number(row[9] || 0)) > 0.000001) throw new Error("03明细收入金额未按U8剔除后-CAATS钩稽");
      if (Math.abs((Number(row[11] || 0) - Number(row[12] || 0)) - Number(row[13] || 0)) > 0.000001) throw new Error("03明细发出数量未按U8剔除后-CAATS钩稽");
      if (Math.abs((Number(row[17] || 0) - Number(row[18] || 0)) - Number(row[19] || 0)) > 0.000001) throw new Error("03明细发出金额未按U8剔除后-连续CAATS钩稽");
      if (Math.abs((Number(row[25] || 0) - Number(row[26] || 0)) - Number(row[27] || 0)) > 0.000001) throw new Error("03明细结存金额未按U8剔除后-连续CAATS钩稽");
    }
  }
  if (kind === "ita") {
    const bridgeSheet = workbook.worksheets.getItem("06_物料月桥接");
    const bridgeHeaders = bridgeSheet.getRange("A1:Y1").values[0];
    if (bridgeHeaders[16] !== "U8月表发出金额（剔除后）" || bridgeHeaders[20] !== "重计价差异（筛选后-CAATS）" || bridgeHeaders[21] !== "总差异（U8剔除后-CAATS）") throw new Error("ITA物料月桥接未按U8剔除后与重计算差异拆分");
    const bridgeSample = bridgeSheet.getRange("Q2:V101").values;
    for (const row of bridgeSample) {
      const u8Amount = Number(row[0] || 0);
      const filteredAmount = Number(row[1] || 0);
      const scopeDifference = Number(row[2] || 0);
      const caatsAmount = Number(row[3] || 0);
      const repricingDifference = Number(row[4] || 0);
      const totalDifference = Number(row[5] || 0);
      if (Math.abs((u8Amount - filteredAmount) - scopeDifference) > 0.000001) throw new Error("ITA筛选范围差异未钩稽");
      if (Math.abs((filteredAmount - caatsAmount) - repricingDifference) > 0.000001) throw new Error("ITA可比口径重计算差异未钩稽");
      if (Math.abs((scopeDifference + repricingDifference) - totalDifference) > 0.000001) throw new Error("ITA总差异未与两段桥接钩稽");
    }
  }
  const errorInspection = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!",
    options: { useRegex: true, maxResults: 500 },
    summary: `${kind} final formula error scan`,
  });
  const errorLines = errorInspection.ndjson.split("\n").filter((line) => line.includes('"kind":"match"'));
  if (errorLines.length) throw new Error(`${kind}存在公式错误：${errorLines.slice(0, 5).join("\n")}`);
  const keyInspection = kind === "caats"
    ? await workbook.inspect({ kind: "table", range: "02_物料月维度结果汇总!A1:J18", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 12, maxChars: 18000 })
    : await workbook.inspect({ kind: "table", range: "10_检查!A1:G12", include: "values,formulas", tableMaxRows: 15, tableMaxCols: 8, maxChars: 14000 });
  return {
    path,
    size: stat.size,
    sha256: await sha256(path),
    sheets: actualSheets,
    zip: "PASS",
    formulaErrors: 0,
    keyInspection: keyInspection.ndjson,
  };
}

const result = {};
result.caats = await verifyWorkbook("caats", files.caats);
result.ita = await verifyWorkbook("ita", files.ita);

const qaFiles = await fs.readdir(qaDir);
for (const [kind, sheets] of Object.entries(expectedSheets)) {
  const missing = sheets.filter((sheet) => !qaFiles.some((file) => file.startsWith(`${kind}_${sheet}`) && file.endsWith(".png")));
  if (missing.length) throw new Error(`${kind}缺少视觉检查图：${missing.join(",")}`);
}

await fs.writeFile(`${qaDir}/verification.json`, JSON.stringify(result, null, 2));
console.log(JSON.stringify({
  caats: { path: result.caats.path, size: result.caats.size, sha256: result.caats.sha256, sheets: result.caats.sheets, formulaErrors: 0 },
  ita: { path: result.ita.path, size: result.ita.size, sha256: result.ita.sha256, sheets: result.ita.sheets, formulaErrors: 0 },
}, null, 2));
