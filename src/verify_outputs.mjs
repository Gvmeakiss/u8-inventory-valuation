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
    if (independentRows !== continuousRows) throw new Error(`连续滚算明细行数不一致：独立=${independentRows}，连续=${continuousRows}`);
    if (differenceRows !== reviewRows) throw new Error(`存在差异明细与REVIEW数量不一致：明细=${differenceRows}，REVIEW=${reviewRows}`);
    if (differenceHeaders.includes("绝对差异")) throw new Error("存在差异明细仍包含绝对差异字段");
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
