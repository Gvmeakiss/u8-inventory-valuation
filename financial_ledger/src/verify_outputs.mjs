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
  excluded: `${outputDir}/U8存货发出计价_调拨出入库剔除明细.xlsx`,
};

const expectedSheets = {
  caats: ["01_物料月维度明细", "02_物料月维度结果汇总", "03_连续滚算明细", "04_连续滚算汇总", "05_月度独立差异明细（对应01）", "06_连续滚算差异明细（对应03）", "07_跨仓抵销"],
  ita: ["01_核对摘要", "02_输入清单", "03_月间衔接", "04_流水桥接", "05_仓库物料月", "06_物料月桥接", "07_仓库组特殊单据", "08_仓库档案", "09_口径说明", "10_检查"],
  excluded: ["01_收入剔除明细", "02_发出剔除明细", "03_收入金额差异明细"],
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
    const independentDifferenceSheet = workbook.worksheets.getItem("05_月度独立差异明细（对应01）");
    const continuousDifferenceSheet = workbook.worksheets.getItem("06_连续滚算差异明细（对应03）");
    const independentDifferenceRows = independentDifferenceSheet.getUsedRange(true).values.length - 1;
    const continuousDifferenceRows = continuousDifferenceSheet.getUsedRange(true).values.length - 1;
    const reviewRows = Number(workbook.worksheets.getItem("02_物料月维度结果汇总").getRange("B6").values[0][0]);
    const independentDifferenceHeaders = independentDifferenceSheet.getRange("A1:I1").values[0];
    const continuousDifferenceHeaders = continuousDifferenceSheet.getRange("A1:M1").values[0];
    const independentSheet = workbook.worksheets.getItem("01_物料月维度明细");
    const independentHeaders = independentSheet.getRange("A1:AW1").values[0];
    const continuousSheet = workbook.worksheets.getItem("03_连续滚算明细");
    const continuousHeaders = continuousSheet.getRange("A1:BD1").values[0];
    if (continuousRows < independentRows) throw new Error(`完整连续面板不应少于原始月度明细：独立=${independentRows}，连续=${continuousRows}`);
    if (independentDifferenceRows !== reviewRows) throw new Error(`05月度独立差异明细与01 REVIEW数量不一致：明细=${independentDifferenceRows}，REVIEW=${reviewRows}`);
    if (independentDifferenceHeaders.some((value) => String(value || "").includes("绝对差异"))) throw new Error("05月度独立差异明细仍包含绝对差异字段");
    if (continuousDifferenceHeaders.some((value) => String(value || "").includes("绝对差异"))) throw new Error("06连续滚算差异明细不应包含绝对差异字段");
    if (continuousDifferenceHeaders[7] !== "本期发出金额差异（U8剔除后-连续CAATS）" || continuousDifferenceHeaders[8] !== "本期期末结存差异（U8剔除后-连续CAATS）") throw new Error("06连续滚算差异明细标题未明确对应03口径");
    if (independentHeaders.length !== 49) throw new Error(`01明细应至AW列，实际列数=${independentHeaders.length}`);
    if (continuousHeaders.length !== 56) throw new Error(`03明细应至BD列，实际列数=${continuousHeaders.length}`);
    if (independentHeaders[12] !== "用友U8月度收入数量" || independentHeaders[13] !== "剔除收入数量" || independentHeaders[14] !== "用友U8月度收入数量（剔除后）" || independentHeaders[15] !== "CAATS计入收入数量") throw new Error("01明细收入数量字段未形成U8原值-剔除-剔除后-CAATS桥接");
    if (independentHeaders[18] !== "用友U8月度收入金额" || independentHeaders[19] !== "剔除收入金额" || independentHeaders[21] !== "用友U8月度收入金额（剔除后）") throw new Error("01明细收入金额字段未明确剔除桥");
    if (independentHeaders[24] !== "用友U8月度发出数量" || independentHeaders[25] !== "剔除发出数量" || independentHeaders[26] !== "用友U8月度发出数量（剔除后）") throw new Error("01明细发出数量字段未明确剔除桥");
    if (independentHeaders[30] !== "用友U8月度发出金额" || independentHeaders[31] !== "剔除发出金额" || independentHeaders[33] !== "用友U8月度发出金额（剔除后）" || independentHeaders[34] !== "CAATS发出金额") throw new Error("01明细发出金额字段未形成剔除桥");
    if (independentHeaders[38] !== "用友U8月度结存金额" || independentHeaders[41] !== "用友U8月度结存金额（剔除后）" || independentHeaders[43] !== "结存金额差异（U8剔除后-CAATS）") throw new Error("01明细结存字段未按U8剔除后-CAATS展示");
    if (continuousHeaders[49] !== "记录类型" || continuousHeaders[53] !== "期初承接结存差异（U8-连续CAATS）" || continuousHeaders[54] !== "金额滚转钩稽差异" || continuousHeaders[55] !== "金额滚转钩稽状态") throw new Error("03明细完整月份及滚转钩稽字段不正确");
    if (continuousHeaders[43] !== "结存金额差异（U8剔除后-连续CAATS）") throw new Error("03明细结存差异方向不正确");
    const allCaatsHeaders = [independentHeaders, continuousHeaders, independentDifferenceHeaders, continuousDifferenceHeaders, workbook.worksheets.getItem("02_物料月维度结果汇总").getUsedRange(true).values.flat(), workbook.worksheets.getItem("04_连续滚算汇总").getUsedRange(true).values.flat(), workbook.worksheets.getItem("07_跨仓抵销").getRange("A1:Q9").values.flat()].flat();
    if (allCaatsHeaders.some((value) => String(value || "").includes("绝对值") || String(value || "").includes("绝对差异"))) throw new Error("CAATS结果表仍存在面向审阅人的绝对值差异展示");
    const key = (code, month) => `${String(code)}|${String(month)}`;
    const near = (left, right) => Math.abs(Number(left || 0) - Number(right || 0)) <= 0.000001;
    const independentValues = independentSheet.getUsedRange(true).values.slice(1);
    const continuousValues = continuousSheet.getUsedRange(true).values.slice(1);
    for (const [index, label] of [[13, "剔除收入数量"], [19, "剔除收入金额"], [25, "剔除发出数量"], [31, "剔除发出金额"]]) {
      const independentTotal = independentValues.reduce((sum, row) => sum + Number(row[index] || 0), 0);
      const continuousTotal = continuousValues.reduce((sum, row) => sum + Number(row[index] || 0), 0);
      if (!near(independentTotal, continuousTotal)) throw new Error(`01与03${label}合计不一致`);
    }
    const independentByKey = new Map(independentValues.map((row) => [key(row[0], row[1]), row]));
    const continuousByKey = new Map(continuousValues.map((row) => [key(row[0], row[1]), row]));
    const independentDifferenceValues = independentDifferenceSheet.getUsedRange(true).values.slice(1);
    const continuousDifferenceValues = continuousDifferenceSheet.getUsedRange(true).values.slice(1);
    const independentSeen = new Set();
    for (const row of independentDifferenceValues) {
      const rowKey = key(row[1], row[0]);
      const source = independentByKey.get(rowKey);
      if (!source || source[47] !== "REVIEW") throw new Error(`05存在非01 REVIEW项目：${rowKey}`);
      if (independentSeen.has(rowKey)) throw new Error(`05存在重复物料月份：${rowKey}`);
      independentSeen.add(rowKey);
      if (!near(row[4], source[33]) || !near(row[5], source[34]) || !near(row[6], source[35]) || row[7] !== source[47] || row[8] !== source[48]) throw new Error(`05未与01逐字段钩稽：${rowKey}`);
    }
    const independentReviewKeys = new Set(independentValues.filter((row) => row[47] === "REVIEW").map((row) => key(row[0], row[1])));
    if (independentSeen.size !== independentReviewKeys.size || [...independentReviewKeys].some((rowKey) => !independentSeen.has(rowKey))) throw new Error("05未完整覆盖01全部REVIEW项目");
    const continuousSeen = new Set();
    for (const row of continuousDifferenceValues) {
      const rowKey = key(row[1], row[0]);
      const source = continuousByKey.get(rowKey);
      if (!source || source[47] !== "REVIEW") throw new Error(`06存在非03 REVIEW项目：${rowKey}`);
      if (continuousSeen.has(rowKey)) throw new Error(`06存在重复物料月份：${rowKey}`);
      continuousSeen.add(rowKey);
      const numberPairs = [[5, 53], [6, 23], [7, 35], [8, 43], [9, 54]];
      if (numberPairs.some(([targetIndex, sourceIndex]) => !near(row[targetIndex], source[sourceIndex])) || row[4] !== source[49] || row[10] !== source[55] || row[11] !== source[47] || row[12] !== source[52]) throw new Error(`06未与03逐字段钩稽：${rowKey}`);
    }
    const continuousReviewKeys = new Set(continuousValues.filter((row) => row[47] === "REVIEW").map((row) => key(row[0], row[1])));
    if (continuousDifferenceRows !== continuousReviewKeys.size || continuousSeen.size !== continuousReviewKeys.size || [...continuousReviewKeys].some((rowKey) => !continuousSeen.has(rowKey))) throw new Error(`06未完整覆盖03全部REVIEW项目：明细=${continuousDifferenceRows}，03 REVIEW=${continuousReviewKeys.size}`);
    const independentSample = independentValues.slice(0, 100);
    for (const row of independentSample) {
      if (!near(Number(row[12] || 0) - Number(row[13] || 0), row[14])) throw new Error("01收入数量未按U8原值-剔除=剔除后钩稽");
      if (!near(Number(row[14] || 0) - Number(row[15] || 0), row[16])) throw new Error("01收入数量差异未按U8剔除后-CAATS钩稽");
      if (!near(Number(row[18] || 0) - Number(row[19] || 0), row[21])) throw new Error("01收入金额未按U8原值-剔除=剔除后钩稽");
      if (!near(Number(row[21] || 0) - Number(row[22] || 0), row[23])) throw new Error("01收入金额差异未按U8剔除后-CAATS钩稽");
      if (!near(Number(row[24] || 0) - Number(row[25] || 0), row[26])) throw new Error("01发出数量未按U8原值-剔除=剔除后钩稽");
      if (!near(Number(row[26] || 0) - Number(row[27] || 0), row[28])) throw new Error("01发出数量差异未按U8剔除后-CAATS钩稽");
      if (!near(Number(row[30] || 0) - Number(row[31] || 0), row[33])) throw new Error("01发出金额未按U8原值-剔除=剔除后钩稽");
      if (!near(Number(row[33] || 0) - Number(row[34] || 0), row[35])) throw new Error("01发出金额差异未按U8剔除后-CAATS钩稽");
      if (!near(Number(row[41] || 0) - Number(row[42] || 0), row[43])) throw new Error("01结存金额未按U8剔除后-CAATS钩稽");
    }
    const periodList = [...new Set(continuousValues.map((row) => String(row[1])))].sort();
    const finalMonth = periodList.at(-1);
    const codes = new Set(continuousValues.map((row) => String(row[0])));
    const finalRows = continuousValues.filter((row) => String(row[1]) === finalMonth);
    if (finalRows.length !== codes.size) throw new Error(`Final月份未覆盖全部物料：Final=${finalRows.length}，物料=${codes.size}`);
    const monthsByCode = new Map();
    for (const row of continuousValues) {
      const code = String(row[0]);
      if (!monthsByCode.has(code)) monthsByCode.set(code, []);
      monthsByCode.get(code).push(String(row[1]));
      const caatsBegin = Number(row[11] || 0);
      const u8Begin = Number(row[51] || 0);
      const beginDifference = Number(row[53] || 0);
      const rollforward = Number(row[54] || 0);
      if (!near(u8Begin - caatsBegin, beginDifference)) throw new Error(`03期初承接差异未钩稽：${code}|${row[1]}`);
      if (!near(beginDifference + Number(row[23] || 0) - Number(row[35] || 0) - Number(row[43] || 0), rollforward)) throw new Error(`03四段式金额滚转未钩稽：${code}|${row[1]}`);
      if (Math.abs(rollforward) > 0.01 || row[55] !== "PASS") throw new Error(`03金额滚转存在异常：${code}|${row[1]}`);
      if (row[49] === "补齐无收发月份") {
        const movement = [...row.slice(12, 24), ...row.slice(24, 36)];
        if (movement.some((value) => Math.abs(Number(value || 0)) > 0.000001)) throw new Error(`补齐月份存在非零收发：${code}|${row[1]}`);
      }
    }
    for (const [code, months] of monthsByCode.entries()) {
      const uniqueMonths = [...new Set(months)].sort();
      const expected = periodList.slice(periodList.indexOf(uniqueMonths[0]));
      if (JSON.stringify(uniqueMonths) !== JSON.stringify(expected)) throw new Error(`物料月份面板不连续：${code}`);
    }
    for (const row of continuousValues.slice(0, 100)) {
      if (!near(Number(row[12] || 0) - Number(row[13] || 0), row[14])) throw new Error("03收入数量未按U8原值-剔除=剔除后钩稽");
      if (!near(Number(row[18] || 0) - Number(row[19] || 0), row[21])) throw new Error("03收入金额未按U8原值-剔除=剔除后钩稽");
      if (!near(Number(row[24] || 0) - Number(row[25] || 0), row[26])) throw new Error("03发出数量未按U8原值-剔除=剔除后钩稽");
      if (!near(Number(row[30] || 0) - Number(row[31] || 0), row[33])) throw new Error("03发出金额未按U8原值-剔除=剔除后钩稽");
      if (!near(Number(row[41] || 0) - Number(row[42] || 0), row[43])) throw new Error("03结存金额未按U8剔除后-连续CAATS钩稽");
    }
    const continuousSummary = workbook.worksheets.getItem("04_连续滚算汇总");
    const finalEndDifference = finalRows.reduce((sum, row) => sum + Number(row[43] || 0), 0);
    const finalSyntheticRows = finalRows.filter((row) => row[49] === "补齐无收发月份").length;
    if (Number(continuousSummary.getRange("B4").values[0][0]) !== continuousRows) throw new Error("04完整连续面板行数未与03钩稽");
    if (Number(continuousSummary.getRange("B6").values[0][0]) !== finalRows.length) throw new Error("04 Final物料数未与03钩稽");
    if (Number(continuousSummary.getRange("B7").values[0][0]) !== finalSyntheticRows) throw new Error("04 Final补齐记录数未与03钩稽");
    if (!near(continuousSummary.getRange("B10").values[0][0], finalEndDifference)) throw new Error("04 Final结存金额差异未与03 Final行钩稽");
    if (Number(continuousSummary.getRange("B11").values[0][0]) !== 0) throw new Error("04金额滚转钩稽异常数不为0");
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
  if (kind === "excluded") {
    const incomeExcludedSheet = workbook.worksheets.getItem("01_收入剔除明细");
    const issueExcludedSheet = workbook.worksheets.getItem("02_发出剔除明细");
    const incomeDifferenceSheet = workbook.worksheets.getItem("03_收入金额差异明细");
    const incomeHeaders = incomeExcludedSheet.getRange("A1:Q1").values[0];
    const issueHeaders = issueExcludedSheet.getRange("A1:Q1").values[0];
    if (incomeHeaders[12] !== "剔除数量" || incomeHeaders[14] !== "剔除金额" || issueHeaders[12] !== "剔除数量" || issueHeaders[14] !== "剔除金额") throw new Error("剔除明细字段结构不正确");
    const itaWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(files.ita));
    const bridgeValues = itaWorkbook.worksheets.getItem("06_物料月桥接").getUsedRange(true).values.slice(1);
    const incomeExcludedValues = incomeExcludedSheet.getUsedRange(true).values.slice(1);
    const issueExcludedValues = issueExcludedSheet.getUsedRange(true).values.slice(1);
    const key = (month, code) => `${String(month)}|${String(code)}`;
    const aggregate = (rows) => {
      const result = new Map();
      for (const row of rows) {
        const rowKey = key(row[0], row[7]);
        const current = result.get(rowKey) || { quantity: 0, amount: 0 };
        current.quantity += Number(row[12] || 0);
        current.amount += Number(row[14] || 0);
        result.set(rowKey, current);
      }
      return result;
    };
    const incomeByKey = aggregate(incomeExcludedValues);
    const issueByKey = aggregate(issueExcludedValues);
    let bridgeIncomeQuantity = 0; let bridgeIncomeAmount = 0; let bridgeIssueQuantity = 0; let bridgeIssueAmount = 0;
    for (const row of bridgeValues) {
      const rowKey = key(row[0], row[1]);
      const incomeQuantity = Number(row[3] || 0) - Number(row[4] || 0);
      const incomeAmount = Number(row[7] || 0) - Number(row[8] || 0);
      const issueQuantity = Number(row[11] || 0) - Number(row[12] || 0);
      const issueAmount = Number(row[15] || 0) - Number(row[16] || 0);
      const incomeDetail = incomeByKey.get(rowKey) || { quantity: 0, amount: 0 };
      const issueDetail = issueByKey.get(rowKey) || { quantity: 0, amount: 0 };
      if (Math.abs(incomeQuantity - incomeDetail.quantity) > 0.000001 || Math.abs(incomeAmount - incomeDetail.amount) > 0.01) throw new Error(`收入剔除明细未与物料月桥接钩稽：${rowKey}`);
      if (Math.abs(issueQuantity - issueDetail.quantity) > 0.000001 || Math.abs(issueAmount - issueDetail.amount) > 0.01) throw new Error(`发出剔除明细未与物料月桥接钩稽：${rowKey}`);
      bridgeIncomeQuantity += incomeQuantity; bridgeIncomeAmount += incomeAmount; bridgeIssueQuantity += issueQuantity; bridgeIssueAmount += issueAmount;
    }
    const sumColumn = (rows, index) => rows.reduce((sum, row) => sum + Number(row[index] || 0), 0);
    if (Math.abs(bridgeIncomeQuantity - sumColumn(incomeExcludedValues, 12)) > 0.000001 || Math.abs(bridgeIncomeAmount - sumColumn(incomeExcludedValues, 14)) > 0.01) throw new Error("收入剔除明细合计未与物料月桥接钩稽");
    if (Math.abs(bridgeIssueQuantity - sumColumn(issueExcludedValues, 12)) > 0.000001 || Math.abs(bridgeIssueAmount - sumColumn(issueExcludedValues, 14)) > 0.01) throw new Error("发出剔除明细合计未与物料月桥接钩稽");
    const differenceHeaders = incomeDifferenceSheet.getRange("A1:V1").values[0];
    if (differenceHeaders[17] !== "收入金额差异（U8剔除后-CAATS）" || differenceHeaders[18] !== "差异状态") throw new Error("收入金额差异明细字段结构不正确");
    const differenceValues = incomeDifferenceSheet.getUsedRange(true).values.slice(1);
    const caatsWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(files.caats));
    const independentValues = caatsWorkbook.worksheets.getItem("01_物料月维度明细").getUsedRange(true).values.slice(1);
    const continuousValues = caatsWorkbook.worksheets.getItem("03_连续滚算明细").getUsedRange(true).values.slice(1);
    const independentByKey = new Map(independentValues.map((row) => [key(row[1], row[0]), row]));
    const continuousByKey = new Map(continuousValues.map((row) => [key(row[1], row[0]), row]));
    const expectedDifferenceRows = bridgeValues.filter((row) => Math.abs(Number(row[8] || 0) - Number(row[9] || 0)) >= 0.005);
    if (differenceValues.length !== expectedDifferenceRows.length) throw new Error("收入金额差异明细行数未与ITA物料月桥接钩稽");
    const expectedByKey = new Map(expectedDifferenceRows.map((row) => [key(row[0], row[1]), {
      quantity: Number(row[4] || 0) - Number(row[5] || 0),
      amount: Number(row[8] || 0) - Number(row[9] || 0),
    }]));
    let differenceAmountTotal = 0;
    for (const row of differenceValues) {
      const rowKey = key(row[0], row[1]);
      const expected = expectedByKey.get(rowKey);
      if (!expected) throw new Error(`收入金额差异明细存在非预期物料月份：${rowKey}`);
      if (Math.abs(Number(row[9] || 0) - expected.quantity) > 0.000001) throw new Error(`收入金额差异明细数量未钩稽：${rowKey}`);
      if (Math.abs(Number(row[17] || 0) - expected.amount) > 0.01) throw new Error(`收入金额差异明细金额未钩稽：${rowKey}`);
      const independentSource = independentByKey.get(rowKey);
      const continuousSource = continuousByKey.get(rowKey);
      if (!independentSource || Math.abs(Number(row[17] || 0) - Number(independentSource[23] || 0)) > 0.000001) throw new Error(`收入金额差异明细未与01逐项钩稽：${rowKey}`);
      if (!continuousSource || Math.abs(Number(row[17] || 0) - Number(continuousSource[23] || 0)) > 0.000001) throw new Error(`收入金额差异明细未与03逐项钩稽：${rowKey}`);
      const expectedStatus = Math.abs(Math.round(Number(row[17] || 0) * 1e10) / 1e10) > 0.01 ? "REVIEW" : "容差内";
      if (row[18] !== expectedStatus) throw new Error(`收入金额差异明细状态不正确：${rowKey}`);
      differenceAmountTotal += Number(row[17] || 0);
    }
    const allIncomeDifference = bridgeValues.reduce((sum, row) => sum + Number(row[8] || 0) - Number(row[9] || 0), 0);
    if (Math.abs(differenceAmountTotal - allIncomeDifference) > 0.01) throw new Error("收入金额差异明细净额未与ITA物料月桥接合计钩稽");
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
    : kind === "ita"
      ? await workbook.inspect({ kind: "table", range: "10_检查!A1:G19", include: "values,formulas", tableMaxRows: 22, tableMaxCols: 8, maxChars: 18000 })
      : await workbook.inspect({ kind: "table", range: "03_收入金额差异明细!A1:V24", include: "values,formulas", tableMaxRows: 24, tableMaxCols: 22, maxChars: 22000 });
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
result.excluded = await verifyWorkbook("excluded", files.excluded);

const qaFiles = await fs.readdir(qaDir);
for (const [kind, sheets] of Object.entries(expectedSheets)) {
  const missing = sheets.filter((sheet) => !qaFiles.some((file) => file.startsWith(`${kind}_${sheet}`) && file.endsWith(".png")));
  if (missing.length) throw new Error(`${kind}缺少视觉检查图：${missing.join(",")}`);
}

await fs.writeFile(`${qaDir}/verification.json`, JSON.stringify(result, null, 2));
console.log(JSON.stringify({
  caats: { path: result.caats.path, size: result.caats.size, sha256: result.caats.sha256, sheets: result.caats.sheets, formulaErrors: 0 },
  ita: { path: result.ita.path, size: result.ita.size, sha256: result.ita.sha256, sheets: result.ita.sheets, formulaErrors: 0 },
  excluded: { path: result.excluded.path, size: result.excluded.size, sha256: result.excluded.sha256, sheets: result.excluded.sheets, formulaErrors: 0 },
}, null, 2));
