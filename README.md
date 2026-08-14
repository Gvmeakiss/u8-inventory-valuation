# U8 存货发出计价工具 📦

> 将 U8 导出、仓库档案与 CAATS 模板处理为「CAATS 审计结果表」与「ITA 核对钩稽数据」两份职责分离的工作簿，支持月度独立重算、连续滚算与跨仓抵销。

[![Language](https://img.shields.io/badge/language-Python%203.8%2B%20%7C%20Node.js%2018%2B-blue)](https://github.com/Gvmeakiss/u8-inventory-valuation) [![License](https://img.shields.io/badge/license-MIT-green)](https://github.com/Gvmeakiss/u8-inventory-valuation/blob/main/LICENSE) [![Domain](https://img.shields.io/badge/domain-Audit%20Analytics-orange)](https://github.com/Gvmeakiss/u8-inventory-valuation)

## 📌 项目简介

针对 U8 存货发出计价的审计复核，本工具把原始 U8 导出（出入库流水账、收发存汇总）、仓库档案与 CAATS 交付模板，转换为两份职责分离的正式工作簿。**CAATS 审计结果表**仅含财务审计需要审阅的重算结果、差异、连续滚算与跨仓抵销；**ITA 核对钩稽数据**保存输入文件清单、月间衔接、流水桥接、仓库/仓库组拆分、配置档案与检查结果等底层核对证据，满足职责分离与可追溯要求。

## ✨ 功能特性

- **XLS/XLSX 标准化与字段校验**：`prepare_data.py` 通过 LibreOffice 将 `.xls` 转为 `.xlsx`，用 `required_columns()` 校验必需字段，保留原始值并另行生成规范化字段。
- **已记账成本流水口径**：按 `记账人非空` 且 `仓库档案.记入成本=是` 筛选成本流水（见 `config/project.json` 的 `posted_rule` / `cost_rule`）。
- **月度独立重算**：基于移动流水、月度收发存汇总与仓库档案，按移动加权平均口径独立重算发出计价。
- **连续滚算**：从首月 U8 期初开始，后续月份采用上月 CAATS 期末衔接（见 `continuous_rollforward_note`）。
- **跨仓抵销**：识别并处理跨仓调拨的抵销。
- **职责分离双工作簿**：由 `build_workbooks.mjs` 借助 `@oai/artifact-tool` 生成 CAATS 结果表与 ITA 钩稽数据两份工作簿。
- **输入留痕与校验**：对输入文件计算 SHA-256（`sha256()`），`verify_outputs.mjs` 做结果、公式错误、结构与视觉检查，`tests/test_calculations.py` 覆盖核心公式。

## 📂 目录结构

```text
u8-inventory-valuation/
├── README.md
├── LICENSE
├── .gitignore
├── run_pipeline.sh              # 一键执行入口（zsh）
├── config/
│   └── project.json             # 期间、特殊单据、容差、口径配置
├── src/
│   ├── prepare_data.py          # XLS→XLSX 转换、字段校验、数据标准化与计算
│   ├── build_workbooks.mjs      # 使用 artifact-tool 生成两份 Excel
│   └── verify_outputs.mjs       # 结果/公式错误/结构/视觉检查
└── tests/
    └── test_calculations.py     # 核心公式单元测试
```

运行时原始业务文件（不在仓库内，由使用者提供）示例：

```text
CAATS交付模板.xlsx
仓库管理-出入库流水账-2026.01-06.xlsx
收发存汇总/*.XLS
仓库档案.XLS
```

## 🔧 环境要求

- **Python** ≥ 3.8（依赖 `openpyxl>=3.1`，见 `requirements.txt`）。
- **Node.js** ≥ 18（提供 `@oai/artifact-tool`，供 `build_workbooks.mjs` 生成工作簿）。
- **LibreOffice**（`soffice`），用于 `.xls` → `.xlsx` 转换。
- 上述二进制路径可通过环境变量覆盖（见下文）。

## 🚀 安装

```bash
git clone https://github.com/Gvmeakiss/u8-inventory-valuation.git
cd u8-inventory-valuation
pip install -r requirements.txt
```

## 💡 快速开始 / 使用示例

在项目根目录执行一键流水线（默认运行版本 `u8_caats_202601_202606_v2`，仓库档案默认 `/Users/aatrox/Downloads/仓库档案.XLS`）：

```bash
zsh project_code/run_pipeline.sh
```

指定运行版本与仓库档案路径：

```bash
zsh project_code/run_pipeline.sh 2026H1_v3 /path/to/仓库档案.XLS
```

运行结果写入 `outputs/<run_id>/`。环境路径变化时，设置以下变量覆盖默认 Codex 运行时路径：

```text
U8_PYTHON_BIN      # Python 解释器
U8_NODE_BIN        # Node.js 可执行文件
U8_NODE_MODULES    # 含 @oai/artifact-tool 的 node_modules
U8_SOFFICE_BIN     # LibreOffice soffice 路径
```

## 🧠 核心逻辑（方法论）

- **数据口径分层**：先按 `posted_rule`（`记账人非空`）与 `cost_rule`（`仓库档案.记入成本=是`）界定「已记账成本流水」，再区分普通单据与特殊单据（`其他入库单`、`其他出库单`）。
- **月度独立重算**：以 `prepare_data.py` 的 `read_summaries()` 读取收发存月度汇总，`read_master()` 读取仓库档案，按物料月维度（期初/收入/发出/结存的数量、单价、金额）重算发出计价。
- **连续滚算**：从首月 U8 期初起算，后续月份衔接上月 CAATS 期末；按 `continuous_rollforward_note` 说明，未模拟 U8 自动调整单与最终取价调整。
- **跨仓抵销**：对跨仓调拨进行抵销处理。
- **差异方向**：`difference_direction` 定义为 `U8金额 - CAATS金额`，金额容差 `amount_tolerance=0.01`、数量容差 `quantity_tolerance=0.000001`（浮点尾差用 `within_amount_tolerance()` 以 `round(value, 10)` 判定）。
- **职责分离**：`build_workbooks.mjs` 分 `caats` 与 `ita` 两部分分别生成两份工作簿，审计审阅与底层核对证据解耦。

## 📋 输入与输出

- **输入**：U8 导出（出入库流水账）、收发存汇总（`*.XLS`）、仓库档案（`仓库档案.XLS`）、CAATS 交付模板；由 `config/project.json` 指定 `periods`（如 `202601`–`202606`）。
- **输出**（位于 `outputs/<run_id>/`）：
  - `U8存货发出计价_CAATS审计结果表.xlsx`：重算结果、差异、连续滚算、跨仓抵销。
  - `U8存货发出计价_ITA核对钩稽数据.xlsx`：输入文件清单、月间期初期末衔接、流水桥接、仓库/仓库组拆分、配置档案、CAATS 结果口径、检查结果。

## ⚙️ 配置说明

`config/project.json` 关键字段（以实际文件为准）：

| 字段 | 含义 |
|---|---|
| `periods` | 处理的期间列表（如 `202601`–`202606`） |
| `special_document_types` | 特殊单据类型（`其他入库单`、`其他出库单`） |
| `amount_tolerance` / `quantity_tolerance` | 金额 / 数量容差 |
| `posted_rule` / `cost_rule` | 已记账、记入成本的筛选规则 |
| `difference_direction` | 差异方向定义（U8金额 − CAATS金额） |
| `continuous_rollforward_note` | 连续滚算口径说明 |

## ⚠️ 注意事项

- **数据脱敏**：不含真实客户业务数据；示例文件为脱敏/合成数据（如 NewHope 等化名场景），实际路径与文件名以使用者提供为准。
- **口径说明**：匹配/重算口径以 `config/project.json` 与源码为准；连续滚算未模拟 U8 自动调整单与最终取价调整。

## 🔗 相关仓库

- https://github.com/Gvmeakiss/kpmg-da-skills
- https://github.com/Gvmeakiss/purchase-three-match-configurable
- https://github.com/Gvmeakiss/purchase-three-match-toolkit
- https://github.com/Gvmeakiss/purchase-three-match-newhope

## 📄 License

MIT

---

<div align="center">

*Disclaimer: Personal project and personal views. Not affiliated with or endorsed by KPMG or any client.*<br>
*本仓库为个人项目与个人观点，与任何前/现雇主及客户无关。*

</div>
