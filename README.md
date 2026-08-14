# U8 存货发出计价工具

> 将 U8 导出、仓库档案与 CAATS 模板处理为「CAATS 审计结果表」与「ITA 核对钩稽数据」两份职责分离的工作簿。

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8%2B-3776AB?logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/Node.js-18%2B-339933?logo=node.js&logoColor=white" alt="Node">
  <img src="https://img.shields.io/badge/Output-Excel%20(xlsx)-217346" alt="Excel">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License">
  <img src="https://img.shields.io/github/last-commit/Gvmeakiss/u8-inventory-valuation?label=updated" alt="Updated">
</p>

## 📋 目录

- [一、项目结构](#一项目结构)
- [二、运行方式](#二运行方式)
- [三、数据口径](#三数据口径)
  - [已记账成本流水](#已记账成本流水)
  - [普通与特殊单据](#普通与特殊单据)
  - [月度独立重算](#月度独立重算)
  - [连续滚算](#连续滚算)
  - [跨仓抵销](#跨仓抵销)
- [四、CAATS与ITA工作簿边界](#四caats与ita工作簿边界)
- [五、关键控制](#五关键控制)
- [六、新数据导入](#六新数据导入)

本目录将原始U8导出、仓库档案和CAATS模板处理为两份职责分离的正式工作簿：

1. `U8存货发出计价_CAATS审计结果表.xlsx`
   - 仅包含财务审计同事需要审阅的重计算结果、差异、连续滚算和跨仓抵销结果。
   - 不包含流水桥接、仓库档案、输入文件哈希等ITA底层核对证据。
2. `U8存货发出计价_ITA核对钩稽数据.xlsx`
   - 保存输入文件清单、月间期初期末衔接、流水桥接、仓库/仓库组拆分、配置档案、CAATS结果口径和检查结果。

## 一、项目结构

```text
project_code/
  config/project.json        # 期间、特殊单据、容差和差异方向
  src/prepare_data.py        # XLS转换、字段校验、数据标准化和计算
  src/build_workbooks.mjs    # 使用artifact-tool生成两份Excel
  src/verify_outputs.mjs     # 结果、公式错误、结构和视觉检查
  tests/test_calculations.py # 核心公式单元测试
  run_pipeline.sh            # 一键执行入口
```

原始业务文件保留在项目根目录，不由程序改写：

```text
CAATS交付模板.xlsx
仓库管理-出入库流水账-2026.01-06.xlsx
收发存汇总/*.XLS
```

仓库档案默认读取：

```text
/Users/aatrox/Downloads/仓库档案.XLS
```

也可以在运行时传入其他路径。

## 二、运行方式

在项目根目录执行：

```bash
zsh project_code/run_pipeline.sh
```

指定运行版本和仓库档案：

```bash
zsh project_code/run_pipeline.sh 2026H1_v3 /path/to/仓库档案.XLS
```

运行结果进入：

```text
outputs/<run_id>/
```

运行脚本默认使用Codex桌面版提供的Node.js、Python、LibreOffice和`@oai/artifact-tool`。环境路径变化时，可设置：

```text
U8_PYTHON_BIN
U8_NODE_BIN
U8_NODE_MODULES
U8_SOFFICE_BIN
```

## 三、数据口径

### 已记账成本流水

```text
记账人非空 AND 仓库档案.记入成本=是
```

借用归还、借出借用、调拨入库、调拨出库不按名称整类剔除。未记账流水保留在ITA表，但不进入已记账成本重算。

### 普通与特殊单据

```text
普通单据：仓库+物料+月份
特殊单据：仓库核算组+物料+月份
```

特殊单据类型从`config/project.json`读取，当前为`其他入库单`、`其他出库单`。

### 月度独立重算

```text
月平均价 = (U8期初金额 + U8收入金额) / (U8期初数量 + U8收入数量)
CAATS发出金额 = U8发出数量 × 月平均价
差异 = U8发出金额 - CAATS发出金额
```

每个月使用客户月表的U8期初，因此该口径不会把上月CAATS差异带入下月。

### 连续滚算

每个物料首次出现时使用当月U8期初；后续月份使用该物料最近一次可追溯的连续CAATS期末数量和金额作为本月期初。当月收入和发出数量仍取客户月表。若中间缺月，追溯链不重置。

连续滚算属于敏感性分析，当前未模拟自动出库调整单、成本卷积调整和最终取价调整，因此不能直接作为账务调整金额。

### 跨仓抵销

对数量、金额均能与U8物料月汇总干净桥接、且同月同物料至少两个仓库发生出库的组合，计算：

```text
仓库贡献 = 仓库U8发出金额 - 仓库发出数量 × 物料统一CAATS平均价
```

同一物料月份同时存在正负仓库贡献时，记录为跨仓抵销。该贡献是相对物料统一价的分析，不代表仓库错账。

## 四、CAATS与ITA工作簿边界

CAATS结果表：

- `01_物料月维度明细`
- `02_物料月维度结果汇总`
- `03_连续滚算明细`
- `04_连续滚算汇总`
- `05_存在差异明细`
- `06_跨仓抵销`

ITA核对表：

- `01_核对摘要`
- `02_输入清单`
- `03_月间衔接`
- `04_流水桥接`
- `05_仓库物料月`
- `06_物料月桥接`
- `07_仓库组特殊单据`
- `08_仓库档案`
- `09_口径说明`
- `10_检查`

CAATS结果表只保留财务审计需要的六张结果Sheet；计算口径、限制及底层钩稽证据统一归入ITA的`09_口径说明`及相关明细Sheet。`05_存在差异明细`释放全部REVIEW记录，不再限定Top200，也不重复展示“绝对差异”字段。

两份工作簿统一使用“深蓝标题、蓝色表头、浅蓝网格、绿色PASS、黄色REVIEW”的视觉层级；有效数据区域全部施加浅色单元格边框，主结果表仅对收入、发出、结存和CAATS结果字段保留分组色。

## 五、关键控制

每次运行至少验证：

1. 六个月收发存文件期间完整且无重复物料月份。
2. 流水仓库全部映射仓库档案。
3. 上月期末数量/金额等于下月期初数量/金额。
4. 流水与收发存的收入、发出数量桥接。
5. CAATS差异汇总等于物料月明细。
6. 跨仓抵销仅使用数量和金额桥接通过的组合。
7. 两份工作簿均无`#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?`、`#N/A`等公式错误。

## 六、新数据导入

新增月份时：

1. 将原始收发存XLS放入`收发存汇总/`，不得覆盖旧文件。
2. 更新`config/project.json`中的`periods`。
3. 流水账路径如变化，在`prepare_data.py`运行参数中指定或更新运行脚本。
4. 保留存货编码、仓库编码为文本，避免前导零丢失。
5. 使用新的`run_id`运行，不能覆盖已交付版本。

程序在中间目录记录原始文件SHA-256、字段、行数和期间；中间数据可重建，不属于正式交付。

---

<div align="center">

**James Li · 审计数据分析工具集**

📫 本工具用于内部审计与数据核对，辅助分析但不替代专业判断，不作为对外签字版本。

</div>
