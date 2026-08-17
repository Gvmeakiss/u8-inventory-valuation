# U8存货发出计价项目代码

更新时间：2026-08-17

本目录保留两套相互隔离的程序。当前应使用财务核算流水账记账日期版；根目录程序仅用于历史口径追溯。

## 当前执行版本

当前入口：

```text
code/financial_ledger/run_pipeline.sh
```

当前流水账：

```text
input/current/财务核算-流水账 记账日期-2026.01-.06.xlsx
```

核心取数规则：

```text
月份 = 记账日期
记账人非空
仓库档案.记入成本 = 是
排除业务类型：调拨入库、调拨出库
```

本版不使用借入、借出移动类别。详细字段、公式、工作簿结构和质量控制参见 [financial_ledger/README.md](financial_ledger/README.md)。

## 运行方式

```zsh
cd /Users/aatrox/Desktop/U8存货发出计价/code/financial_ledger
./run_pipeline.sh <new_run_id>
```

每次必须使用新的`run_id`，避免覆盖已复核结果。输出分为：

```text
output/runs/<run_id>/   # 三份正式工作簿
output/work/<run_id>/   # 转换文件、标准化数据、检查记录和视觉QA
```

当前复核通过的最新运行：

```text
output/runs/u8_caats_202601_202606_financial_ledger_postingdate_summary_balances_v5/
```

## 当前代码结构

```text
code/
  financial_ledger/             # 当前财务核算流水账记账日期版
    config/project.json         # 期间、剔除类型、容差和差异方向
    src/prepare_data.py         # 字段校验、XLS转换、标准化和计算
    src/build_workbooks.mjs     # 生成CAATS、ITA和剔除明细工作簿
    src/verify_outputs.mjs      # 结构、公式、钩稽、哈希和文件完整性检查
    tests/test_calculations.py  # 核心规则单元测试
    run_pipeline.sh             # 当前一键执行入口
    README.md                   # 当前技术说明
  config/ src/ tests/           # 历史仓库管理流水账版，仅供追溯
  run_pipeline.sh               # 历史入口，不作为当前交付重跑入口
```

## 输出边界

1. `U8存货发出计价_CAATS审计结果表.xlsx`：供财务审计审阅结果。
2. `U8存货发出计价_调拨出入库剔除明细.xlsx`：供财务审计追溯剔除流水和收入金额桥接。
3. `U8存货发出计价_ITA核对钩稽数据.xlsx`：内部数据完整性、配置和钩稽证据，不作为财务审计主结果表。

原始业务文件、输出工作簿、转换缓存和客户配置证据不得提交至公开Git仓库。
