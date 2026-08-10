#!/usr/bin/env python3
"""读取U8原始导出，形成可供artifact-tool生成工作簿的标准JSON。"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import subprocess
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook


SUMMARY_FIELDS = [
    "期初数量", "期初单价", "期初金额",
    "收入数量", "收入单价", "收入金额",
    "发出数量", "发出单价", "发出金额",
    "结存数量", "结存单价", "结存金额",
]


def number(value: Any) -> float:
    try:
        if value in (None, ""):
            return 0.0
        result = float(value)
        return 0.0 if not math.isfinite(result) else result
    except (TypeError, ValueError):
        return 0.0


def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def within_amount_tolerance(value: float, tolerance: float) -> bool:
    """按可审计精度消除二进制浮点尾差后判断金额容差。"""
    return abs(round(value, 10)) <= tolerance


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def required_columns(headers: Iterable[Any], required: Iterable[str], label: str) -> dict[str, int]:
    mapping = {text(value): index for index, value in enumerate(headers) if text(value)}
    missing = [name for name in required if name not in mapping]
    if missing:
        raise ValueError(f"{label}缺少必需字段：{', '.join(missing)}")
    return mapping


def month_from_filename(path: Path) -> str:
    match = re.search(r"(20\d{2})[._-]?(0[1-9]|1[0-2])", path.stem)
    if not match:
        raise ValueError(f"无法从文件名识别期间：{path.name}")
    return f"{match.group(1)}{match.group(2)}"


def convert_xls(source: Path, output_dir: Path, soffice: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    if source.suffix.lower() == ".xlsx":
        return source
    if source.suffix.lower() != ".xls":
        raise ValueError(f"仅支持XLS/XLSX：{source}")
    target = output_dir / f"{source.stem}.xlsx"
    if target.exists():
        target.unlink()
    completed = subprocess.run(
        [str(soffice), "--headless", "--convert-to", "xlsx", "--outdir", str(output_dir), str(source)],
        check=False,
        capture_output=True,
        text=True,
    )
    if completed.returncode != 0 or not target.exists():
        raise RuntimeError(
            f"XLS转换失败：{source}\nstdout={completed.stdout}\nstderr={completed.stderr}"
        )
    return target


def read_master(path: Path) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook.active
    rows = worksheet.iter_rows(values_only=True)
    headers = list(next(rows))
    index = required_columns(
        headers,
        ["仓库编码", "仓库名称", "计价方式", "仓库核算组", "记入成本"],
        "仓库档案",
    )
    master: list[dict[str, Any]] = []
    by_name: dict[str, dict[str, Any]] = {}
    for row in rows:
        if row[index["仓库编码"]] in (None, ""):
            continue
        record = {name: row[column] for name, column in index.items()}
        record["仓库编码"] = text(record["仓库编码"])
        record["仓库名称"] = text(record["仓库名称"])
        record["仓库核算组"] = text(record.get("仓库核算组"))
        if record.get("停用日期") is not None:
            record["停用日期"] = str(record["停用日期"])[:10]
        master.append(record)
        if record["仓库名称"] in by_name:
            raise ValueError(f"仓库档案仓库名称重复：{record['仓库名称']}")
        by_name[record["仓库名称"]] = record
    return master, by_name


def read_summaries(paths: list[Path], periods: list[str]) -> tuple[list[dict[str, Any]], dict[tuple[str, str], dict[str, Any]]]:
    summaries: list[dict[str, Any]] = []
    by_key: dict[tuple[str, str], dict[str, Any]] = {}
    found_periods: list[str] = []
    for path in sorted(paths):
        month = month_from_filename(path)
        if month not in periods:
            continue
        found_periods.append(month)
        workbook = load_workbook(path, read_only=True, data_only=True)
        worksheet = workbook.active
        rows = worksheet.iter_rows(values_only=True)
        headers = list(next(rows))
        index = required_columns(
            headers,
            ["存货编码", "存货名称", "存货规格", "存货单位", *SUMMARY_FIELDS],
            path.name,
        )
        for row_number, row in enumerate(rows, start=2):
            code = text(row[index["存货编码"]])
            if not code or code == "合计":
                continue
            record: dict[str, Any] = {
                "month": month,
                "code": code,
                "stock_code": text(row[index["存货代码"]]) if "存货代码" in index else "",
                "name": text(row[index["存货名称"]]),
                "spec": text(row[index["存货规格"]]),
                "unit": text(row[index["存货单位"]]),
                "weight": row[index["存货重量"]] if "存货重量" in index else None,
                "reg": text(row[index["存货产品注册证号"]]) if "存货产品注册证号" in index else "",
                "manufacturer": text(row[index["存货生产厂商"]]) if "存货生产厂商" in index else "",
                "source_file": path.name,
                "source_row": row_number,
            }
            for field in SUMMARY_FIELDS:
                record[field] = number(row[index[field]])
            key = (month, code)
            if key in by_key:
                raise ValueError(f"收发存存在重复物料月份：{month}/{code}")
            summaries.append(record)
            by_key[key] = record
    if sorted(found_periods) != sorted(periods):
        raise ValueError(f"收发存期间不完整。期望={periods}，实际={sorted(found_periods)}")
    return summaries, by_key


def build_transition_checks(
    material_rows: list[dict[str, Any]], periods: list[str], quantity_tolerance: float, amount_tolerance: float
) -> list[dict[str, Any]]:
    row_by_key = {(row["month"], row["code"]): row for row in material_rows}
    codes = sorted({row["code"] for row in material_rows})
    checks: list[dict[str, Any]] = []
    for position in range(1, len(periods)):
        previous_month = periods[position - 1]
        month = periods[position]
        quantity_mismatch = 0
        amount_mismatch = 0
        quantity_abs = 0.0
        amount_abs = 0.0
        compared = 0
        for code in codes:
            previous = row_by_key.get((previous_month, code))
            current = row_by_key.get((month, code))
            if previous is None and current is None:
                continue
            compared += 1
            quantity_difference = number(current and current["期初数量"]) - number(previous and previous["结存数量"])
            amount_difference = number(current and current["期初金额"]) - number(previous and previous["结存金额"])
            quantity_abs += abs(quantity_difference)
            amount_abs += abs(amount_difference)
            quantity_mismatch += abs(quantity_difference) > quantity_tolerance
            amount_mismatch += abs(amount_difference) > amount_tolerance
        checks.append(
            {
                "previous_month": previous_month,
                "month": month,
                "compared": compared,
                "quantity_mismatch": quantity_mismatch,
                "amount_mismatch": amount_mismatch,
                "quantity_abs": quantity_abs,
                "amount_abs": amount_abs,
                "status": "PASS" if quantity_mismatch == 0 and amount_mismatch == 0 else "REVIEW",
            }
        )
    return checks


def build_continuous_summary(material_rows: list[dict[str, Any]], periods: list[str]) -> list[dict[str, Any]]:
    row_by_key = {(row["month"], row["code"]): row for row in material_rows}
    codes = sorted({row["code"] for row in material_rows})
    rolled: dict[str, dict[str, Any]] = {}
    result: list[dict[str, Any]] = []
    for month_position, month in enumerate(periods):
        u8_issue = 0.0
        independent_caats_issue = 0.0
        independent_difference = 0.0
        independent_abs = 0.0
        continuous_caats_issue = 0.0
        continuous_difference = 0.0
        continuous_abs = 0.0
        continuous_end_difference = 0.0
        continuous_end_abs = 0.0
        rows = 0
        resets = 0
        for code in codes:
            record = row_by_key.get((month, code))
            if record is None:
                continue
            rows += 1
            prior = rolled.get(code)
            can_roll = prior is not None
            begin_quantity = prior["end_quantity"] if can_roll else record["期初数量"]
            begin_amount = prior["end_amount"] if can_roll else record["期初金额"]
            if month_position and not can_roll:
                resets += 1
            denominator = begin_quantity + record["收入数量"]
            average = (begin_amount + record["收入金额"]) / denominator if abs(denominator) > 1e-12 else 0.0
            calculated_issue = record["发出数量"] * average
            end_quantity = begin_quantity + record["收入数量"] - record["发出数量"]
            end_amount = begin_amount + record["收入金额"] - calculated_issue
            issue_difference = record["发出金额"] - calculated_issue
            end_difference = record["结存金额"] - end_amount
            u8_issue += record["发出金额"]
            independent_caats_issue += record["caats_issue"]
            independent_difference += record["issue_diff"]
            independent_abs += abs(record["issue_diff"])
            continuous_caats_issue += calculated_issue
            continuous_difference += issue_difference
            continuous_abs += abs(issue_difference)
            continuous_end_difference += end_difference
            continuous_end_abs += abs(end_difference)
            rolled[code] = {"month": month, "end_quantity": end_quantity, "end_amount": end_amount}
        result.append(
            {
                "month": month,
                "rows": rows,
                "u8_issue": u8_issue,
                "independent_caats_issue": independent_caats_issue,
                "independent_difference": independent_difference,
                "independent_abs": independent_abs,
                "continuous_caats_issue": continuous_caats_issue,
                "continuous_difference": continuous_difference,
                "continuous_abs": continuous_abs,
                "continuous_end_difference": continuous_end_difference,
                "continuous_end_abs": continuous_end_abs,
                "carry_impact": continuous_difference - independent_difference,
                "resets": resets,
            }
        )
    return result


def build_offset_rows(
    material_rows: list[dict[str, Any]], warehouse_rows: list[dict[str, Any]], quantity_tolerance: float, amount_tolerance: float
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    material_by_key = {(row["month"], row["code"]): row for row in material_rows}
    warehouses_by_key: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in warehouse_rows:
        warehouses_by_key[(row["month"], row["code"])].append(row)
    result: list[dict[str, Any]] = []
    multi_issue_keys = 0
    clean_multi_keys = 0
    both_sign_keys = 0
    total_offset = 0.0
    for key, rows in warehouses_by_key.items():
        issue_rows = [row for row in rows if abs(row["oq"]) > quantity_tolerance]
        if len(issue_rows) < 2:
            continue
        multi_issue_keys += 1
        material = material_by_key.get(key)
        if material is None:
            continue
        warehouse_quantity = sum(row["oq"] for row in rows)
        if abs(warehouse_quantity - material["发出数量"]) > quantity_tolerance:
            continue
        if abs(material["issue_amount_gap"]) > amount_tolerance:
            continue
        clean_multi_keys += 1
        parts: list[dict[str, Any]] = []
        for row in rows:
            if abs(row["oq"]) <= quantity_tolerance and abs(row["oa"]) <= amount_tolerance:
                continue
            contribution = row["oa"] - row["oq"] * material["caats_avg"]
            parts.append({"warehouse_row": row, "contribution": contribution})
        positive = sum(max(part["contribution"], 0.0) for part in parts)
        negative_abs = sum(max(-part["contribution"], 0.0) for part in parts)
        if positive <= amount_tolerance or negative_abs <= amount_tolerance:
            continue
        both_sign_keys += 1
        offset = min(positive, negative_abs)
        total_offset += offset
        for part in parts:
            row = part["warehouse_row"]
            result.append(
                {
                    "month": material["month"],
                    "code": material["code"],
                    "name": material["name"],
                    "spec": material["spec"],
                    "unit": material["unit"],
                    "pooled_avg": material["caats_avg"],
                    "material_issue_difference": material["issue_diff"],
                    "warehouse": row["warehouse"],
                    "warehouse_code": row["warehouse_code"],
                    "group": row["group"],
                    "warehouse_issue_quantity": row["oq"],
                    "warehouse_issue_amount": row["oa"],
                    "pooled_expected_amount": row["oq"] * material["caats_avg"],
                    "warehouse_contribution": part["contribution"],
                    "key_positive": positive,
                    "key_negative_abs": negative_abs,
                    "key_offset": offset,
                    "key_net": positive - negative_abs,
                }
            )
    result.sort(key=lambda row: (-row["key_offset"], row["month"], row["code"], -abs(row["warehouse_contribution"])))
    summary = {
        "multi_issue_keys": multi_issue_keys,
        "clean_multi_keys": clean_multi_keys,
        "both_sign_keys": both_sign_keys,
        "offset_rows": len(result),
        "total_offset": total_offset,
    }
    return result, summary


def prepare(args: argparse.Namespace) -> dict[str, Any]:
    project_root = args.project_root.resolve()
    staging_dir = args.staging_dir.resolve()
    staging_dir.mkdir(parents=True, exist_ok=True)
    converted_dir = staging_dir / "converted_xls"
    config = json.loads(args.config.read_text(encoding="utf-8"))
    periods = [str(value) for value in config["periods"]]
    amount_tolerance = float(config["amount_tolerance"])
    quantity_tolerance = float(config["quantity_tolerance"])
    special_document_types = set(config["special_document_types"])

    ledger_path = (args.ledger or project_root / "仓库管理-出入库流水账-2026.01-06.xlsx").resolve()
    summary_dir = (args.summary_dir or project_root / "收发存汇总").resolve()
    warehouse_master_source = args.warehouse_master.resolve()
    warehouse_master_path = convert_xls(warehouse_master_source, converted_dir / "warehouse_master", args.soffice)
    summary_sources = sorted(path for path in summary_dir.iterdir() if path.suffix.lower() in {".xls", ".xlsx"})
    converted_summaries = [convert_xls(path, converted_dir / "summaries", args.soffice) for path in summary_sources]

    master, master_by_name = read_master(warehouse_master_path)
    summaries, summary_by_key = read_summaries(converted_summaries, periods)

    workbook = load_workbook(ledger_path, read_only=True, data_only=True)
    worksheet = workbook.active
    rows = worksheet.iter_rows(min_row=2, values_only=True)
    headers = list(next(rows))
    ledger_index = required_columns(
        headers,
        ["日期", "单据类型", "单据号", "仓库", "存货编码", "收发类别", "记账人", "入库数量", "入库金额", "出库数量", "出库金额"],
        ledger_path.name,
    )

    material_aggregate: dict[tuple[str, str], defaultdict[str, float]] = defaultdict(lambda: defaultdict(float))
    warehouse_aggregate: dict[tuple[str, str, str], defaultdict[str, float]] = defaultdict(lambda: defaultdict(float))
    warehouse_categories: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    warehouse_document_types: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    group_aggregate: dict[tuple[str, str, str], defaultdict[str, float]] = defaultdict(lambda: defaultdict(float))
    group_categories: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    valid_rows = posted_cost_rows = unposted_rows = no_cost_rows = 0
    missing_master: set[str] = set()

    for row in rows:
        if row[ledger_index["存货编码"]] in (None, "") or row[ledger_index["日期"]] in (None, ""):
            continue
        valid_rows += 1
        month = text(row[ledger_index["日期"]])[:7].replace("-", "")
        if month not in periods:
            continue
        code = text(row[ledger_index["存货编码"]])
        warehouse = text(row[ledger_index["仓库"]])
        category = text(row[ledger_index["收发类别"]])
        document_type = text(row[ledger_index["单据类型"]])
        warehouse_master = master_by_name.get(warehouse)
        if warehouse_master is None:
            missing_master.add(warehouse)
            continue
        posted = row[ledger_index["记账人"]] not in (None, "")
        costed = warehouse_master.get("记入成本") == "是"
        inbound_quantity = number(row[ledger_index["入库数量"]])
        inbound_amount = number(row[ledger_index["入库金额"]])
        outbound_quantity = number(row[ledger_index["出库数量"]])
        outbound_amount = number(row[ledger_index["出库金额"]])
        warehouse_key = (month, warehouse, code)
        material_key = (month, code)
        warehouse_categories[warehouse_key].add(category)
        warehouse_document_types[warehouse_key].add(document_type)
        warehouse_aggregate[warehouse_key]["all_rows"] += 1
        if not posted:
            unposted_rows += 1
            warehouse_aggregate[warehouse_key]["unposted_rows"] += 1
            warehouse_aggregate[warehouse_key]["unposted_iq"] += inbound_quantity
            warehouse_aggregate[warehouse_key]["unposted_oq"] += outbound_quantity
            continue
        if not costed:
            no_cost_rows += 1
            warehouse_aggregate[warehouse_key]["posted_no_cost_rows"] += 1
            warehouse_aggregate[warehouse_key]["posted_no_cost_iq"] += inbound_quantity
            warehouse_aggregate[warehouse_key]["posted_no_cost_oq"] += outbound_quantity
            continue
        posted_cost_rows += 1
        for aggregate, key in ((material_aggregate, material_key), (warehouse_aggregate, warehouse_key)):
            aggregate[key]["posted_rows"] += 1
            aggregate[key]["iq"] += inbound_quantity
            aggregate[key]["ia"] += inbound_amount
            aggregate[key]["oq"] += outbound_quantity
            aggregate[key]["oa"] += outbound_amount
        is_special = document_type in special_document_types
        if is_special:
            for name, value in (("special_rows", 1), ("special_iq", inbound_quantity), ("special_ia", inbound_amount), ("special_oq", outbound_quantity), ("special_oa", outbound_amount)):
                warehouse_aggregate[warehouse_key][name] += value
            group = text(warehouse_master.get("仓库核算组")) or "(无组)"
            group_key = (month, group, code)
            group_aggregate[group_key]["rows"] += 1
            group_aggregate[group_key]["iq"] += inbound_quantity
            group_aggregate[group_key]["ia"] += inbound_amount
            group_aggregate[group_key]["oq"] += outbound_quantity
            group_aggregate[group_key]["oa"] += outbound_amount
            group_categories[group_key].add(category)
        else:
            for name, value in (("normal_rows", 1), ("normal_iq", inbound_quantity), ("normal_ia", inbound_amount), ("normal_oq", outbound_quantity), ("normal_oa", outbound_amount)):
                warehouse_aggregate[warehouse_key][name] += value
        if category == "调拨入库":
            material_aggregate[material_key]["transfer_iq"] += inbound_quantity
            material_aggregate[material_key]["transfer_ia"] += inbound_amount
            warehouse_aggregate[warehouse_key]["transfer_iq"] += inbound_quantity
            warehouse_aggregate[warehouse_key]["transfer_ia"] += inbound_amount
        if category == "调拨出库":
            material_aggregate[material_key]["transfer_oq"] += outbound_quantity
            material_aggregate[material_key]["transfer_oa"] += outbound_amount
            warehouse_aggregate[warehouse_key]["transfer_oq"] += outbound_quantity
            warehouse_aggregate[warehouse_key]["transfer_oa"] += outbound_amount

    price_by_material: dict[tuple[str, str], list[float]] = defaultdict(list)
    warehouses_by_material: dict[tuple[str, str], set[str]] = defaultdict(set)
    issue_warehouses_by_material: dict[tuple[str, str], set[str]] = defaultdict(set)
    for (month, warehouse, code), aggregate in warehouse_aggregate.items():
        if aggregate.get("posted_rows", 0):
            warehouses_by_material[(month, code)].add(warehouse)
        if aggregate.get("oq", 0):
            issue_warehouses_by_material[(month, code)].add(warehouse)
            price_by_material[(month, code)].append(aggregate.get("oa", 0) / aggregate["oq"])

    material_rows: list[dict[str, Any]] = []
    for record in sorted(summaries, key=lambda item: (item["month"], item["code"])):
        key = (record["month"], record["code"])
        aggregate = material_aggregate.get(key, {})
        denominator = record["期初数量"] + record["收入数量"]
        average = (record["期初金额"] + record["收入金额"]) / denominator if abs(denominator) > 1e-12 else 0.0
        calculated_issue = record["发出数量"] * average
        calculated_end = record["期初金额"] + record["收入金额"] - calculated_issue
        issue_difference = record["发出金额"] - calculated_issue
        prices = price_by_material.get(key, [])
        minimum_price = min(prices) if prices else None
        maximum_price = max(prices) if prices else None
        income_gap = record["收入金额"] - aggregate.get("ia", 0)
        issue_gap = record["发出金额"] - aggregate.get("oa", 0)
        multi_price = len(issue_warehouses_by_material.get(key, set())) > 1 and minimum_price is not None and maximum_price is not None and abs(maximum_price - minimum_price) > amount_tolerance
        amount_gap = abs(income_gap) > amount_tolerance or abs(issue_gap) > amount_tolerance
        if multi_price and amount_gap:
            reason = "跨仓价格混合+流水金额调整缺口"
        elif multi_price:
            reason = "跨仓价格混合"
        elif amount_gap:
            reason = "流水未含调整/最终取价差异"
        elif within_amount_tolerance(issue_difference, amount_tolerance):
            reason = "物料维度可解释"
        else:
            reason = "单仓特殊单据/舍入或其他配置"
        item = dict(record)
        item.update(
            {
                "ledger_iq": aggregate.get("iq", 0),
                "ledger_ia": aggregate.get("ia", 0),
                "ledger_oq": aggregate.get("oq", 0),
                "ledger_oa": aggregate.get("oa", 0),
                "transfer_iq": aggregate.get("transfer_iq", 0),
                "transfer_ia": aggregate.get("transfer_ia", 0),
                "transfer_oq": aggregate.get("transfer_oq", 0),
                "transfer_oa": aggregate.get("transfer_oa", 0),
                "warehouse_count": len(warehouses_by_material.get(key, set())),
                "issue_warehouse_count": len(issue_warehouses_by_material.get(key, set())),
                "warehouse_price_min": minimum_price,
                "warehouse_price_max": maximum_price,
                "warehouse_price_range": maximum_price - minimum_price if prices else None,
                "caats_avg": average,
                "caats_issue": calculated_issue,
                "issue_diff": issue_difference,
                "caats_end": calculated_end,
                "end_diff": record["结存金额"] - calculated_end,
                "income_amount_gap": income_gap,
                "issue_amount_gap": issue_gap,
                "reason": reason,
                "status": "PASS" if within_amount_tolerance(issue_difference, amount_tolerance) else "REVIEW",
            }
        )
        material_rows.append(item)

    warehouse_rows: list[dict[str, Any]] = []
    for (month, warehouse, code), aggregate in sorted(warehouse_aggregate.items()):
        if not aggregate.get("all_rows", 0):
            continue
        warehouse_master = master_by_name[warehouse]
        summary = summary_by_key.get((month, code), {})
        warehouse_rows.append(
            {
                "month": month,
                "warehouse_code": warehouse_master["仓库编码"],
                "warehouse": warehouse,
                "group": text(warehouse_master.get("仓库核算组")),
                "method": text(warehouse_master.get("计价方式")),
                "costed": text(warehouse_master.get("记入成本")),
                "stopped": text(warehouse_master.get("停用日期")),
                "code": code,
                "name": summary.get("name", ""),
                "spec": summary.get("spec", ""),
                "unit": summary.get("unit", ""),
                "posted_rows": aggregate.get("posted_rows", 0),
                "unposted_rows": aggregate.get("unposted_rows", 0),
                "posted_no_cost_rows": aggregate.get("posted_no_cost_rows", 0),
                "iq": aggregate.get("iq", 0),
                "ia": aggregate.get("ia", 0),
                "oq": aggregate.get("oq", 0),
                "oa": aggregate.get("oa", 0),
                "out_unit": aggregate.get("oa", 0) / aggregate["oq"] if aggregate.get("oq", 0) else None,
                "normal_iq": aggregate.get("normal_iq", 0),
                "normal_ia": aggregate.get("normal_ia", 0),
                "normal_oq": aggregate.get("normal_oq", 0),
                "normal_oa": aggregate.get("normal_oa", 0),
                "special_iq": aggregate.get("special_iq", 0),
                "special_ia": aggregate.get("special_ia", 0),
                "special_oq": aggregate.get("special_oq", 0),
                "special_oa": aggregate.get("special_oa", 0),
                "transfer_iq": aggregate.get("transfer_iq", 0),
                "transfer_ia": aggregate.get("transfer_ia", 0),
                "transfer_oq": aggregate.get("transfer_oq", 0),
                "transfer_oa": aggregate.get("transfer_oa", 0),
                "unposted_iq": aggregate.get("unposted_iq", 0),
                "unposted_oq": aggregate.get("unposted_oq", 0),
                "categories": "、".join(sorted(value for value in warehouse_categories[(month, warehouse, code)] if value)),
                "doc_types": "、".join(sorted(value for value in warehouse_document_types[(month, warehouse, code)] if value)),
                "u8_mat_iq": summary.get("收入数量", 0),
                "u8_mat_oq": summary.get("发出数量", 0),
                "mat_sum_iq": material_aggregate.get((month, code), {}).get("iq", 0),
                "mat_sum_oq": material_aggregate.get((month, code), {}).get("oq", 0),
            }
        )

    group_rows: list[dict[str, Any]] = []
    for (month, group, code), aggregate in sorted(group_aggregate.items()):
        summary = summary_by_key.get((month, code), {})
        group_rows.append(
            {
                "month": month,
                "group": group,
                "code": code,
                "name": summary.get("name", ""),
                "spec": summary.get("spec", ""),
                "unit": summary.get("unit", ""),
                "rows": aggregate.get("rows", 0),
                "iq": aggregate.get("iq", 0),
                "ia": aggregate.get("ia", 0),
                "oq": aggregate.get("oq", 0),
                "oa": aggregate.get("oa", 0),
                "in_unit": aggregate.get("ia", 0) / aggregate["iq"] if aggregate.get("iq", 0) else None,
                "out_unit": aggregate.get("oa", 0) / aggregate["oq"] if aggregate.get("oq", 0) else None,
                "categories": "、".join(sorted(value for value in group_categories[(month, group, code)] if value)),
            }
        )

    monthly: list[dict[str, Any]] = []
    for month in periods:
        rows_for_month = [row for row in material_rows if row["month"] == month]
        monthly.append(
            {
                "month": month,
                "rows": len(rows_for_month),
                "review": sum(row["status"] == "REVIEW" for row in rows_for_month),
                "pass": sum(row["status"] == "PASS" for row in rows_for_month),
                "signed_issue_diff": sum(row["issue_diff"] for row in rows_for_month),
                "abs_issue_diff": sum(abs(row["issue_diff"]) for row in rows_for_month),
            }
        )

    transition_checks = build_transition_checks(material_rows, periods, quantity_tolerance, amount_tolerance)
    continuous_monthly = build_continuous_summary(material_rows, periods)
    offset_rows, offset_summary = build_offset_rows(material_rows, warehouse_rows, quantity_tolerance, amount_tolerance)
    positive_difference = sum(max(row["issue_diff"], 0.0) for row in material_rows)
    negative_difference_abs = sum(max(-row["issue_diff"], 0.0) for row in material_rows)
    metrics = {
        "material_month_rows": len(material_rows),
        "warehouse_material_month_rows": len(warehouse_rows),
        "group_special_rows": len(group_rows),
        "valid_ledger_rows": valid_rows,
        "posted_cost_rows": posted_cost_rows,
        "unposted_rows": unposted_rows,
        "no_cost_rows": no_cost_rows,
        "missing_master": sorted(missing_master),
        "review_rows": sum(row["status"] == "REVIEW" for row in material_rows),
        "pass_rows": sum(row["status"] == "PASS" for row in material_rows),
        "signed_issue_diff": sum(row["issue_diff"] for row in material_rows),
        "abs_issue_diff": sum(abs(row["issue_diff"]) for row in material_rows),
        "max_abs_issue_diff": max(abs(row["issue_diff"]) for row in material_rows),
        "amount_gap_rows": sum(abs(row["income_amount_gap"]) > amount_tolerance or abs(row["issue_amount_gap"]) > amount_tolerance for row in material_rows),
        "positive_issue_diff": positive_difference,
        "negative_issue_diff_abs": negative_difference_abs,
        "global_netting": min(positive_difference, negative_difference_abs),
    }
    input_files = [ledger_path, warehouse_master_source, *summary_sources, project_root / "CAATS交付模板.xlsx"]
    source_manifest = [
        {"role": "输入文件", "path": str(path), "name": path.name, "size": path.stat().st_size, "sha256": sha256(path)}
        for path in input_files
    ]
    return {
        "config": config,
        "metrics": metrics,
        "monthly": monthly,
        "continuous_monthly": continuous_monthly,
        "transition_checks": transition_checks,
        "offset_rows": offset_rows,
        "offset_summary": offset_summary,
        "material_rows": material_rows,
        "difference_rows": sorted(
            (row for row in material_rows if row["status"] == "REVIEW"),
            key=lambda row: abs(row["issue_diff"]),
            reverse=True,
        ),
        "warehouse_rows": warehouse_rows,
        "group_rows": group_rows,
        "master": master,
        "source_manifest": source_manifest,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-root", type=Path, required=True)
    parser.add_argument("--warehouse-master", type=Path, required=True)
    parser.add_argument("--config", type=Path, required=True)
    parser.add_argument("--soffice", type=Path, required=True)
    parser.add_argument("--staging-dir", type=Path, required=True)
    parser.add_argument("--output-json", type=Path, required=True)
    parser.add_argument("--ledger", type=Path)
    parser.add_argument("--summary-dir", type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    payload = prepare(args)
    args.output_json.parent.mkdir(parents=True, exist_ok=True)
    args.output_json.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":"), default=str), encoding="utf-8")
    print(json.dumps({"metrics": payload["metrics"], "offset_summary": payload["offset_summary"], "output_json": str(args.output_json)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
