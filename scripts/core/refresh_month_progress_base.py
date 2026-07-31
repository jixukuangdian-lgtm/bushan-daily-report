#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import sync_feishu_base_workbook_may as base_sync


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = SCRIPT_DIR / "daily_report_pipeline_config.json"
DEFAULT_INPUT_DIR = SCRIPT_DIR / "daily_report_inputs"
DEFAULT_ENV = SCRIPT_DIR.parent / ".env"

LABEL_MAP = {
    ("小红书", "直播"): "小红书直播",
    ("抖音", "直播间"): "抖音直播",
    ("小红书", "商品卡"): "商品卡",
    ("抖音", "商品卡"): "商品卡",
    ("视频号", "大号"): "视频号大号直播",
    ("视频号", "小号"): "视频号小号直播",
    ("有赞", "商城"): "有赞",
}

ROW_ORDER = [
    ("小红书直播", "小红书直播", "小红书直播"),
    ("抖音直播", "抖音直播间", "抖音直播"),
    ("商品卡", "商品卡", "商品卡"),
    ("视频号大号直播", "视频号大号", "视频号大号直播"),
    ("视频号小号直播", "视频号小号", "视频号小号直播"),
    ("有赞", "有赞", "有赞"),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="重算并替换飞书 Base 的月度平台GMV完成情况表")
    parser.add_argument("--date", required=True, help="目标日期，格式 YYYY-MM-DD")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="配置文件路径")
    parser.add_argument("--input-dir", default=str(DEFAULT_INPUT_DIR), help="daily_report_inputs 目录")
    parser.add_argument("--env", default=str(DEFAULT_ENV), help=".env 路径")
    parser.add_argument("--dry-run", action="store_true", help="只输出将写入的数据，不实际修改飞书")
    return parser.parse_args()


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def round2(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.01")))


def round6(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.000001")))


def iter_month_input_paths(input_dir: Path, target_date: datetime) -> list[Path]:
    paths = []
    for path in sorted(input_dir.glob("20??-??-??.json")):
        try:
            day = datetime.strptime(path.stem, "%Y-%m-%d")
        except ValueError:
            continue
        if day.year == target_date.year and day.month == target_date.month and day.date() <= target_date.date():
            paths.append(path)
    return paths


def build_records(config: dict, input_dir: Path, target_date: datetime) -> list[dict]:
    agg: dict[str, dict[str, Decimal]] = {}
    total_gmv = Decimal("0")
    total_refund = Decimal("0")
    total_actual = Decimal("0")

    for path in iter_month_input_paths(input_dir, target_date):
        payload = load_json(path)
        for entry in payload.get("entries", []):
            key = LABEL_MAP.get((entry.get("platform"), entry.get("account")))
            if not key:
                continue
            row = agg.setdefault(key, {"gmv": Decimal("0"), "refund": Decimal("0"), "actual": Decimal("0")})
            gmv = Decimal(str(entry.get("gmv", 0) or 0))
            refund = Decimal(str(entry.get("refund", 0) or 0))
            actual = Decimal(str(entry.get("actual", 0) or 0))
            row["gmv"] += gmv
            row["refund"] += refund
            row["actual"] += actual
            total_gmv += gmv
            total_refund += refund
            total_actual += actual

    targets = config.get("month_targets", {}) or {}
    records: list[dict] = []
    for platform_label, target_key, remark_label in ROW_ORDER:
        row = agg.get(platform_label, {"gmv": Decimal("0"), "refund": Decimal("0"), "actual": Decimal("0")})
        goal = Decimal(str(targets.get(target_key, 0) or 0))
        actual = row["actual"]
        records.append(
            {
                "平台": platform_label,
                "成交金额(GMV)": round2(row["gmv"]),
                "退款金额": round2(row["refund"]),
                "实际成交额": round2(actual),
                "月度目标": round2(goal),
                "完成度": round6(actual / goal) if goal else 0,
                "描述（1）": f"{remark_label}累计至{target_date.month}/{target_date.day}",
            }
        )

    month_goal = Decimal(str(config.get("month_goal", 0) or 0))
    records.append(
        {
            "平台": "月度合计",
            "成交金额(GMV)": round2(total_gmv),
            "退款金额": round2(total_refund),
            "实际成交额": round2(total_actual),
            "月度目标": round2(month_goal),
            "完成度": round6(total_actual / month_goal) if month_goal else 0,
            "描述（1）": f"累计至{target_date.month}/{target_date.day}",
        }
    )
    return records


def main() -> int:
    args = parse_args()
    target_date = datetime.strptime(args.date, "%Y-%m-%d")
    config = load_json(Path(args.config))
    base_token = str(config.get("base_token", "")).strip()
    table_name = str((config.get("base_tables") or {}).get("month_progress") or "6月平台GMV完成情况").strip()
    records = build_records(config, Path(args.input_dir), target_date)

    if args.dry_run:
        print(json.dumps({"table": table_name, "records": records}, ensure_ascii=False, indent=2))
        return 0

    class Args:
        env = args.env
        app_id = ""
        app_secret = ""
        token_cache = ""

    app_id, app_secret = base_sync.resolve_app_credentials(Args())
    token = base_sync.get_tenant_access_token(app_id, app_secret)
    table_items = base_sync.list_tables(token, base_token)
    table_id = None
    for item in table_items:
        if str(item.get("name") or item.get("table_name") or "").strip() == table_name:
            table_id = str(item.get("table_id") or item.get("id") or "").strip()
            break
    if not table_id:
        raise RuntimeError(f"未找到目标表：{table_name}")

    field_names = base_sync.resolve_field_names(token, base_token, table_id)
    result = base_sync.replace_table(token, base_token, table_id, records, field_names, table_name)
    print(json.dumps({"table": table_name, "result": result, "records": records}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
