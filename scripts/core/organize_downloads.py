#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


DATE_PATTERNS = [
    re.compile(r"(20\d{2}-\d{2}-\d{2})"),
    re.compile(r"(20\d{6})"),
]

DATA_SUFFIXES = {".xlsx", ".xls", ".csv"}


@dataclass(frozen=True)
class PlatformRule:
    platform: str
    includes: tuple[str, ...]
    excludes: tuple[str, ...] = ()


RULES = [
    PlatformRule("小红书", ("商家成交数据概览", "商家经营数据总览", "载体构成账号列表")),
    PlatformRule("抖音", ("抖音电商罗盘-成交分析", "成交分析")),
    PlatformRule("有赞", ("Order_youzan", "Refund_youzan")),
    PlatformRule("视频号大号", ("场景构成", "成交数据"), excludes=("小号",)),
    PlatformRule("视频号小号", ("场景构成",), excludes=("大号", "B1", "B2", "b1", "b2")),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Organize downloaded raw files into Bushan date/platform folders.")
    parser.add_argument("--source-dir", required=True, help="Directory containing newly downloaded raw files")
    parser.add_argument("--target-root", required=True, help="Root directory for organized YYYY-MM-DD folders")
    parser.add_argument("--date", help="Override target date, format YYYY-MM-DD")
    parser.add_argument("--copy", action="store_true", help="Copy files instead of moving them")
    parser.add_argument("--dry-run", action="store_true", help="Preview actions only")
    parser.add_argument("--manifest", help="Optional path to write a JSON manifest of the organization result")
    return parser.parse_args()


def detect_date(name: str, override: str | None) -> str | None:
    if override:
        return override
    for pattern in DATE_PATTERNS:
        match = pattern.search(name)
        if not match:
            continue
        value = match.group(1)
        if len(value) == 8 and value.isdigit():
            return f"{value[:4]}-{value[4:6]}-{value[6:]}"
        return value
    return None


def detect_platform(name: str) -> str | None:
    lower_name = name.lower()
    for rule in RULES:
        if not any(token.lower() in lower_name for token in rule.includes):
            continue
        if any(token.lower() in lower_name for token in rule.excludes):
            continue
        return rule.platform
    return None


def iter_data_files(source_dir: Path) -> Iterable[Path]:
    for path in sorted(source_dir.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith(".") or path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in DATA_SUFFIXES:
            continue
        yield path


def main() -> int:
    args = parse_args()
    source_dir = Path(args.source_dir).expanduser().resolve()
    target_root = Path(args.target_root).expanduser().resolve()

    if not source_dir.exists():
        raise SystemExit(f"Source directory does not exist: {source_dir}")

    manifest: dict[str, list[dict[str, str]]] = {"organized": [], "skipped": []}

    for path in iter_data_files(source_dir):
        detected_date = detect_date(path.name, args.date)
        detected_platform = detect_platform(path.name)
        if not detected_date or not detected_platform:
            manifest["skipped"].append(
                {
                    "file": str(path),
                    "reason": f"date={detected_date or 'unknown'}, platform={detected_platform or 'unknown'}",
                }
            )
            continue

        target_dir = target_root / detected_date / detected_platform
        target_path = target_dir / path.name
        manifest["organized"].append(
            {
                "file": str(path),
                "date": detected_date,
                "platform": detected_platform,
                "target": str(target_path),
                "mode": "copy" if args.copy else "move",
            }
        )

        if args.dry_run:
            continue

        target_dir.mkdir(parents=True, exist_ok=True)
        if args.copy:
            shutil.copy2(path, target_path)
        else:
            shutil.move(str(path), str(target_path))

    if args.manifest:
        manifest_path = Path(args.manifest).expanduser().resolve()
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
