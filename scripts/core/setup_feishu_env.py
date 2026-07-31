#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Any
from urllib import error, request


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent.parent
CONFIG_DIR = SKILL_ROOT / "config"

ENV_KEYS = [
    ("BUSHAN_FEISHU_APP_ID", "Feishu app id"),
    ("BUSHAN_FEISHU_APP_SECRET", "Feishu app secret"),
    ("BUSHAN_FEISHU_FOLDER_TOKEN", "Feishu root folder token for daily uploads"),
    ("BUSHAN_REPORT_WEBHOOK", "Feishu bot webhook"),
    ("BUSHAN_BASE_TOKEN", "Feishu Base token"),
    ("BUSHAN_REPORT_OPERATOR_OPEN_ID", "Feishu operator open id"),
    ("BUSHAN_LARK_CLI_PATH", "Local lark-cli path"),
    ("BUSHAN_MONTH_WORKBOOK", "Workbook template path"),
]

PLATFORM_FOLDER_KEYS = [
    ("小红书", "Xiaohongshu folder token"),
    ("抖音", "Douyin folder token"),
    ("有赞", "Youzan folder token"),
    ("视频号大号", "WeChat Video big-account folder token"),
    ("视频号小号", "WeChat Video small-account folder token"),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Interactive setup wizard for Bushan Feishu configuration.")
    parser.add_argument("--output-dir", default=str(CONFIG_DIR), help="Directory to write generated config files")
    parser.add_argument("--env-file", default=".env", help="Generated env file name")
    parser.add_argument("--config-file", default="daily_report_pipeline_config.json", help="Generated pipeline config file name")
    parser.add_argument("--folder-map-file", default="folder_mapping.json", help="Generated folder mapping file name")
    parser.add_argument("--month-targets-file", default="month_targets.json", help="Generated month targets file name")
    parser.add_argument("--skip-validate", action="store_true", help="Skip live validation of Feishu app credentials")
    return parser.parse_args()


def prompt(label: str, default: str | None = None, secret: bool = False) -> str:
    suffix = f" [{default}]" if default else ""
    raw = input(f"{label}{suffix}: ").strip()
    if raw:
        return raw
    return default or ""


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def request_json(url: str, payload: dict[str, Any]) -> dict[str, Any]:
    body = json.dumps(payload).encode("utf-8")
    req = request.Request(
        url,
        data=body,
        method="POST",
        headers={"Content-Type": "application/json; charset=utf-8"},
    )
    try:
        with request.urlopen(req, timeout=20) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code}: {detail}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"Network error: {exc.reason}") from exc


def validate_app_credentials(app_id: str, app_secret: str) -> tuple[bool, str]:
    try:
        result = request_json(
            "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
            {"app_id": app_id, "app_secret": app_secret},
        )
    except Exception as exc:
        return False, str(exc)
    code = result.get("code")
    if code not in (None, 0):
        return False, f"code={code}, msg={result.get('msg')}"
    token = result.get("tenant_access_token") or result.get("data", {}).get("tenant_access_token")
    if not token:
        return False, "response did not contain tenant access token"
    return True, "tenant access token acquired"


def main() -> int:
    args = parse_args()
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    env_template = load_json(CONFIG_DIR / "daily_report_pipeline_config.template.json")
    folder_template = load_json(CONFIG_DIR / "folder_mapping.example.json")
    targets_template = load_json(CONFIG_DIR / "month_targets.example.json")

    print("Bushan Feishu setup wizard")
    print("This will generate .env, daily report config, folder mapping, and month target files.")
    print()

    env_values: dict[str, str] = {}
    env_defaults = {
        "BUSHAN_LARK_CLI_PATH": os.environ.get("BUSHAN_LARK_CLI_PATH", str(Path.home() / ".npm-global/bin/lark-cli")),
        "BUSHAN_MONTH_WORKBOOK": os.environ.get("BUSHAN_MONTH_WORKBOOK", "./assets/templates/month_workbook.xlsx"),
    }
    for key, label in ENV_KEYS:
        env_values[key] = prompt(label, env_defaults.get(key) or os.environ.get(key, ""))

    folder_mapping = dict(folder_template)
    folder_mapping["local_root"] = prompt("Local organized data root", folder_mapping.get("local_root", "./data"))
    folder_mapping["feishu_root_folder_token"] = prompt(
        "Feishu root folder token",
        env_values.get("BUSHAN_FEISHU_FOLDER_TOKEN") or folder_mapping.get("feishu_root_folder_token", ""),
    )
    platform_folders: dict[str, str] = {}
    for platform, label in PLATFORM_FOLDER_KEYS:
        platform_folders[platform] = prompt(label, folder_mapping.get("platform_folders", {}).get(platform, ""))
    folder_mapping["platform_folders"] = platform_folders

    report_config = dict(env_template)
    report_config["webhook"] = env_values["BUSHAN_REPORT_WEBHOOK"]
    report_config["base_token"] = env_values["BUSHAN_BASE_TOKEN"]
    report_config["workbook_path"] = env_values["BUSHAN_MONTH_WORKBOOK"]
    report_config["month_goal"] = float(prompt("Month goal", str(report_config.get("month_goal", 4400000))))
    month_targets: dict[str, float] = {}
    current_targets = targets_template.get("month_targets", report_config.get("month_targets", {}))
    for key, current_value in current_targets.items():
        month_targets[key] = float(prompt(f"Month target for {key}", str(current_value)))
    report_config["month_targets"] = month_targets
    month_targets_payload = {"month_goal": report_config["month_goal"], "month_targets": month_targets}

    env_path = output_dir / args.env_file
    config_path = output_dir / args.config_file
    folder_map_path = output_dir / args.folder_map_file
    targets_path = output_dir / args.month_targets_file

    env_lines = [f"{key}={value}" for key, value in env_values.items()]
    env_path.write_text("\n".join(env_lines) + "\n", encoding="utf-8")
    write_json(config_path, report_config)
    write_json(folder_map_path, folder_mapping)
    write_json(targets_path, month_targets_payload)

    print()
    print(f"Wrote env file: {env_path}")
    print(f"Wrote report config: {config_path}")
    print(f"Wrote folder mapping: {folder_map_path}")
    print(f"Wrote month targets: {targets_path}")

    if not args.skip_validate:
        print()
        print("Validating Feishu app credentials...")
        ok, detail = validate_app_credentials(
            env_values["BUSHAN_FEISHU_APP_ID"],
            env_values["BUSHAN_FEISHU_APP_SECRET"],
        )
        if ok:
            print(f"Validation succeeded: {detail}")
        else:
            print(f"Validation failed: {detail}")
            print("Config files were still written. Fix the values and rerun if needed.")
    else:
        print()
        print("Skipped live validation.")

    print()
    print("Next recommended steps:")
    print("1. Place your workbook template at the configured workbook path.")
    print("2. Confirm Feishu folder tokens in folder_mapping.json.")
    print("3. Run the local ingest or month resume entrypoint.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
