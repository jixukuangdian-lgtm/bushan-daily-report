#!/usr/bin/env python3
import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any, Dict
from urllib import error, request


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = SCRIPT_DIR / "daily_report_pipeline_config.json"
DEFAULT_OUTPUT_DIR = SCRIPT_DIR / "daily_report_outputs"
FREQUENCY_LIMIT_CODE = "11232"


def load_json(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"文件不存在：{path}")
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_card_path(output_dir: Path, date_text: str) -> Path:
    normalized = date_text.replace("-", "_")
    return output_dir / f"card_{normalized}.json"


def post_webhook(webhook: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = request.Request(
        webhook,
        data=body,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )

    try:
        with request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"发送失败，HTTP {exc.code}：{detail}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"发送失败，网络错误：{exc.reason}") from exc
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"发送后返回的不是合法 JSON：{exc}") from exc


def ensure_webhook_success(result: Dict[str, Any]) -> None:
    code = result.get("code")
    if code in (None, 0, "0"):
        return

    msg = str(result.get("msg", "")).strip() or "未知错误"
    raise RuntimeError(f"发送失败，飞书返回 code={code}，msg={msg}")


def is_frequency_limited(result: Dict[str, Any]) -> bool:
    code = str(result.get("code", "")).strip()
    msg = str(result.get("msg", "")).strip().lower()
    return code == FREQUENCY_LIMIT_CODE or "frequency limited" in msg


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="单独发送已生成的飞书日报卡片 JSON")
    parser.add_argument("--date", help="日期，格式 YYYY-MM-DD，例如 2026-05-19")
    parser.add_argument("--card", help="卡片 JSON 的绝对路径或相对路径")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="配置文件路径")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="日报输出目录")
    parser.add_argument("--max-retries", type=int, default=4, help="飞书限流时最多重试次数，默认 4")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        if not args.date and not args.card:
            raise ValueError("请至少提供 --date 或 --card")

        config = load_json(Path(args.config))
        webhook = str(config.get("webhook", "")).strip()
        if not webhook.startswith("https://"):
            raise ValueError(f"webhook 不合法：{webhook}")

        if args.card:
            card_path = Path(args.card).expanduser().resolve()
        else:
            card_path = resolve_card_path(Path(args.output_dir), args.date)

        payload = load_json(card_path)
        print(f"准备发送卡片：{card_path}")

        attempt = 0
        while True:
            result = post_webhook(webhook, payload)
            if not is_frequency_limited(result):
                ensure_webhook_success(result)
                break

            if attempt >= args.max_retries:
                ensure_webhook_success(result)

            wait_seconds = min(60, 15 * (2 ** attempt))
            print(
                f"命中飞书限流 code={result.get('code')}，"
                f"{wait_seconds} 秒后重试（{attempt + 1}/{args.max_retries}）..."
            )
            time.sleep(wait_seconds)
            attempt += 1

        print("发送成功：")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print(f"执行失败：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
