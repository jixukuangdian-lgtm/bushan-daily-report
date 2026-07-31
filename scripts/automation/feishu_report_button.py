#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import signal
import subprocess
import sys
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo


PROJECT_DIR = Path(__file__).resolve().parents[2]
LARK_CLI = Path(os.environ.get("BUSHAN_LARK_CLI_PATH", "~/.npm-global/bin/lark-cli")).expanduser()
RUNNER = Path(os.environ.get("BUSHAN_REPORT_RUNNER", str(PROJECT_DIR / "scripts" / "entrypoints" / "run_month_resume.sh"))).expanduser()
LOG_DIR = PROJECT_DIR / "runtime" / "logs"
STATE_DIR = PROJECT_DIR / "runtime" / "state"
AUTHORIZED_OPERATOR = os.environ.get("BUSHAN_REPORT_OPERATOR_OPEN_ID", "YOUR_OPERATOR_OPEN_ID")
ACTION_NAME = "run_yesterday_report"
TIMEZONE = ZoneInfo("Asia/Shanghai")

stop_requested = threading.Event()
run_lock = threading.Lock()
event_process: subprocess.Popen[str] | None = None


def log(message: str) -> None:
    stamp = datetime.now(TIMEZONE).strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{stamp}] {message}", flush=True)


def target_date() -> str:
    return (datetime.now(TIMEZONE).date() - timedelta(days=1)).isoformat()


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def save_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def parse_action_value(raw: str) -> dict[str, Any]:
    if not raw:
        return {}
    try:
        value = json.loads(raw)
    except json.JSONDecodeError:
        return {}
    return value if isinstance(value, dict) else {}


def status_card(
    *,
    title: str,
    subtitle: str,
    date_text: str,
    detail: str,
    template: str,
    tag_text: str,
    tag_color: str,
    show_button: bool = False,
    button_text: str = "重新尝试",
) -> dict[str, Any]:
    elements: list[dict[str, Any]] = [
        {
            "tag": "column_set",
            "flex_mode": "none",
            "columns": [
                {
                    "tag": "column",
                    "width": "weighted",
                    "weight": 1,
                    "background_style": f"{tag_color}-50",
                    "padding": "12px",
                    "vertical_spacing": "4px",
                    "elements": [
                        {
                            "tag": "markdown",
                            "content": f"**处理日期：{date_text}**",
                        },
                        {
                            "tag": "markdown",
                            "content": f"<font color='grey'>{detail}</font>",
                            "text_size": "notation",
                        },
                    ],
                }
            ],
        }
    ]
    if show_button:
        elements.append(
            {
                "tag": "button",
                "element_id": "retry_report",
                "text": {"tag": "plain_text", "content": button_text},
                "type": "primary_filled",
                "size": "large",
                "width": "fill",
                "behaviors": [
                    {
                        "type": "callback",
                        "value": {"action": ACTION_NAME, "version": 1},
                    }
                ],
            }
        )
    return {
        "schema": "2.0",
        "config": {
            "update_multi": True,
            "width_mode": "default",
            "enable_forward": False,
            "summary": {"content": title},
        },
        "header": {
            "title": {"tag": "plain_text", "content": title},
            "subtitle": {"tag": "plain_text", "content": subtitle},
            "template": template,
            "icon": {"tag": "standard_icon", "token": "chart_colorful"},
            "text_tag_list": [
                {
                    "tag": "text_tag",
                    "text": {"tag": "plain_text", "content": tag_text},
                    "color": tag_color,
                }
            ],
        },
        "body": {
            "direction": "vertical",
            "padding": "12px 12px 20px 12px",
            "vertical_spacing": "12px",
            "elements": elements,
        },
    }


def update_card(token: str, card: dict[str, Any]) -> bool:
    payload = json.dumps(
        {"token": token, "card": card},
        ensure_ascii=False,
        separators=(",", ":"),
    )
    result = subprocess.run(
        [
            str(LARK_CLI),
            "api",
            "POST",
            "/open-apis/interactive/v1/card/update",
            "--as",
            "bot",
            "--data",
            payload,
        ],
        cwd=PROJECT_DIR,
        text=True,
        capture_output=True,
        timeout=30,
    )
    if result.returncode != 0:
        log(f"卡片更新失败：{result.stderr.strip() or result.stdout.strip()}")
        return False
    return True


def already_completed(date_text: str) -> bool:
    state_file = PROJECT_DIR / ".run_state" / f"{date_text}.state"
    try:
        return int(state_file.read_text(encoding="utf-8").strip()) >= 4
    except (FileNotFoundError, ValueError):
        return False


def run_report(event: dict[str, Any]) -> None:
    token = str(event.get("token", ""))
    date_text = target_date()
    if not run_lock.acquire(blocking=False):
        update_card(
            token,
            status_card(
                title="昨日日报正在生成",
                subtitle="已有一个任务在执行",
                date_text=date_text,
                detail="请等待当前任务完成，不会重复执行。",
                template="yellow",
                tag_text="处理中",
                tag_color="yellow",
            ),
        )
        return

    try:
        if already_completed(date_text):
            update_card(
                token,
                status_card(
                    title="昨日日报已经完成",
                    subtitle="无需重复执行",
                    date_text=date_text,
                    detail="该日期的四个步骤均已完成。",
                    template="green",
                    tag_text="已完成",
                    tag_color="green",
                    show_button=True,
                    button_text="下次上传后再次生成",
                ),
            )
            return

        update_card(
            token,
            status_card(
                title="正在生成昨日日报",
                subtitle="请稍候，不需要打开终端",
                date_text=date_text,
                detail="正在检查并处理五个平台数据，完成后会更新本卡片。",
                template="carmine",
                tag_text="处理中",
                tag_color="carmine",
            ),
        )

        LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_path = LOG_DIR / f"report-{date_text}.log"
        started_at = datetime.now(TIMEZONE).isoformat()
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(f"\n===== {started_at} button trigger =====\n")
            handle.flush()
            result = subprocess.run(
                [str(RUNNER), "--date", date_text],
                cwd=PROJECT_DIR,
                stdout=handle,
                stderr=subprocess.STDOUT,
                text=True,
            )

        ledger_path = STATE_DIR / f"{date_text}.json"
        save_json(
            ledger_path,
            {
                "date": date_text,
                "started_at": started_at,
                "finished_at": datetime.now(TIMEZONE).isoformat(),
                "exit_code": result.returncode,
                "log_path": str(log_path),
            },
        )

        if result.returncode == 0 and already_completed(date_text):
            final = status_card(
                title="昨日日报生成完成",
                subtitle="数据、工作簿、Base 和日报卡片均已处理",
                date_text=date_text,
                detail="本次任务执行成功；同一天再次点击不会重复执行。",
                template="green",
                tag_text="成功",
                tag_color="green",
                show_button=True,
                button_text="下次上传后再次生成",
            )
            log(f"{date_text} 日报执行成功")
        else:
            final = status_card(
                title="昨日日报生成失败",
                subtitle="已保留日志，可安全续跑",
                date_text=date_text,
                detail=f"执行未完成（退出码 {result.returncode}）。请检查数据文件或运行日志后重新尝试。",
                template="red",
                tag_text="失败",
                tag_color="red",
                show_button=True,
            )
            log(f"{date_text} 日报执行失败，退出码 {result.returncode}")
        update_card(token, final)
    except Exception as exc:
        log(f"处理按钮事件异常：{exc}")
        if token:
            update_card(
                token,
                status_card(
                    title="昨日日报生成失败",
                    subtitle="后台程序遇到异常",
                    date_text=date_text,
                    detail="异常信息已写入本地日志，请稍后重新尝试。",
                    template="red",
                    tag_text="失败",
                    tag_color="red",
                    show_button=True,
                ),
            )
    finally:
        run_lock.release()


def handle_event(event: dict[str, Any]) -> None:
    if event.get("type") != "card.action.trigger":
        return
    if event.get("operator_id") != AUTHORIZED_OPERATOR:
        log(f"忽略未授权用户操作：{event.get('operator_id', '')}")
        return
    if event.get("action_tag") != "button":
        return
    action = parse_action_value(str(event.get("action_value", "")))
    if action.get("action") != ACTION_NAME:
        return
    threading.Thread(target=run_report, args=(event,), daemon=True).start()


def consume_events() -> int:
    global event_process
    env = os.environ.copy()
    env["LARKSUITE_CLI_NO_UPDATE_NOTIFIER"] = "1"
    env["LARKSUITE_CLI_NO_SKILLS_NOTIFIER"] = "1"
    event_process = subprocess.Popen(
        [
            str(LARK_CLI),
            "event",
            "consume",
            "card.action.trigger",
            "--as",
            "bot",
        ],
        cwd=PROJECT_DIR,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1,
        env=env,
    )

    def relay_stderr() -> None:
        assert event_process is not None
        assert event_process.stderr is not None
        for line in event_process.stderr:
            log(line.rstrip())

    threading.Thread(target=relay_stderr, daemon=True).start()
    assert event_process.stdout is not None
    for line in event_process.stdout:
        if stop_requested.is_set():
            break
        try:
            payload = json.loads(line)
        except json.JSONDecodeError:
            log("忽略无法解析的事件")
            continue
        handle_event(payload)

    if event_process.poll() is None:
        event_process.terminate()
    return event_process.wait(timeout=15)


def request_stop(signum: int, _frame: Any) -> None:
    stop_requested.set()
    log(f"收到停止信号 {signum}")
    if event_process is not None and event_process.poll() is None:
        event_process.terminate()


def main() -> int:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    signal.signal(signal.SIGTERM, request_stop)
    signal.signal(signal.SIGINT, request_stop)
    if not LARK_CLI.exists():
        log(f"未找到 lark-cli：{LARK_CLI}")
        return 1
    if not RUNNER.exists():
        log(f"未找到日报脚本：{RUNNER}")
        return 1

    log("飞书日报按钮监听器启动")
    while not stop_requested.is_set():
        try:
            exit_code = consume_events()
        except Exception as exc:
            log(f"事件监听异常：{exc}")
            exit_code = 1
        if stop_requested.is_set():
            break
        log(f"事件监听退出（{exit_code}），5 秒后重连")
        time.sleep(5)
    log("飞书日报按钮监听器停止")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
