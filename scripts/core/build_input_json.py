#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from urllib import error, parse, request

import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parent
INPUT_DIR = SCRIPT_DIR / "daily_report_inputs"
RUN_SCRIPT = SCRIPT_DIR / "run_daily_report.sh"
YOUZAN_CUMULATIVE_DIR = SCRIPT_DIR / "youzan_cumulative_sources"

DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")
DATA_SUFFIXES = (".xlsx", ".xls", ".csv")
PLATFORM_ENTRY_MAP = {
    "小红书": [("小红书", "直播"), ("小红书", "商品卡")],
    "抖音": [("抖音", "直播间"), ("抖音", "商品卡")],
    "有赞": [("有赞", "商城")],
    "视频号大号": [("视频号", "大号")],
    "视频号小号": [("视频号", "小号")],
}
PLATFORM_ORDER = ["小红书", "抖音", "有赞", "视频号大号", "视频号小号"]
ONLINE_DOC_EXPORT_EXT = {
    "doc": "docx",
    "docx": "docx",
    "sheet": "xlsx",
    "bitable": "xlsx",
    "slides": "pdf",
    "mindnote": "pdf",
}


def load_env() -> Tuple[Dict[str, str], Path]:
    candidates = [
        SCRIPT_DIR.parent / ".env",
        SCRIPT_DIR / ".env",
        Path.cwd() / ".env",
        Path.cwd() / "config" / ".env",
    ]

    env_path = next((path for path in candidates if path.exists()), None)
    if env_path is None:
        searched = "\n".join(f"- {path}" for path in candidates)
        raise RuntimeError(f"没有找到 .env 文件，已检查这些位置：\n{searched}")

    env: Dict[str, str] = {}
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        env[key.strip()] = value.strip().strip('"').strip("'")

    required = ["FEISHU_APP_ID", "FEISHU_APP_SECRET", "FEISHU_FOLDER_TOKEN"]
    missing = [key for key in required if not env.get(key)]
    if missing:
        raise RuntimeError(
            f".env 缺少必填项：{', '.join(missing)}，当前读取文件：{env_path}"
        )

    return env, env_path


def request_json(
    url: str,
    *,
    method: str = "GET",
    payload: Optional[Dict] = None,
    headers: Optional[Dict[str, str]] = None,
) -> Dict:
    body = None
    req_headers = {"Content-Type": "application/json; charset=utf-8"}
    if headers:
        req_headers.update(headers)
    if payload is not None:
        body = json.dumps(payload).encode("utf-8")

    req = request.Request(url, data=body, headers=req_headers, method=method)

    try:
        with request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} 调用失败：{detail}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"网络请求失败：{exc.reason}") from exc
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"接口返回不是合法 JSON：{exc}") from exc


def get_tenant_access_token(app_id: str, app_secret: str) -> str:
    data = request_json(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        method="POST",
        payload={"app_id": app_id, "app_secret": app_secret},
    )
    if data.get("code") not in (None, 0):
        raise RuntimeError(
            f"获取 tenant_access_token 失败：code={data.get('code')} msg={data.get('msg')}"
        )

    token = data.get("tenant_access_token") or data.get("data", {}).get("tenant_access_token")
    if not token:
        raise RuntimeError(f"接口未返回 tenant_access_token：{data}")
    return token


def list_folder_items(folder_token: str, tenant_access_token: str) -> List[Dict]:
    items: List[Dict] = []
    page_token = ""

    while True:
        query = {
            "folder_token": folder_token,
            "page_size": 200,
            "order_by": "EditedTime",
            "direction": "DESC",
        }
        if page_token:
            query["page_token"] = page_token

        url = "https://open.feishu.cn/open-apis/drive/v1/files?" + parse.urlencode(query)
        data = request_json(
            url,
            headers={"Authorization": f"Bearer {tenant_access_token}"},
        )

        if data.get("code") not in (None, 0):
            raise RuntimeError(
                f"读取文件夹列表失败：code={data.get('code')} msg={data.get('msg')}"
            )

        payload = data.get("data", {})
        files = payload.get("files", [])
        if not isinstance(files, list):
            raise RuntimeError(f"接口返回的 files 字段格式异常：{data}")
        items.extend(files)

        has_more = payload.get("has_more") or payload.get("has_next_page") or False
        page_token = payload.get("page_token") or payload.get("next_page_token") or ""
        if not has_more:
            break

    return items


def collect_token_candidates(item: Dict) -> List[str]:
    candidates: List[str] = []

    def add_value(value: object) -> None:
        if isinstance(value, str) and value and value not in candidates:
            candidates.append(value)

    add_value(item.get("token"))
    add_value(item.get("folder_token"))
    add_value(item.get("file_token"))
    for key, value in item.items():
        if "token" in key.lower():
            add_value(value)
    return candidates


def list_child_items(folder_item: Dict, tenant_access_token: str) -> Tuple[List[Dict], str]:
    token_candidates = collect_token_candidates(folder_item)
    if not token_candidates:
        raise RuntimeError(f"文件夹缺少可用 token：{folder_item}")

    last_result: List[Dict] = []
    for candidate in token_candidates:
        items = list_folder_items(candidate, tenant_access_token)
        if items:
            return items, candidate
        last_result = items
    return last_result, token_candidates[0]


def find_date_folder(items: List[Dict], target_date: Optional[str]) -> Optional[Dict]:
    date_folders: List[Dict] = []
    for item in items:
        if item.get("type") != "folder":
            continue
        name = str(item.get("name", "")).strip()
        if not DATE_PATTERN.match(name):
            continue
        try:
            datetime.strptime(name, "%Y-%m-%d")
        except ValueError:
            continue
        date_folders.append(item)

    if not date_folders:
        return None

    date_folders.sort(key=lambda item: item.get("name", ""), reverse=True)
    if target_date:
        for item in date_folders:
            if item.get("name") == target_date:
                return item
        return None
    return date_folders[0]


def get_folder_items(items: List[Dict]) -> List[Dict]:
    folders: List[Dict] = []
    for item in items:
        if item.get("type") != "folder":
            continue
        name = str(item.get("name", "")).strip()
        if name:
            folders.append(item)
    return folders


def is_data_file(item: Dict) -> bool:
    if item.get("type") == "folder":
        return False
    if is_online_document(item):
        return True
    name = str(item.get("name", "")).strip()
    if not name or name.startswith("."):
        return False
    lower_name = name.lower()
    return lower_name.endswith(DATA_SUFFIXES)


def round2(value: float) -> float:
    return round(float(value or 0), 2)


def parse_number(value) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("¥", "").replace("%", "")
    return float(text) if text else 0.0


def normalize_day(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and float(value).is_integer():
        text = str(int(value))
        if len(text) >= 8 and text[:8].isdigit():
            return datetime.strptime(text[:8], "%Y%m%d").date()
    text = str(value).strip()
    if len(text) >= 8 and text[:8].isdigit():
        return datetime.strptime(text[:8], "%Y%m%d").date()
    return pd.to_datetime(value).date()


def extract_date_from_text(text: str) -> date:
    match = re.search(r"(20\d{2})[-_]?(\d{2})[-_]?(\d{2})", text)
    if not match:
        raise RuntimeError(f"无法从文本中识别日期：{text}")
    year, month, day = match.groups()
    return date(int(year), int(month), int(day))


def sort_platform_folders(items: List[Dict]) -> List[Dict]:
    order_map = {name: index for index, name in enumerate(PLATFORM_ORDER)}
    folders = get_folder_items(items)
    return sorted(
        folders,
        key=lambda item: (
            order_map.get(str(item.get("name", "")).strip(), 999),
            str(item.get("name", "")).strip(),
        ),
    )


def collect_platform_files(platform_folders: List[Dict], tenant_access_token: str) -> Dict[str, List[str]]:
    result: Dict[str, List[str]] = {}

    for folder in platform_folders:
        platform_name = str(folder.get("name", "")).strip()
        child_items, _ = list_child_items(folder, tenant_access_token)
        file_names = sorted(str(item.get("name", "")).strip() for item in child_items if is_data_file(item))
        result[platform_name] = [name for name in file_names if name]
    return result


def get_item_mime_type(item: Dict) -> str:
    mime_type = (
        item.get("mime_type")
        or item.get("mimeType")
        or item.get("file_mime_type")
        or ""
    )
    return str(mime_type).strip()


def print_item_metadata(platform_name: str, item: Dict) -> None:
    print(
        f"[{platform_name}]"
        f" name={item.get('name', '')}"
        f" | type={item.get('type', '')}"
        f" | token={choose_primary_token(item)}"
        f" | mime_type={get_item_mime_type(item) or '-'}"
    )


def is_online_document(item: Dict) -> bool:
    item_type = str(item.get("type", "")).strip().lower()
    if item_type in ONLINE_DOC_EXPORT_EXT:
        return True
    mime_type = get_item_mime_type(item).lower()
    if mime_type.startswith("application/vnd.larksuite") or mime_type.startswith("application/vnd.bytedance"):
        return True
    return False


def choose_primary_token(item: Dict) -> str:
    candidates = collect_token_candidates(item)
    if not candidates:
        raise RuntimeError(f"文件缺少可用 token：{item}")
    return candidates[0]


def get_tmp_download_url(file_token: str, tenant_access_token: str) -> str:
    url = "https://open.feishu.cn/open-apis/drive/v1/medias/batch_get_tmp_download_url"
    data = request_json(
        url,
        method="POST",
        payload={"file_tokens": [file_token]},
        headers={"Authorization": f"Bearer {tenant_access_token}"},
    )
    if data.get("code") not in (None, 0):
        raise RuntimeError(
            f"获取文件下载链接失败：code={data.get('code')} msg={data.get('msg')}"
        )

    items = data.get("data", {}).get("tmp_download_urls", [])
    if not items:
        raise RuntimeError(f"接口未返回下载链接，file_token={file_token}，响应={data}")
    download_url = items[0].get("tmp_download_url")
    if not download_url:
        raise RuntimeError(f"下载链接为空：{data}")
    return download_url


def create_export_task(file_token: str, file_type: str, tenant_access_token: str) -> str:
    data = request_json(
        "https://open.feishu.cn/open-apis/drive/v1/export_tasks",
        method="POST",
        payload={
            "token": file_token,
            "type": file_type,
            "file_extension": ONLINE_DOC_EXPORT_EXT[file_type],
        },
        headers={"Authorization": f"Bearer {tenant_access_token}"},
    )
    if data.get("code") not in (None, 0):
        raise RuntimeError(
            f"创建导出任务失败：code={data.get('code')} msg={data.get('msg')} token={file_token} type={file_type}"
        )
    ticket = data.get("data", {}).get("ticket")
    if not ticket:
        raise RuntimeError(f"导出任务未返回 ticket：{data}")
    return str(ticket)


def get_export_task_file_token(ticket: str, tenant_access_token: str) -> str:
    data = request_json(
        f"https://open.feishu.cn/open-apis/drive/v1/export_tasks/{ticket}",
        headers={"Authorization": f"Bearer {tenant_access_token}"},
    )
    if data.get("code") not in (None, 0):
        raise RuntimeError(
            f"查询导出任务失败：code={data.get('code')} msg={data.get('msg')} ticket={ticket}"
        )

    payload = data.get("data", {})
    file_token = (
        payload.get("file_token")
        or payload.get("result", {}).get("file_token")
        or payload.get("result", {}).get("token")
    )
    if file_token:
        return str(file_token)

    job_status = (
        payload.get("job_status")
        or payload.get("status")
        or payload.get("result", {}).get("job_status")
        or payload.get("result", {}).get("status")
    )
    raise RuntimeError(f"导出任务尚未返回文件 token：ticket={ticket} status={job_status} payload={payload}")


def export_online_document(item: Dict, tenant_access_token: str, target_path: Path) -> Path:
    file_token = choose_primary_token(item)
    file_type = str(item.get("type", "")).strip().lower()
    if file_type not in ONLINE_DOC_EXPORT_EXT:
        raise RuntimeError(f"暂不支持导出该在线文档类型：type={file_type} name={item.get('name', '')}")

    ticket = create_export_task(file_token, file_type, tenant_access_token)
    exported_file_token = get_export_task_file_token(ticket, tenant_access_token)
    download_url = f"https://open.feishu.cn/open-apis/drive/v1/export_tasks/file/{exported_file_token}/download"
    return download_file(
        download_url,
        target_path,
        headers={"Authorization": f"Bearer {tenant_access_token}"},
    )


def download_file(download_url: str, target_path: Path, headers: Optional[Dict[str, str]] = None) -> Path:
    req = request.Request(download_url, headers=headers or {}, method="GET")
    try:
        with request.urlopen(req, timeout=60) as resp:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            target_path.write_bytes(resp.read())
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"下载文件失败 HTTP {exc.code}：{detail}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"下载文件失败：{exc.reason}") from exc
    return target_path


def sanitize_filename(name: str) -> str:
    return re.sub(r"[\\\\/:*?\"<>|]", "_", name.strip())


def build_target_path(download_dir: Path, platform_name: str, item: Dict) -> Path:
    file_name = sanitize_filename(str(item.get("name", "")).strip() or "unnamed")
    item_type = str(item.get("type", "")).strip().lower()
    if is_online_document(item):
        export_ext = ONLINE_DOC_EXPORT_EXT.get(item_type)
        if export_ext and not file_name.lower().endswith(f".{export_ext}"):
            file_name = f"{file_name}.{export_ext}"
    return download_dir / platform_name / file_name


def download_file_with_fallback(
    item: Dict,
    tenant_access_token: str,
    target_path: Path,
) -> Path:
    file_name = str(item.get("name", "")).strip()
    file_token = choose_primary_token(item)
    item_type = str(item.get("type", "")).strip().lower()
    errors: List[str] = []

    if is_online_document(item):
        return export_online_document(item, tenant_access_token, target_path)

    try:
        tmp_url = get_tmp_download_url(file_token, tenant_access_token)
        return download_file(tmp_url, target_path)
    except Exception as exc:
        errors.append(f"tmp_url 匿名下载失败: {exc}")

    try:
        tmp_url = get_tmp_download_url(file_token, tenant_access_token)
        return download_file(
            tmp_url,
            target_path,
            headers={"Authorization": f"Bearer {tenant_access_token}"},
        )
    except Exception as exc:
        errors.append(f"tmp_url 鉴权下载失败: {exc}")

    fallback_urls = [
        f"https://open.feishu.cn/open-apis/drive/v1/files/{file_token}/download",
        f"https://open.feishu.cn/open-apis/drive/v1/medias/{file_token}/download",
    ]
    for url in fallback_urls:
        try:
            return download_file(
                url,
                target_path,
                headers={"Authorization": f"Bearer {tenant_access_token}"},
            )
        except Exception as exc:
            errors.append(f"直接下载失败 {url}: {exc}")

    raise RuntimeError(
        f"文件下载全部失败：{file_name} type={item_type} token={file_token}；" + "；".join(errors)
    )


def download_platform_files(
    platform_folders: List[Dict],
    tenant_access_token: str,
    download_dir: Path,
) -> Dict[str, List[Path]]:
    downloaded: Dict[str, List[Path]] = {}

    for folder in platform_folders:
        platform_name = str(folder.get("name", "")).strip()
        child_items, _ = list_child_items(folder, tenant_access_token)
        files: List[Path] = []
        for item in child_items:
            if not is_data_file(item):
                continue
            print_item_metadata(platform_name, item)
            local_path = build_target_path(download_dir, platform_name, item)
            download_file_with_fallback(item, tenant_access_token, local_path)
            files.append(local_path)
        downloaded[platform_name] = sorted(files)
    return downloaded


def build_xhs_map(path: Path) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    xls = pd.ExcelFile(path)
    preferred_sheets = [
        name for name in xls.sheet_names if name.startswith("商家成交数据概览-all")
    ] or [
        name for name in xls.sheet_names if name.startswith("商家经营数据总览-all")
    ]
    if not preferred_sheets:
        raise RuntimeError(f"小红书文件缺少可识别 sheet：{path}")
    sheet = preferred_sheets[0]
    df = pd.read_excel(path, sheet_name=sheet).fillna(0)
    live_account_sheet = next((name for name in xls.sheet_names if name.startswith("载体构成账号列表")), None)
    live_account_df = pd.read_excel(path, sheet_name=live_account_sheet).fillna(0) if live_account_sheet else None
    result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    if "时间" in df.columns:
        rows = df.iterrows()
    else:
        rows = [(0, df.iloc[0])]
    for _, row in rows:
        day = normalize_day(row["时间"]) if "时间" in df.columns else extract_date_from_text(path.name)
        new_live_cols = [
            "不山直播支付金额",
            "云南不山直播支付金额",
            "不山直播退款金额（支付时间）",
            "云南不山直播退款金额（支付时间）",
        ]
        old_live_cols = [
            "直播支付金额",
            "直播退款金额（支付时间）",
        ]
        if live_account_df is not None and {"账号名称", "载体类型", "支付金额", "支付订单数", "支付买家数", "退款金额（支付时间）"}.issubset(set(live_account_df.columns)):
            target_accounts = {"不山", "云南不山"}
            filtered_live_df = live_account_df[
                live_account_df["账号名称"].astype(str).str.strip().isin(target_accounts)
                & (live_account_df["载体类型"].astype(str).str.strip() == "直播")
            ]
            live_gmv = filtered_live_df["支付金额"].apply(parse_number).sum()
            live_refund = filtered_live_df["退款金额（支付时间）"].apply(parse_number).sum()
            live_orders = int(filtered_live_df["支付订单数"].apply(parse_number).sum())
            live_buyers = int(filtered_live_df["支付买家数"].apply(parse_number).sum())
        elif all(col in row.index for col in new_live_cols):
            live_gmv = parse_number(row["不山直播支付金额"]) + parse_number(row["云南不山直播支付金额"])
            live_refund = parse_number(row["不山直播退款金额（支付时间）"]) + parse_number(row["云南不山直播退款金额（支付时间）"])
            live_orders = int(parse_number(row.get("直播支付订单数")))
            live_buyers = int(parse_number(row.get("直播支付买家数")))
        elif all(col in row.index for col in old_live_cols):
            live_gmv = parse_number(row["直播支付金额"])
            live_refund = parse_number(row["直播退款金额（支付时间）"])
            live_orders = int(parse_number(row.get("直播支付订单数")))
            live_buyers = int(parse_number(row.get("直播支付买家数")))
        else:
            missing_live_cols = [col for col in new_live_cols if col not in row.index]
            raise RuntimeError(
                f"小红书文件既缺少新的直播口径字段，也缺少旧版直播字段；缺少新字段：{missing_live_cols}，文件：{path.name}"
            )
        card_gmv = parse_number(row["笔记支付金额"]) + parse_number(row["商卡支付金额"])
        card_refund = parse_number(row["笔记退款金额（支付时间）"]) + parse_number(row["商卡退款金额（支付时间）"])
        result[day] = {
            ("小红书", "直播"): {
                "gmv": round2(live_gmv),
                "refund": round2(live_refund),
                "actual": round2(live_gmv - live_refund),
                "orders": live_orders,
                "buyers": live_buyers,
            },
            ("小红书", "商品卡"): {
                "gmv": round2(card_gmv),
                "refund": round2(card_refund),
                "actual": round2(card_gmv - card_refund),
                "orders": int(parse_number(row.get("笔记支付订单数")) + parse_number(row.get("商卡支付订单数"))),
                "buyers": int(parse_number(row.get("笔记支付买家数")) + parse_number(row.get("商卡支付买家数"))),
            },
        }
    return result


def build_douyin_map(path: Path) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    xls = pd.ExcelFile(path)
    sheet_name = None
    for candidate in ("成交概览", "收支分析"):
        if candidate in xls.sheet_names:
            sheet_name = candidate
            break
    if sheet_name is None:
        raise RuntimeError(f"抖音文件缺少可识别 sheet，当前 sheets：{xls.sheet_names}")

    df = pd.read_excel(path, sheet_name=sheet_name).fillna(0)
    period_col = "投放时段" if "投放时段" in df.columns else "推广时段" if "推广时段" in df.columns else None
    if period_col is None:
        raise RuntimeError(f"抖音文件缺少投放/推广时段字段：{path.name}")
    df = df[df[period_col] == "不限"].copy()
    result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    for day_value, day_df in df.groupby("日期"):
        day = normalize_day(day_value)
        live_df = day_df[day_df["载体类型"] == "直播"]
        card_df = day_df[day_df["载体类型"].isin(["商品卡", "其他"])]
        live_gmv = round2(live_df["成交金额"].apply(parse_number).sum())
        live_refund = round2(live_df["成交退款金额(支付时间)"].apply(parse_number).sum())
        card_gmv = round2(card_df["成交金额"].apply(parse_number).sum())
        card_refund = round2(card_df["成交退款金额(支付时间)"].apply(parse_number).sum())
        result[day] = {
            ("抖音", "直播间"): {
                "gmv": live_gmv,
                "refund": live_refund,
                "actual": round2(live_gmv - live_refund),
                "orders": int(round(live_df["成交订单数"].apply(parse_number).sum())),
                "buyers": int(round(live_df["成交人数"].apply(parse_number).sum())),
            },
            ("抖音", "商品卡"): {
                "gmv": card_gmv,
                "refund": card_refund,
                "actual": round2(card_gmv - card_refund),
                "orders": int(round(card_df["成交订单数"].apply(parse_number).sum())),
                "buyers": int(round(card_df["成交人数"].apply(parse_number).sum())),
            },
        }
    return result


def build_video_small_map(path: Path) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    df = pd.read_excel(path).fillna(0)
    if "开播时间" in df.columns:
        df["日期"] = pd.to_datetime(df["开播时间"]).dt.date
        gmv_col = "成交金额"
        refund_col = "退款金额"
        orders_col = "成交订单数"
        buyers_col = "成交人数"
    elif "时间" in df.columns:
        df["日期"] = df["时间"].apply(normalize_day)
        gmv_col = "成交金额"
        refund_col = "成交退款金额"
        orders_col = None
        buyers_col = None
    else:
        raise RuntimeError(f"视频号小号文件缺少可识别字段：{path}")
    result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    for day, day_df in df.groupby("日期"):
        gmv = round2(day_df[gmv_col].apply(parse_number).sum())
        refund = round2(day_df[refund_col].apply(parse_number).sum())
        result[day] = {
            ("视频号", "小号"): {
                "gmv": gmv,
                "refund": refund,
                "actual": round2(gmv - refund),
                "orders": int(round(day_df[orders_col].apply(parse_number).sum())) if orders_col else None,
                "buyers": int(round(day_df[buyers_col].apply(parse_number).sum())) if buyers_col else None,
            }
        }
    return result


def build_video_scene_map(
    path: Path,
    account: str,
) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    df = pd.read_excel(path).fillna(0)
    required_columns = {"时间", "场景", "成交金额", "成交退款金额"}
    missing_columns = required_columns - set(df.columns)
    if missing_columns:
        raise RuntimeError(f"视频号{account}场景构成文件缺少字段 {sorted(missing_columns)}：{path}")

    df["日期"] = df["时间"].apply(normalize_day)
    df["场景"] = df["场景"].astype(str).str.strip()

    result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    for day, day_df in df.groupby("日期"):
        live_df = day_df[day_df["场景"] == "直播间"]
        card_df = day_df[day_df["场景"] != "直播间"]

        live_gmv = round2(live_df["成交金额"].apply(parse_number).sum())
        live_refund = round2(live_df["成交退款金额"].apply(parse_number).sum())
        card_gmv = round2(card_df["成交金额"].apply(parse_number).sum())
        card_refund = round2(card_df["成交退款金额"].apply(parse_number).sum())
        total_refund = round2(live_refund + card_refund)
        live_actual = round2(live_gmv - live_refund)
        card_actual = round2(card_gmv - card_refund)
        result[day] = {
            ("视频号", account): {
                "gmv": round2(live_gmv + card_gmv),
                "refund": total_refund,
                "actual": round2(live_actual + card_actual),
                "orders": int(round(day_df["成交订单数"].apply(parse_number).sum())) if "成交订单数" in day_df.columns else None,
                "buyers": int(round(day_df["成交人数"].apply(parse_number).sum())) if "成交人数" in day_df.columns else None,
                "live_actual": live_actual,
                "live_gmv": live_gmv,
                "live_refund": live_refund,
                "card_gmv": card_gmv,
                "card_actual": card_actual,
            }
        }
    return result


def merge_metric_maps(
    maps: List[Dict[date, Dict[Tuple[str, str], Dict]]],
) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    merged: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    metric_fields = ("gmv", "refund", "actual", "orders", "buyers", "live_actual", "live_gmv", "live_refund", "card_gmv", "card_actual")
    for item_map in maps:
        for day, payload in item_map.items():
            day_bucket = merged.setdefault(day, {})
            for key, metrics in payload.items():
                current = day_bucket.setdefault(
                    key,
                    {field: 0.0 for field in metric_fields},
                )
                for field in metric_fields:
                    value = metrics.get(field)
                    if value is None:
                        continue
                    current[field] = round2(parse_number(current.get(field)) + parse_number(value))
    for payload in merged.values():
        for metrics in payload.values():
            for count_field in ("orders", "buyers"):
                if metrics.get(count_field) is not None:
                    metrics[count_field] = int(round(metrics[count_field]))
    return merged


def build_video_small_csv_map(path: Path) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    df = pd.read_csv(path, header=1).fillna(0)
    df["日期"] = pd.to_datetime(df["直播开始时间"], format="%Y年%m月%d日 %H:%M", errors="coerce").dt.date
    result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    for day, day_df in df.groupby("日期"):
        gmv = round2(day_df["成交金额"].apply(parse_number).sum())
        result[day] = {
            ("视频号", "小号"): {
                "gmv": gmv,
                "refund": 0.0,
                "actual": gmv,
                "orders": int(round(day_df["成交订单数"].apply(parse_number).sum())),
                "buyers": int(round(day_df["成交人数"].apply(parse_number).sum())),
            }
        }
    return result


def build_video_live_csv_stats(path: Path) -> Dict[date, Dict[str, object]]:
    df = pd.read_csv(path, header=1).fillna(0)
    df["日期"] = pd.to_datetime(df["直播开始时间"], format="%Y年%m月%d日 %H:%M", errors="coerce").dt.date
    result: Dict[date, Dict[str, object]] = {}
    for day, day_df in df.groupby("日期"):
        refund_col = "退款金额" if "退款金额" in day_df.columns else None
        refund_value = round2(day_df[refund_col].apply(parse_number).sum()) if refund_col else 0.0
        result[day] = {
            "gmv": round2(day_df["成交金额"].apply(parse_number).sum()),
            "refund": refund_value,
            "orders": int(round(day_df["成交订单数"].apply(parse_number).sum())) if "成交订单数" in day_df.columns else 0,
            "buyers": int(round(day_df["成交人数"].apply(parse_number).sum())) if "成交人数" in day_df.columns else 0,
        }
    return result


def build_video_small_pair_map(total_path: Path, live_path: Path) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    total_map = build_video_small_map(total_path)
    live_map = (
        build_video_small_csv_map(live_path)
        if live_path.suffix.lower() == ".csv"
        else build_video_small_map(live_path)
    )

    result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    for day in sorted(set(total_map) | set(live_map)):
        total = total_map.get(day, {}).get(("视频号", "小号"), {})
        live = live_map.get(day, {}).get(("视频号", "小号"), {})
        total_gmv = parse_number(total.get("gmv"))
        total_refund = parse_number(total.get("refund"))
        live_gmv = parse_number(live.get("gmv"))
        live_refund = parse_number(live.get("refund"))
        if total_gmv >= live_gmv:
            card_gmv = round2(total_gmv - live_gmv)
            card_refund = round2(total_refund - live_refund)
        else:
            card_gmv = round2(total_gmv)
            card_refund = round2(total_refund)

        gmv = round2(live_gmv + card_gmv)
        refund = round2(live_refund + card_refund)
        actual = round2(gmv - refund)
        result[day] = {
            ("视频号", "小号"): {
                "gmv": gmv,
                "refund": refund,
                "actual": actual,
                "orders": live.get("orders") or total.get("orders"),
                "buyers": live.get("buyers") or total.get("buyers"),
            }
        }
    return result


def build_video_big_map(total_path: Path, live_path: Path) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    total_df = pd.read_excel(total_path).fillna(0)
    total_df["日期"] = total_df["时间"].apply(normalize_day)

    total_map: Dict[date, Dict[str, object]] = {}
    for day, day_df in total_df.groupby("日期"):
        total_map[day] = {
            "gmv": round2(day_df["成交金额"].apply(parse_number).sum()),
            "refund": round2(day_df["成交退款金额"].apply(parse_number).sum()),
            "orders": int(round(day_df["成交订单数"].apply(parse_number).sum())),
            "buyers": int(round(day_df["成交人数"].apply(parse_number).sum())),
        }

    live_map: Dict[date, Dict[str, object]] = {}
    if live_path.suffix.lower() == ".csv":
        live_map = build_video_live_csv_stats(live_path)
    else:
        live_df = pd.read_excel(live_path).fillna(0)
        live_df["日期"] = pd.to_datetime(live_df["开播时间"]).dt.date
        for day, day_df in live_df.groupby("日期"):
            live_map[day] = {
                "gmv": round2(day_df["成交金额"].apply(parse_number).sum()),
                "refund": round2(day_df["退款金额"].apply(parse_number).sum()),
                "orders": int(round(day_df["成交订单数"].apply(parse_number).sum())),
                "buyers": int(round(day_df["成交人数"].apply(parse_number).sum())),
            }

    result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    for day in sorted(set(total_map) | set(live_map)):
        total = total_map.get(day, {"gmv": 0.0, "refund": 0.0, "orders": 0, "buyers": 0})
        live = live_map.get(day, {"gmv": 0.0, "refund": 0.0, "orders": 0, "buyers": 0})
        total_gmv = float(total["gmv"])
        total_refund = float(total["refund"])
        live_gmv = float(live["gmv"])
        live_refund = float(live["refund"])
        if total_gmv >= live_gmv:
            card_gmv = round2(total_gmv - live_gmv)
            card_refund = round2(total_refund - live_refund)
        else:
            # Some exported "数据趋势" files are card-only rather than all-channel totals.
            card_gmv = round2(total_gmv)
            card_refund = round2(total_refund)
        if abs(card_gmv) < 0.01:
            card_gmv = 0.0
        if abs(card_refund) < 0.01:
            card_refund = 0.0
        result[day] = {
            ("视频号", "大号直播"): {
                "gmv": round2(live["gmv"]),
                "refund": round2(live["refund"]),
                "actual": round2(float(live["gmv"]) - float(live["refund"])),
                "orders": live["orders"],
                "buyers": live["buyers"],
            },
            ("视频号", "大号商品卡"): {
                "gmv": card_gmv,
                "refund": card_refund,
                "actual": round2(card_gmv - card_refund),
                "orders": None,
                "buyers": None,
            },
        }
    return result


def build_youzan_map(path: Path) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    if path.suffix.lower() == ".csv":
        order_csv = pd.read_csv(path).fillna(0)
        order_csv["买家付款时间"] = pd.to_datetime(order_csv["买家付款时间"], errors="coerce")
        order_csv["订单实付金额"] = pd.to_numeric(order_csv["订单实付金额"], errors="coerce").fillna(0)
        result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
        for day, day_df in order_csv.groupby(order_csv["买家付款时间"].dt.date):
            gmv = round2(day_df["订单实付金额"].sum())
            result[day] = {
                ("有赞", "商城"): {
                    "gmv": gmv,
                    "refund": 0.0,
                    "actual": gmv,
                    "orders": int(day_df["订单号"].count()),
                    "buyers": None,
                }
            }
        return result

    df = pd.read_excel(path).fillna(0)
    order_df = df[df["类型"] == "订单入账"].copy()
    order_df["日期"] = pd.to_datetime(order_df["下单时间"], errors="coerce").dt.date
    order_df["收入(元)"] = order_df["收入(元)"].apply(parse_number)
    refund_df = df[df["类型"] == "退款"].copy()
    refund_df["日期"] = pd.to_datetime(refund_df["入账时间"], errors="coerce").dt.date
    refund_df["支出(元)"] = refund_df["支出(元)"].apply(parse_number)

    order_group = order_df.groupby("日期")
    refund_group = refund_df.groupby("日期")
    result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    for day in sorted(set(order_group.groups) | set(refund_group.groups)):
        day_orders = order_group.get_group(day) if day in order_group.groups else None
        day_refunds = refund_group.get_group(day) if day in refund_group.groups else None
        gmv = round2(day_orders["收入(元)"].sum()) if day_orders is not None else 0.0
        refund = round2(day_refunds["支出(元)"].sum()) if day_refunds is not None else 0.0
        result[day] = {
            ("有赞", "商城"): {
                "gmv": gmv,
                "refund": refund,
                "actual": round2(gmv - refund),
                "orders": int(len(day_orders)) if day_orders is not None else 0,
                "buyers": None,
            }
        }
    return result


def build_youzan_map_from_files(paths: List[Path]) -> Dict[date, Dict[Tuple[str, str], Dict]]:
    excel_paths = [path for path in paths if path.suffix.lower() in {".xlsx", ".xls"}]
    if excel_paths:
        return build_youzan_map(excel_paths[0])

    order_candidates = [path for path in paths if "order" in path.name.lower()]
    refund_candidates = [path for path in paths if "refund" in path.name.lower()]
    order_csv = max(order_candidates, key=lambda path: path.stat().st_mtime) if order_candidates else None
    refund_csv = max(refund_candidates, key=lambda path: path.stat().st_mtime) if refund_candidates else None
    if order_csv is None and refund_csv is None:
        raise RuntimeError(f"有赞缺少可解析文件：{[path.name for path in paths]}")

    result: Dict[date, Dict[Tuple[str, str], Dict]] = {}
    if order_csv is not None:
        order_df = pd.read_csv(order_csv).fillna(0)
        order_df["买家付款时间"] = pd.to_datetime(order_df["买家付款时间"], errors="coerce")
        order_df["订单实付金额"] = pd.to_numeric(order_df["订单实付金额"], errors="coerce").fillna(0)
        for day, day_df in order_df.groupby(order_df["买家付款时间"].dt.date):
            gmv = round2(day_df["订单实付金额"].sum())
            result[day] = {
                ("有赞", "商城"): {
                    "gmv": gmv,
                    "refund": 0.0,
                    "actual": gmv,
                    "orders": int(day_df["订单号"].count()),
                    "buyers": None,
                }
            }

    if refund_csv is not None:
        refund_df = pd.read_csv(refund_csv).fillna(0)
        refund_date_col = "退款完成时间" if "退款完成时间" in refund_df.columns else "付款时间"
        refund_df[refund_date_col] = pd.to_datetime(refund_df[refund_date_col], errors="coerce")
        refund_df["退款金额"] = pd.to_numeric(refund_df["退款金额"], errors="coerce").fillna(0)
        if "退款资金状态" in refund_df.columns:
            refund_df = refund_df[refund_df["退款资金状态"].astype(str).eq("退款成功")]
        refund_map = {
            day: round2(day_df["退款金额"].sum())
            for day, day_df in refund_df.groupby(refund_df[refund_date_col].dt.date)
        }
        for day in set(result) | set(refund_map):
            if day not in result:
                result[day] = {
                    ("有赞", "商城"): {
                        "gmv": 0.0,
                        "refund": 0.0,
                        "actual": 0.0,
                        "orders": 0,
                        "buyers": None,
                    }
                }
            row = result[day][("有赞", "商城")]
            row["refund"] = refund_map.get(day, 0.0)
            row["actual"] = round2(row["gmv"] - row["refund"])

    return result


def build_youzan_month_override(report_date: str, current_paths: List[Path]) -> Dict:
    """Build a month-to-date Youzan snapshot from archived sources, deduped by IDs."""
    target_day = datetime.strptime(report_date, "%Y-%m-%d").date()
    month_start = target_day.replace(day=1)
    YOUZAN_CUMULATIVE_DIR.mkdir(parents=True, exist_ok=True)

    for path in current_paths:
        if path.suffix.lower() != ".csv":
            continue
        destination = YOUZAN_CUMULATIVE_DIR / path.name
        if not destination.exists():
            shutil.copy2(path, destination)

    order_frames = []
    refund_frames = []
    cumulative_refund_files = []
    source_names = []
    for path in sorted(YOUZAN_CUMULATIVE_DIR.glob("*.csv")):
        lower_name = path.name.lower()
        if "order" in lower_name:
            frame = pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
            if {"订单号", "买家付款时间", "订单实付金额"} <= set(frame.columns):
                frame = frame[["订单号", "买家付款时间", "订单实付金额"]].copy()
                frame["买家付款时间"] = pd.to_datetime(frame["买家付款时间"], errors="coerce")
                frame["订单实付金额"] = pd.to_numeric(frame["订单实付金额"], errors="coerce").fillna(0)
                order_frames.append(frame)
                source_names.append(path.name)
        elif "refund" in lower_name:
            frame = pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
            required = {"售后编号", "退款完成时间", "退款金额"}
            if required <= set(frame.columns):
                if "退款资金状态" in frame.columns:
                    frame = frame[frame["退款资金状态"].astype(str).str.strip().eq("退款成功")]
                coverage_dates = pd.to_datetime(frame["退款完成时间"], errors="coerce").dt.date.dropna()
                if not coverage_dates.empty:
                    cumulative_refund_files.append(
                        {
                            "path": path.name,
                            "min_date": coverage_dates.min(),
                            "max_date": coverage_dates.max(),
                        }
                    )
                frame = frame[["售后编号", "退款完成时间", "退款金额"]].copy()
                frame["退款完成时间"] = pd.to_datetime(frame["退款完成时间"], errors="coerce")
                frame["退款金额"] = pd.to_numeric(frame["退款金额"], errors="coerce").fillna(0)
                refund_frames.append(frame)
                source_names.append(path.name)

    if not order_frames or not refund_frames:
        raise RuntimeError("有赞累计源不完整：必须同时存在订单和退款 CSV")

    orders = pd.concat(order_frames, ignore_index=True)
    orders = orders.dropna(subset=["买家付款时间"])
    orders["订单号"] = orders["订单号"].astype(str).str.strip()
    orders = orders.drop_duplicates(subset=["订单号"], keep="last")
    orders = orders[
        (orders["买家付款时间"].dt.date >= month_start)
        & (orders["买家付款时间"].dt.date <= target_day)
    ]

    refunds = pd.concat(refund_frames, ignore_index=True)
    refunds = refunds.dropna(subset=["退款完成时间"])
    refunds["售后编号"] = refunds["售后编号"].astype(str).str.strip()
    refunds = refunds.drop_duplicates(subset=["售后编号"], keep="last")
    refunds = refunds[
        (refunds["退款完成时间"].dt.date >= month_start)
        & (refunds["退款完成时间"].dt.date <= target_day)
    ]

    complete_refund_files = [
        item
        for item in cumulative_refund_files
        if item["min_date"] <= month_start and item["max_date"] >= target_day
    ]
    if not complete_refund_files:
        raise RuntimeError(
            "有赞退款累计源不完整：必须提供一份覆盖本月1日至报表日的退款完成明细，"
            "已停止发布，避免用单日退款文件生成错误月累计。"
        )

    gmv = round2(orders["订单实付金额"].sum())
    refund = round2(refunds["退款金额"].sum())
    daily_orders = orders[orders["买家付款时间"].dt.date == target_day]
    daily_refunds = refunds[refunds["退款完成时间"].dt.date == target_day]
    daily_gmv = round2(daily_orders["订单实付金额"].sum())
    daily_refund = round2(daily_refunds["退款金额"].sum())
    return {
        "gmv": gmv,
        "refund": refund,
        "actual": round2(gmv - refund),
        "orders": int(len(orders)),
        "refund_records": int(len(refunds)),
        "daily": {
            "gmv": daily_gmv,
            "refund": daily_refund,
            "actual": round2(daily_gmv - daily_refund),
            "orders": int(len(daily_orders)),
            "refund_records": int(len(daily_refunds)),
        },
        "as_of": report_date,
        "sources": sorted(set(source_names)),
        "complete_refund_source": complete_refund_files[-1]["path"],
    }


def pick_single_file(paths: List[Path], platform_name: str) -> Path:
    if not paths:
        raise RuntimeError(f"{platform_name} 缺少可解析文件")
    if len(paths) == 1:
        return paths[0]
    xlsx_paths = [path for path in paths if path.suffix.lower() in {".xlsx", ".xls"}]
    if len(xlsx_paths) == 1:
        return xlsx_paths[0]
    raise RuntimeError(f"{platform_name} 文件数量不明确，请手工确认：{[path.name for path in paths]}")


def pick_video_scene_files(paths: List[Path]) -> List[Path]:
    scene_paths = [
        path
        for path in paths
        if path.suffix.lower() in {".xlsx", ".xls"} and "场景构成" in path.name and not path.name.startswith("~$")
    ]
    return sorted(scene_paths, key=lambda path: path.name)


def pick_video_big_files(paths: List[Path]) -> Tuple[Path, Path]:
    live_path = None
    total_path = None
    for path in paths:
        name = path.name.lower()
        if path.suffix.lower() == ".csv":
            live_path = path
        elif "直播" in path.name or "历史" in path.name:
            live_path = path
        elif (
            "核心" in path.name
            or "总" in path.name
            or "概览" in path.name
            or "趋势" in path.name
        ):
            total_path = path
    if total_path is None or live_path is None:
        raise RuntimeError(f"视频号大号需要总盘文件和直播明细文件，当前文件：{[path.name for path in paths]}")
    return total_path, live_path


def pick_video_small_files(paths: List[Path]) -> Tuple[Optional[Path], Optional[Path], Optional[Path]]:
    total_path = None
    live_xlsx_path = None
    live_csv_path = None
    for path in paths:
        name = path.name.lower()
        suffix = path.suffix.lower()
        if suffix == ".csv":
            live_csv_path = path
        elif "直播" in path.name or "历史" in path.name:
            live_xlsx_path = path
        elif "趋势" in path.name or "统计" in path.name or "总" in path.name or "概览" in path.name:
            total_path = path
        elif total_path is None:
            total_path = path
    return total_path, live_xlsx_path, live_csv_path


def merge_day_payload(target: Dict[Tuple[str, str], Dict], source: Dict[Tuple[str, str], Dict]) -> None:
    for key, value in source.items():
        target[key] = value


def format_file_list(paths: List[Path]) -> str:
    return "、".join(path.name for path in paths)


def build_source_text(platform: str, account: str, file_paths: List[Path]) -> str:
    file_text = format_file_list(file_paths)

    if platform == "小红书" and account == "直播":
        return (
            f"小红书下载 Excel《{file_text}》；"
            "直播口径优先取不山直播支付金额+云南不山直播支付金额，"
            "直播退款优先取不山直播退款金额（支付时间）+云南不山直播退款金额（支付时间）；"
            "若源文件仍为旧版结构，则回退到直播支付金额、直播退款金额（支付时间）；"
            "直播实际销售额=直播GMV-直播退款。"
        )
    if platform == "小红书" and account == "商品卡":
        return (
            f"小红书下载 Excel《{file_text}》；"
            "商品卡口径=笔记支付金额+商卡支付金额，退款取笔记退款金额（支付时间）+商卡退款金额（支付时间）。"
        )
    if platform == "抖音" and account == "直播间":
        return (
            f"抖音下载 Excel《{file_text}》；"
            "直播间口径取成交概览中载体类型=直播、投放时段=不限。"
        )
    if platform == "抖音" and account == "商品卡":
        return (
            f"抖音下载 Excel《{file_text}》；"
            "商品卡口径取成交概览中载体类型=商品卡+其他、投放时段=不限。"
        )
    if platform == "有赞":
        return (
            f"有赞下载文件《{file_text}》；"
            "GMV 按买家付款时间汇总订单实付金额，退款按退款完成时间汇总退款成功金额。"
        )
    if platform == "视频号" and account == "大号":
        return (
            f"视频号大号下载文件《{file_text}》；"
            "口径=两个场景构成文件汇总；实际成交按成交金额-成交退款金额，直播间实际成交按直播间成交金额-直播间成交退款金额，其余场景计入商品卡。"
        )
    if platform == "视频号" and account == "小号":
        return (
            f"视频号小号下载文件《{file_text}》；"
            "口径=场景构成文件；实际成交按成交金额-成交退款金额，直播间实际成交按直播间成交金额-直播间成交退款金额，其余场景计入商品卡。"
        )
    return f"飞书云盘自动解析：{file_text}"


def get_parsed_platform_names(entries: List[Dict]) -> List[str]:
    entry_keys = {(str(entry.get("platform", "")).strip(), str(entry.get("account", "")).strip()) for entry in entries}
    parsed = []
    for platform_name in PLATFORM_ORDER:
        mappings = PLATFORM_ENTRY_MAP.get(platform_name, [])
        if any(mapping in entry_keys for mapping in mappings):
            parsed.append(platform_name)
    return parsed


def build_coverage_note(entries: List[Dict], downloaded_files: Dict[str, List[Path]]) -> str:
    parsed_platforms = get_parsed_platform_names(entries)
    parsed_text = "、".join(parsed_platforms) if parsed_platforms else "暂无"
    missing_file_platforms = [name for name in PLATFORM_ORDER if name in downloaded_files and not downloaded_files.get(name)]
    unresolved_platforms = [
        name for name in PLATFORM_ORDER if downloaded_files.get(name) and name not in parsed_platforms
    ]

    notes = [
        f"自动口径：当前已纳入 {parsed_text}。",
        "小红书、抖音按下载 Excel 口径。",
    ]

    if downloaded_files.get("有赞"):
        notes.append("有赞按订单 CSV 与退款 CSV 汇总。")
    if downloaded_files.get("视频号大号"):
        notes.append("视频号大号按两个场景构成文件汇总，实际成交按成交金额减成交退款金额，直播间实际成交单独拆分，其余场景并入商品卡。")
    if downloaded_files.get("视频号小号"):
        notes.append("视频号小号按场景构成文件统计，实际成交按成交金额减成交退款金额，直播间实际成交单独拆分，其余场景并入商品卡。")

    if missing_file_platforms:
        notes.append("以下平台当日未上传源文件，未纳入汇总：" + "、".join(missing_file_platforms) + "。")
    if unresolved_platforms:
        notes.append("以下平台已发现文件但未解析出结果，暂未纳入汇总：" + "、".join(unresolved_platforms) + "。")

    notes.append("当前仅复用既有平台解析逻辑生成单日输入 JSON，不重建整月工作簿。")
    return " ".join(notes)


def build_entries_from_downloads(report_date: str, downloaded_files: Dict[str, List[Path]]) -> List[Dict]:
    target_day = datetime.strptime(report_date, "%Y-%m-%d").date()
    day_map: Dict[Tuple[str, str], Dict] = {}

    if downloaded_files.get("小红书"):
        xhs_map = build_xhs_map(pick_single_file(downloaded_files["小红书"], "小红书"))
        merge_day_payload(day_map, xhs_map.get(target_day, {}))

    if downloaded_files.get("抖音"):
        douyin_map = build_douyin_map(pick_single_file(downloaded_files["抖音"], "抖音"))
        merge_day_payload(day_map, douyin_map.get(target_day, {}))

    if downloaded_files.get("视频号小号"):
        scene_paths = pick_video_scene_files(downloaded_files["视频号小号"])
        small_day = {}
        if scene_paths:
            small_day = build_video_scene_map(scene_paths[0], "小号").get(target_day, {})
        else:
            total_path, live_xlsx_path, live_csv_path = pick_video_small_files(downloaded_files["视频号小号"])
            if live_xlsx_path is not None:
                small_day = build_video_small_map(live_xlsx_path).get(target_day, {})
            elif live_csv_path is not None:
                small_day = build_video_small_csv_map(live_csv_path).get(target_day, {})
            elif total_path is not None:
                small_day = build_video_small_map(total_path).get(target_day, {})
        merge_day_payload(day_map, small_day)

    if downloaded_files.get("视频号大号"):
        try:
            scene_paths = pick_video_scene_files(downloaded_files["视频号大号"])
            big_day = {}
            if scene_paths:
                if len(scene_paths) < 2:
                    raise RuntimeError(f"视频号大号需要两个场景构成文件，当前文件：{[path.name for path in scene_paths]}")
                big_map = merge_metric_maps([build_video_scene_map(path, "大号") for path in scene_paths])
                big_day = big_map.get(target_day, {})
            else:
                total_path, live_path = pick_video_big_files(downloaded_files["视频号大号"])
                legacy_big_map = build_video_big_map(total_path, live_path)
                legacy_big_day = legacy_big_map.get(target_day, {})
                if legacy_big_day:
                    live = legacy_big_day.get(("视频号", "大号直播"), {})
                    card = legacy_big_day.get(("视频号", "大号商品卡"), {})
                    big_day = {
                        ("视频号", "大号"): {
                            "gmv": round2(parse_number(live.get("gmv")) + parse_number(card.get("gmv"))),
                            "refund": round2(parse_number(live.get("refund"))),
                            "actual": round2(parse_number(live.get("actual")) + parse_number(card.get("gmv"))),
                            "orders": live.get("orders"),
                            "buyers": live.get("buyers"),
                        }
                    }
            if big_day:
                merged = big_day.get(("视频号", "大号"), {})
                day_map[("视频号", "大号")] = {
                    "gmv": round2(parse_number(merged.get("gmv"))),
                    "refund": round2(parse_number(merged.get("refund"))),
                    "actual": round2(parse_number(merged.get("actual"))),
                    "orders": merged.get("orders"),
                    "buyers": merged.get("buyers"),
                    "live_actual": round2(parse_number(merged.get("live_actual"))),
                    "card_actual": round2(parse_number(merged.get("card_actual"))),
                }
        except RuntimeError as exc:
            print(f"跳过视频号大号：{exc}", flush=True)

    if downloaded_files.get("有赞"):
        youzan_map = build_youzan_map_from_files(downloaded_files["有赞"])
        merge_day_payload(day_map, youzan_map.get(target_day, {}))

    entries: List[Dict] = []
    for platform_name in PLATFORM_ORDER:
        mappings = PLATFORM_ENTRY_MAP.get(platform_name, [])
        platform_file_names = [path.name for path in downloaded_files.get(platform_name, [])]
        for platform, account in mappings:
            payload = day_map.get((platform, account))
            if not payload:
                continue
            entries.append(
                {
                    "platform": platform,
                    "account": account,
                    "gmv": round2(payload.get("gmv", 0)),
                    "actual": round2(payload.get("actual", 0)),
                    "refund": round2(payload.get("refund", 0)),
                    "orders": payload.get("orders"),
                    "buyers": payload.get("buyers"),
                    "status": "完整",
                    "source": build_source_text(platform, account, downloaded_files.get(platform_name, [])),
                    "live_actual": round2(payload.get("live_actual", 0)) if payload.get("live_actual") is not None else None,
                    "card_actual": round2(payload.get("card_actual", 0)) if payload.get("card_actual") is not None else None,
                }
            )
    if not entries:
        raise RuntimeError(f"{report_date} 未解析出任何平台明细，请检查飞书源文件格式")
    return entries


def build_input_payload(report_date: str, downloaded_files: Dict[str, List[Path]]) -> Dict:
    entries = build_entries_from_downloads(report_date, downloaded_files)
    available_platforms = [name for name, file_names in downloaded_files.items() if file_names]
    parsed_platforms = get_parsed_platform_names(entries)
    missing_file_platforms = [name for name in PLATFORM_ORDER if name in downloaded_files and not downloaded_files.get(name)]
    unresolved_platforms = [
        name for name in PLATFORM_ORDER if downloaded_files.get(name) and name not in parsed_platforms
    ]
    coverage_note = build_coverage_note(entries, downloaded_files)
    conclusion = (
        f"当前已解析到 {len(entries)} 条平台口径明细，覆盖 {len(available_platforms)} 个有文件的平台："
        + ("、".join(available_platforms) if available_platforms else "暂无")
        + "。"
    )
    if missing_file_platforms:
        conclusion += " 未上传源文件的平台：" + "、".join(missing_file_platforms) + "。"
    if unresolved_platforms:
        conclusion += " 有文件但未解析出结果的平台：" + "、".join(unresolved_platforms) + "。"

    payload = {
        "date": report_date,
        "coverage_note": coverage_note,
        "conclusion": conclusion,
        "entries": entries,
    }
    if downloaded_files.get("有赞"):
        payload["youzan_month_override"] = build_youzan_month_override(
            report_date, downloaded_files["有赞"]
        )
        daily_override = payload["youzan_month_override"]["daily"]
        for entry in payload["entries"]:
            if entry.get("platform") == "有赞":
                entry.update(daily_override)
                entry["source_note"] = f"有赞累计明细校准；退款来源覆盖本月1日至{report_date}"
                break
    return payload


def save_json(path: Path, payload: Dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def run_daily_report(input_path: Path, mode: str) -> int:
    if not RUN_SCRIPT.exists():
        raise FileNotFoundError(f"未找到脚本：{RUN_SCRIPT}")

    cmd = ["zsh", str(RUN_SCRIPT), "--input", str(input_path), "--mode", mode]
    print("调用日报脚本：")
    print(" ".join(cmd))
    completed = subprocess.run(cmd, cwd=str(SCRIPT_DIR))
    return completed.returncode


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从飞书云盘生成 daily_report_inputs/YYYY-MM-DD.json，并调用 run_daily_report.sh"
    )
    parser.add_argument("--date", help="指定日期，格式 YYYY-MM-DD；不传则自动取最新日期文件夹")
    parser.add_argument(
        "--mode",
        default="dry-run",
        help="传给 run_daily_report.sh 的 mode，默认 dry-run，避免直接同步或发送",
    )
    parser.add_argument(
        "--skip-run",
        action="store_true",
        help="只生成 input JSON，不调用 run_daily_report.sh",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        env, env_path = load_env()
        print(f"已读取配置：{env_path}")

        if args.date:
            datetime.strptime(args.date, "%Y-%m-%d")

        tenant_token = get_tenant_access_token(
            env["FEISHU_APP_ID"],
            env["FEISHU_APP_SECRET"],
        )
        print("tenant_access_token 获取成功")

        root_items = list_folder_items(env["FEISHU_FOLDER_TOKEN"], tenant_token)
        print(f"根目录读取成功，共 {len(root_items)} 个对象")

        date_folder = find_date_folder(root_items, args.date)
        if date_folder is None:
            if args.date:
                raise RuntimeError(f"未找到指定日期文件夹：{args.date}")
            raise RuntimeError("未找到日期格式文件夹")

        report_date = str(date_folder.get("name", "")).strip()
        print(f"目标日期：{report_date}")

        platform_items, _ = list_child_items(date_folder, tenant_token)
        platform_folders = sort_platform_folders(platform_items)
        print(f"平台文件夹数量：{len(platform_folders)}")
        for folder in platform_folders:
            print(f"- {folder.get('name', '<未命名>')}")

        platform_files = collect_platform_files(platform_folders, tenant_token)
        for platform_name in PLATFORM_ORDER:
            file_names = platform_files.get(platform_name, [])
            print(f"{platform_name} 文件数：{len(file_names)}")
            for file_name in file_names:
                print(f"  - {file_name}")

        download_root = Path(tempfile.gettempdir()) / "feishu_daily_report_inputs" / report_date
        downloaded_files = download_platform_files(platform_folders, tenant_token, download_root)
        payload = build_input_payload(report_date, downloaded_files)
        output_path = INPUT_DIR / f"{report_date}.json"
        save_json(output_path, payload)
        print(f"已生成输入 JSON：{output_path}")

        if args.skip_run:
            return 0

        return run_daily_report(output_path, args.mode)
    except ValueError as exc:
        print(f"参数错误：{exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"执行失败：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
