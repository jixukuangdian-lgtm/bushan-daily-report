#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import mimetypes
import os
import uuid
from pathlib import Path
from typing import Any
from urllib import error, parse, request


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Upload organized Bushan folders or files to Feishu Drive.")
    parser.add_argument("--source", required=True, help="Local file or folder to upload")
    parser.add_argument("--parent-token", required=True, help="Feishu parent folder token")
    parser.add_argument("--app-id", default=os.environ.get("BUSHAN_FEISHU_APP_ID"))
    parser.add_argument("--app-secret", default=os.environ.get("BUSHAN_FEISHU_APP_SECRET"))
    parser.add_argument("--dry-run", action="store_true", help="Preview upload plan only")
    return parser.parse_args()


def request_json(url: str, *, method: str = "GET", payload: dict[str, Any] | None = None, headers: dict[str, str] | None = None) -> dict[str, Any]:
    body = None if payload is None else json.dumps(payload).encode("utf-8")
    req_headers = {"Content-Type": "application/json; charset=utf-8"}
    if headers:
        req_headers.update(headers)
    req = request.Request(url, data=body, headers=req_headers, method=method)
    try:
        with request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code}: {detail}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"Network error: {exc.reason}") from exc


def get_tenant_access_token(app_id: str, app_secret: str) -> str:
    payload = {"app_id": app_id, "app_secret": app_secret}
    data = request_json(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        method="POST",
        payload=payload,
    )
    if data.get("code") not in (None, 0):
        raise RuntimeError(f"Failed to get tenant access token: {data}")
    token = data.get("tenant_access_token") or data.get("data", {}).get("tenant_access_token")
    if not token:
        raise RuntimeError(f"Missing tenant access token in response: {data}")
    return token


def create_folder(name: str, parent_token: str, token: str) -> str:
    payload = {"name": name, "folder_token": parent_token}
    data = request_json(
        "https://open.feishu.cn/open-apis/drive/v1/files/create_folder",
        method="POST",
        payload=payload,
        headers={"Authorization": f"Bearer {token}"},
    )
    if data.get("code") not in (None, 0):
        # tolerate duplicate-name-like cases by surfacing a clear error for manual handling
        raise RuntimeError(f"Failed to create folder {name}: {data}")
    return str(data.get("data", {}).get("token") or data.get("data", {}).get("folder_token") or "")


def upload_file(path: Path, parent_token: str, token: str) -> dict[str, Any]:
    mime_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    boundary = f"----CodexBoundary{uuid.uuid4().hex}"
    body = []

    def add_field(name: str, value: str) -> None:
        body.extend(
            [
                f"--{boundary}\r\n".encode(),
                f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode(),
                value.encode(),
                b"\r\n",
            ]
        )

    add_field("file_name", path.name)
    add_field("parent_type", "explorer")
    add_field("parent_node", parent_token)
    add_field("size", str(path.stat().st_size))

    body.extend(
        [
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="file"; filename="{path.name}"\r\n'.encode(),
            f"Content-Type: {mime_type}\r\n\r\n".encode(),
            path.read_bytes(),
            b"\r\n",
            f"--{boundary}--\r\n".encode(),
        ]
    )

    req = request.Request(
        "https://open.feishu.cn/open-apis/drive/v1/files/upload_all",
        data=b"".join(body),
        method="POST",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        },
    )

    try:
        with request.urlopen(req, timeout=120) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Failed to upload {path.name}: HTTP {exc.code}: {detail}") from exc


def upload_tree(path: Path, parent_token: str, token: str, dry_run: bool, result: list[dict[str, str]]) -> None:
    if path.is_dir():
        if dry_run:
            result.append({"type": "folder", "name": path.name, "parent_token": parent_token, "action": "create"})
            child_parent = f"dry-run:{path.name}"
        else:
            child_parent = create_folder(path.name, parent_token, token)
            result.append({"type": "folder", "name": path.name, "parent_token": parent_token, "token": child_parent})

        for child in sorted(path.iterdir()):
            if child.name.startswith("."):
                continue
            upload_tree(child, child_parent, token, dry_run, result)
        return

    if dry_run:
        result.append({"type": "file", "name": path.name, "parent_token": parent_token, "action": "upload"})
        return

    upload_result = upload_file(path, parent_token, token)
    result.append(
        {
            "type": "file",
            "name": path.name,
            "parent_token": parent_token,
            "token": str(upload_result.get("data", {}).get("file_token") or upload_result.get("data", {}).get("token") or ""),
        }
    )


def main() -> int:
    args = parse_args()
    source = Path(args.source).expanduser().resolve()
    if not source.exists():
        raise SystemExit(f"Source does not exist: {source}")
    if not args.app_id or not args.app_secret:
        raise SystemExit("Missing Feishu app credentials. Set BUSHAN_FEISHU_APP_ID and BUSHAN_FEISHU_APP_SECRET.")

    token = "" if args.dry_run else get_tenant_access_token(args.app_id, args.app_secret)
    result: list[dict[str, str]] = []
    upload_tree(source, args.parent_token, token, args.dry_run, result)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
