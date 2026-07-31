# Local Ingest Rules

## Purpose

Normalize raw downloaded files into a stable local structure before upload or parsing.

## Standard folder structure

```text
YYYY-MM-DD/
├── 小红书/
├── 抖音/
├── 有赞/
├── 视频号大号/
└── 视频号小号/
```

## Classification rules

### Xiaohongshu

Typical file names:

- `商家成交数据概览-all(...)`
- `商家经营数据总览-all(...)`

### Douyin

Typical file names:

- `抖音电商罗盘-成交分析-...`

### Youzan

Typical file names:

- `Order_youzan_...csv`
- `Refund_youzan_...csv`

Use two layers of source management:

- same-day files for one-day totals
- cumulative refund sources for month-safe refund accumulation

### WeChat Video big account

Preferred file set:

- two `场景构成` files

### WeChat Video small account

Preferred file set:

- one `场景构成` file

## Upload handoff

After local organization, upload the date folder or its platform subfolders to the matching Feishu drive date folder.

Keep file names unchanged whenever possible so the parsing logic can rely on the naming patterns.

## Included skill scripts

- `scripts/core/organize_downloads.py`
- `scripts/core/upload_to_feishu_drive.py`
- `scripts/entrypoints/run_local_ingest.sh`

These scripts are the reusable first-pass implementation for the ingest stage. Adapt the naming tokens and folder-token mapping through config, not by hardcoding new business rules into the runner.
