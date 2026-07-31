# Setup Guide

## Fastest first-time setup

Prefer the setup wizard instead of editing JSON files by hand.

Run:

```bash
./scripts/entrypoints/guide_configure_feishu_bot.sh
```

The wizard will:

1. Ask for Feishu app id and app secret
2. Ask for root folder token and platform folder tokens
3. Ask for webhook, Base token, and operator open id
4. Ask for workbook template path
5. Ask for month goal and month targets
6. Generate:
   - `.env`
   - `daily_report_pipeline_config.json`
   - `folder_mapping.json`
   - `month_targets.json`
7. Try to validate Feishu app credentials unless skipped
8. Check whether `lark-cli` and workbook template path already exist locally

## Recommended preparation

Prepare these values before running the wizard:

- Feishu app id
- Feishu app secret
- Feishu bot webhook
- Feishu Base token
- Feishu root folder token
- One folder token for each platform
- Operator open id
- Workbook template path

## Friendly entrypoint name

When the operator only needs the plain-language instruction, use this phrasing:

- "指引我配置飞书机器人"

The matching shell entrypoint is:

```bash
./scripts/entrypoints/guide_configure_feishu_bot.sh
```

## If validation fails

- Keep the generated files
- Correct the values
- Run the wizard again
- Or rerun with `--skip-validate` if network is temporarily unavailable
