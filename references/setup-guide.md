# Setup Guide

## Fastest first-time setup

Prefer the setup wizard instead of editing JSON files by hand.

Run:

```bash
./scripts/entrypoints/setup_feishu.sh
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

## If validation fails

- Keep the generated files
- Correct the values
- Run the wizard again
- Or rerun with `--skip-validate` if network is temporarily unavailable
