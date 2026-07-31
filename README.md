# Bushan Daily Report Skill

This repository packages Bushan's daily data workflow into a reusable Skill bundle. It is designed for migration, reuse, and long-term maintenance rather than one-off script execution.

## What this skill covers

The skill models the full reporting chain:

1. Download raw platform files
2. Organize them into local date/platform folders
3. Upload organized files to Feishu Drive
4. Trigger Feishu button or local runner
5. Scan Feishu files and build normalized daily input JSON
6. Update workbook and monthly progress
7. Sync Feishu Base
8. Send the daily Feishu report card

## Three supported entrypoints

### 1. Local full workflow

Use this when raw files were just downloaded locally and still need to be organized and uploaded.

Main files:

- `scripts/core/organize_downloads.py`
- `scripts/core/upload_to_feishu_drive.py`
- `scripts/entrypoints/run_local_ingest.sh`
- `scripts/entrypoints/run_month_resume.sh`

### 2. Feishu button workflow

Use this when files are already uploaded and the operator should only need one click in Feishu.

Main files:

- `scripts/automation/send_report_button.sh`
- `scripts/automation/feishu_report_button.py`
- `assets/cards/report_button_card.json`

### 3. Historical rerun workflow

Use this when one day must be recalculated after corrected source files, rule changes, or failed month-progress updates.

Main files:

- `scripts/entrypoints/run_month_resume.sh`
- `scripts/core/build_input_json.py`
- `scripts/core/daily_report_pipeline.py`

## Repository layout

- `SKILL.md`: the reusable agent instructions.
- `scripts/core/`: platform parsing, ingest, workbook, Base-sync, and send logic.
- `scripts/entrypoints/`: operator-facing shell entrypoints.
- `scripts/automation/`: Feishu trigger listener and sender.
- `references/`: business rules, runbooks, and troubleshooting.
- `config/`: sanitized template configuration.
- `examples/`: sample normalized input and sample card payload.
- `assets/`: static artifacts such as templates and trigger cards.

## Quick start

1. Run the setup wizard:
   - `./scripts/entrypoints/setup_feishu.sh`
2. Add a sanitized workbook template to `assets/templates/`.
3. For local-ingest mode:
   - run `scripts/entrypoints/run_local_ingest.sh`
4. For rerun mode:
   - run `scripts/entrypoints/run_month_resume.sh --date YYYY-MM-DD`
5. For Feishu button mode:
   - send the trigger card with `scripts/automation/send_report_button.sh`

If manual setup is preferred, use:

- `config/.env.example`
- `config/daily_report_pipeline_config.template.json`
- `config/folder_mapping.example.json`
- `config/month_targets.example.json`

## Important safety rules

- Never publish real webhook URLs, base tokens, folder tokens, or operator open IDs.
- Never generate month totals from incomplete Youzan cumulative refund sources.
- Rewrite existing workbook rows when rerunning historical dates.
- Rerun affected dates one by one after workbook corruption recovery.

## Packaging

This skill already validates and packages cleanly with Python 3.11.

Example:

```bash
python3.11 /Users/jixukuangdian/.agents/skills/skill-creator/scripts/package_skill.py /path/to/bushan-data-ops ./dist
```
