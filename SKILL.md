---
name: bushan-data-ops
description: This skill should be used when building, migrating, or running Bushan's end-to-end daily data workflow, including local file organization, Feishu drive scanning, daily report generation, Base/table updates, and Feishu bot delivery.
---

# Bushan Data Ops

Build and maintain a reusable Bushan daily data-operations workflow. Treat the skill as one end-to-end operating package with three entrypoints:

1. Local full workflow: downloaded source files -> local classification -> upload to Feishu drive -> report generation.
2. Feishu button workflow: click one interactive card -> scan new files in Feishu drive -> continue report generation automatically.
3. Rerun workflow: reprocess one historical date after source corrections, rule changes, or table-sync failures.

Keep business-rule documents in `references/` as the source of truth. Keep executable logic in `scripts/`. Keep templates and sample artifacts in `assets/`, `config/`, and `examples/`.

## When to Use This Skill

Use this skill when any of the following is needed:

- Package the current daily-report project into a reusable GitHub-ready skill.
- Set up the workflow in a new machine, new workspace, or new month.
- Explain or modify the file-ingest rules for Xiaohongshu, Douyin, Youzan, or WeChat Video accounts.
- Maintain the Feishu button trigger that lets one click resume report generation.
- Diagnose rerun problems such as stale workbook values, incomplete Youzan refund sources, or Feishu send failures.

## How to Use

Follow this order:

1. Read `references/workflow-overview.md` to determine which of the three entrypoints applies.
2. Read `references/business-rules.md` before changing any parsing or aggregation logic.
3. Read `references/local-ingest-rules.md` when the task starts from downloaded raw files.
4. Read `references/feishu-trigger-rules.md` when the task starts from the Feishu robot/button/card.
5. Read `references/troubleshooting.md` when the task involves reruns, stale month totals, workbook corruption, or Youzan cumulative refund checks.
6. Use the scripts in `scripts/core/` for parsing, workbook updates, and send logic.
7. Use the scripts in `scripts/entrypoints/` for month/day execution.
8. Use the scripts in `scripts/automation/` for the Feishu button listener and trigger card.
9. Before publishing to GitHub, replace all real tokens, webhook URLs, open IDs, and absolute local paths with template values from `config/` and `.env.example`.

## References

- `references/workflow-overview.md`: full architecture, three entrypoints, and recommended execution order.
- `references/business-rules.md`: reporting rules for each platform, monthly progress logic, product-card definitions, and Youzan refund guardrails.
- `references/local-ingest-rules.md`: how raw downloads are classified into local date/platform folders and uploaded to Feishu drive.
- `references/feishu-trigger-rules.md`: how the Feishu robot card, listener, and click-to-run flow are wired.
- `references/troubleshooting.md`: common failure patterns and safe recovery methods.
- `references/github-publish.md`: how to sanitize, version, zip, and publish the skill package.

## Scripts

- `scripts/core/build_input_json.py`: build one-day normalized input JSON from Feishu or local source files.
- `scripts/core/organize_downloads.py`: classify freshly downloaded raw files into `YYYY-MM-DD/平台/` local folders.
- `scripts/core/upload_to_feishu_drive.py`: upload organized local folders or files into Feishu Drive.
- `scripts/core/daily_report_pipeline.py`: update workbook, compute month progress, sync Base, and build/send report artifacts.
- `scripts/core/prepare_month_workbook.py`: create or refresh monthly workbook structure.
- `scripts/core/refresh_month_progress_base.py`: refresh month progress in Base from workbook data.
- `scripts/core/send_card.py`: send a generated interactive report card with retry logic.
- `scripts/entrypoints/run_month_all.sh`: one-shot day runner.
- `scripts/entrypoints/run_month_resume.sh`: resumable day runner.
- `scripts/entrypoints/run_local_ingest.sh`: local file organization plus optional Feishu upload handoff.
- `scripts/automation/feishu_report_button.py`: listener for the Feishu interactive card button.
- `scripts/automation/send_report_button.sh`: send the trigger card to the operator.

## Assets

- `assets/cards/report_button_card.json`: base card used for the Feishu trigger flow.
- `assets/templates/`: place sanitized workbook templates or month setup templates here.
- `examples/sample_daily_input.json`: sample normalized one-day input payload.
- `examples/sample_daily_card.json`: sample generated Feishu daily card payload.
