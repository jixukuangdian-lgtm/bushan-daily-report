# Workflow Overview

## Goal

Package Bushan's daily data workflow into a reusable operating system that can be copied, migrated, and maintained.

## Three entrypoints

### 1. Local full workflow

Use this when raw files were just downloaded locally and still need to be organized.

Order:

1. Classify raw source files into a `YYYY-MM-DD/平台/` folder structure.
2. Upload the classified files to the matching Feishu drive folder.
3. Build the normalized daily input JSON.
4. Update the workbook.
5. Sync Feishu Base or sheets.
6. Send or update the Feishu report card.

### 2. Feishu button workflow

Use this when files are already in Feishu drive and the operator should only need one click.

Order:

1. Send the trigger card.
2. Operator clicks the button.
3. Listener validates operator identity.
4. Listener chooses the target date, usually yesterday.
5. Listener runs the resumable day workflow.
6. Status card updates to success or failure.

### 3. Rerun workflow

Use this when a historical day must be recalculated.

Typical triggers:

- Corrected source files
- New parsing rule
- Product-card rule change
- Wrong month progress
- Wrong Youzan refund accumulation

Order:

1. Confirm the corrected source files are in place.
2. Rebuild the one-day input JSON.
3. Rewrite the workbook rows for that date.
4. Recompute month progress.
5. Resync Feishu Base.
6. Resend card only if needed.

## Recommended default runner

Prefer the resumable runner for daily use because interruptions are common.

- `scripts/entrypoints/run_month_resume.sh`

Use the non-resumable runner only when the environment is stable and the user explicitly wants one-shot execution.
