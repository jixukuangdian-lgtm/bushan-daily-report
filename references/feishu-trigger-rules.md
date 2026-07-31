# Feishu Trigger Rules

## Purpose

Allow the operator to trigger daily reporting by clicking a Feishu interactive card instead of opening a terminal.

## Core components

- `scripts/automation/send_report_button.sh`
- `scripts/automation/feishu_report_button.py`
- `assets/cards/report_button_card.json`

## Expected behavior

1. Send one trigger card to the operator.
2. Operator clicks the retry/run button.
3. Listener validates operator open ID.
4. Listener selects the target date, usually yesterday.
5. Listener runs the resumable day workflow.
6. Card updates to one of:
   - running
   - success
   - failed but resumable

## Important constraints

- Do not hardcode a real operator open ID when publishing to GitHub.
- Read the operator ID, CLI path, and runner path from environment variables or a local setup file.
- Keep logs and state files so that a failed click can be resumed safely.

## Safe failure behavior

If parsing or source validation fails:

- keep the state file
- keep the logs
- update the card with a resumable failure message
- do not claim the report was sent
