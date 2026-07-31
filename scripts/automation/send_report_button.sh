#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/../.." && pwd)"
USER_OPEN_ID="${BUSHAN_REPORT_OPERATOR_OPEN_ID:-YOUR_OPERATOR_OPEN_ID}"
LARK_CLI_PATH="${BUSHAN_LARK_CLI_PATH:-$HOME/.npm-global/bin/lark-cli}"
CARD_PATH="$PROJECT_DIR/assets/cards/report_button_card.json"

CARD_JSON="$(python3 -c 'import json,sys; print(json.dumps(json.load(open(sys.argv[1], encoding="utf-8")), ensure_ascii=False, separators=(",",":")))' "$CARD_PATH")"

exec "$LARK_CLI_PATH" im +messages-send \
  --user-id "$USER_OPEN_ID" \
  --msg-type interactive \
  --content "$CARD_JSON" \
  --idempotency-key "bushan-report-button-v2" \
  --as bot
