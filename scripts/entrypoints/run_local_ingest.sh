#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILL_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

SOURCE_DIR="${1:-}"
TARGET_ROOT="${2:-$SKILL_ROOT/data}"
TARGET_DATE="${3:-}"
PARENT_TOKEN="${BUSHAN_FEISHU_FOLDER_TOKEN:-}"

if [[ -z "$SOURCE_DIR" ]]; then
  echo "Usage: ./run_local_ingest.sh <source-dir> [target-root] [YYYY-MM-DD]" >&2
  exit 1
fi

ARGS=(
  python3
  "$SKILL_ROOT/scripts/core/organize_downloads.py"
  --source-dir "$SOURCE_DIR"
  --target-root "$TARGET_ROOT"
)

if [[ -n "$TARGET_DATE" ]]; then
  ARGS+=(--date "$TARGET_DATE")
fi

"${ARGS[@]}"

if [[ -n "$PARENT_TOKEN" ]]; then
  if [[ -n "$TARGET_DATE" ]]; then
    python3 "$SKILL_ROOT/scripts/core/upload_to_feishu_drive.py" \
      --source "$TARGET_ROOT/$TARGET_DATE" \
      --parent-token "$PARENT_TOKEN"
  else
    echo "Skipping Feishu upload because no date override was provided." >&2
  fi
else
  echo "Skipping Feishu upload because BUSHAN_FEISHU_FOLDER_TOKEN is not set." >&2
fi
