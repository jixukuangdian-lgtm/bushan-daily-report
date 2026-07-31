#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TARGET_DATE=""
SEND_ENABLED=1

CONFIG_PATH="$SCRIPT_DIR/daily_report_pipeline_config_2026_07.json"
WORKBOOK_PATH="$SCRIPT_DIR/不山电商日报_2026年07.xlsx"
OUTPUT_DIR="$SCRIPT_DIR/daily_report_outputs"
INPUT_DIR="$SCRIPT_DIR/daily_report_inputs"

usage() {
  cat <<'EOF'
用法：
  ./run_2026_07_all.sh --date 2026-07-01
  ./run_2026_07_all.sh --date 2026-07-01 --no-send

说明：
  这是 2026 年 7 月通用脚本，会：
  1. 准备 7月 workbook
  2. 生成 daily_report_inputs/YYYY-MM-DD.json
  3. 更新 7月 workbook 并同步 7月飞书 Base
  4. 发送飞书日报卡片

可选参数：
  --date YYYY-MM-DD  指定 7 月内要处理的日期
  --no-send          只算数据、更新 workbook、同步 Base，不发卡片
  --help             显示帮助
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --date)
      TARGET_DATE="${2:-}"
      shift 2
      ;;
    --no-send)
      SEND_ENABLED=0
      shift
      ;;
    --help|-h)
      usage
      exit 0
      ;;
    *)
      echo "不支持的参数：$1" >&2
      echo >&2
      usage >&2
      exit 1
      ;;
  esac
done

if [[ -z "$TARGET_DATE" ]]; then
  echo "必须传 --date，例如：./run_2026_07_all.sh --date 2026-07-01" >&2
  exit 1
fi

if [[ "$TARGET_DATE" != 2026-07-* ]]; then
  echo "这个脚本只处理 2026-07 的日期，当前收到：$TARGET_DATE" >&2
  exit 1
fi

echo "执行日期：$TARGET_DATE"
echo "流程 0/3：准备 7月 workbook"
python3 "$SCRIPT_DIR/prepare_2026_07_workbook.py"

echo
echo "流程 1/3：生成输入 JSON"
python3 "$SCRIPT_DIR/feishu_to_input_json.py" --date "$TARGET_DATE" --skip-run

echo
echo "流程 2/3：更新 7月 workbook 并同步飞书 Base"
python3 "$SCRIPT_DIR/daily_report_pipeline.py" \
  --input "$INPUT_DIR/$TARGET_DATE.json" \
  --config "$CONFIG_PATH" \
  --workbook "$WORKBOOK_PATH" \
  --output-dir "$OUTPUT_DIR" \
  --sync-base

if [[ "$SEND_ENABLED" == "1" ]]; then
  echo
  echo "流程 3/3：发送飞书机器人"
  python3 "$SCRIPT_DIR/send_card.py" --date "$TARGET_DATE"
else
  echo
  echo "流程 3/3：已跳过飞书机器人发送"
fi

echo
echo "执行完成：$TARGET_DATE"
