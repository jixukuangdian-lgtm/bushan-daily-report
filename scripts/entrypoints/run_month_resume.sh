#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TARGET_DATE=""
SEND_ENABLED=1
RESTART_ENABLED=0

CONFIG_PATH="$SCRIPT_DIR/daily_report_pipeline_config_2026_07.json"
WORKBOOK_PATH="$SCRIPT_DIR/不山电商日报_2026年07.xlsx"
OUTPUT_DIR="$SCRIPT_DIR/daily_report_outputs"
INPUT_DIR="$SCRIPT_DIR/daily_report_inputs"
STATE_DIR="$SCRIPT_DIR/.run_state"

usage() {
  cat <<'EOF'
用法：
  ./run_2026_07_resume_all.sh --date 2026-07-12
  ./run_2026_07_resume_all.sh --date 2026-07-12 --no-send
  ./run_2026_07_resume_all.sh --date 2026-07-12 --restart

说明：
  这是 2026 年 7 月一键可续跑脚本，默认按当前已确认口径执行：
  1. 准备 7月 workbook
  2. 生成 daily_report_inputs/YYYY-MM-DD.json
  3. 更新 7月 workbook 并同步飞书 Base
  4. 发送飞书日报卡片

  中途如果被中断，再次执行同一个日期会自动跳过已完成步骤，继续往后跑。

可选参数：
  --date YYYY-MM-DD  指定 7 月内要处理的日期
  --no-send          只算数据、更新 workbook、同步 Base，不发卡片
  --restart          清空该日期的续跑状态，从头执行
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
    --restart)
      RESTART_ENABLED=1
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
  echo "必须传 --date，例如：./run_2026_07_resume_all.sh --date 2026-07-12" >&2
  exit 1
fi

if [[ "$TARGET_DATE" != 2026-07-* ]]; then
  echo "这个脚本只处理 2026-07 的日期，当前收到：$TARGET_DATE" >&2
  exit 1
fi

mkdir -p "$STATE_DIR"
STATE_FILE="$STATE_DIR/$TARGET_DATE.state"

if [[ "$RESTART_ENABLED" == "1" && -f "$STATE_FILE" ]]; then
  rm -f "$STATE_FILE"
fi

CURRENT_STEP=0
if [[ -f "$STATE_FILE" ]]; then
  CURRENT_STEP="$(cat "$STATE_FILE" 2>/dev/null || echo 0)"
fi

save_step() {
  printf '%s\n' "$1" > "$STATE_FILE"
}

run_step() {
  local step_num="$1"
  local title="$2"
  shift 2

  if (( CURRENT_STEP >= step_num )); then
    echo "跳过步骤 ${step_num}：${title}（已完成）"
    return 0
  fi

  echo
  echo "步骤 ${step_num}/4：${title}"
  "$@"
  save_step "$step_num"
  CURRENT_STEP="$step_num"
}

echo "执行日期：$TARGET_DATE"
echo "工作簿：$WORKBOOK_PATH"
echo "输出目录：$OUTPUT_DIR"
echo "续跑状态：步骤 $CURRENT_STEP"

run_step 1 "准备 7月 workbook" \
  python3 "$SCRIPT_DIR/prepare_2026_07_workbook.py"

run_step 2 "生成输入 JSON" \
  python3 "$SCRIPT_DIR/feishu_to_input_json.py" --date "$TARGET_DATE" --skip-run

run_step 3 "更新 7月 workbook 并同步飞书 Base" \
  python3 "$SCRIPT_DIR/daily_report_pipeline.py" \
    --input "$INPUT_DIR/$TARGET_DATE.json" \
    --config "$CONFIG_PATH" \
    --workbook "$WORKBOOK_PATH" \
    --output-dir "$OUTPUT_DIR" \
    --sync-base

if [[ "$SEND_ENABLED" == "1" ]]; then
  run_step 4 "发送飞书日报卡片" \
    python3 "$SCRIPT_DIR/send_card.py" --date "$TARGET_DATE"
else
  echo
  echo "步骤 4/4：已跳过飞书日报卡片发送"
  save_step 4
fi

echo
echo "执行完成：$TARGET_DATE"
echo "如需重头再跑：./run_2026_07_resume_all.sh --date $TARGET_DATE --restart"
