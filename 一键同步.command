#!/bin/zsh

set -u

SCRIPT_DIR="${0:A:h}"
cd "$SCRIPT_DIR" || exit 1

export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8

wait_to_close() {
  printf '\n按回车键关闭窗口...'
  read -r _
}

fail() {
  printf '\n[失败] %s\n' "$1"
  wait_to_close
  exit "${2:-1}"
}

printf '%s\n' '========================================'
printf '%s\n' 'ERP 月度实发数量与退货率同步'
printf '配置文件：%s\n' "$SCRIPT_DIR/config.json"
printf '%s\n' '========================================'

[[ -f "$SCRIPT_DIR/erp_excel_sync.py" ]] || fail '找不到 erp_excel_sync.py'
[[ -f "$SCRIPT_DIR/config.json" ]] || fail '找不到 config.json'

PYTHON_CMD=''
for candidate in \
  "$HOME/.local/bin/python3.12" \
  /opt/homebrew/bin/python3.12 /opt/homebrew/bin/python3.11 /opt/homebrew/bin/python3.10 \
  /usr/local/bin/python3.12 /usr/local/bin/python3.11 /usr/local/bin/python3.10 \
  python3.12 python3.11 python3.10 python3; do
  if command -v "$candidate" >/dev/null 2>&1 && \
    "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)' >/dev/null 2>&1; then
    PYTHON_CMD="$candidate"
    break
  fi
done

[[ -n "$PYTHON_CMD" ]] || fail '未找到 Python 3.10 或更高版本，请先安装 Python。'

printf '使用 Python：%s\n\n' "$($PYTHON_CMD --version 2>&1)"
"$PYTHON_CMD" "$SCRIPT_DIR/erp_excel_sync.py" --config "$SCRIPT_DIR/config.json"
SYNC_EXIT_CODE=$?

if [[ "$SYNC_EXIT_CODE" -eq 0 ]]; then
  printf '\n[成功] 月度同步已完成，原工作簿已保留为 .bak 备份。\n'
elif [[ "$SYNC_EXIT_CODE" -eq 2 ]]; then
  printf '\n[未替换] 关键货号校验未通过，请查看 reports 目录。\n'
else
  printf '\n[失败] 同步程序退出，错误码：%s\n' "$SYNC_EXIT_CODE"
fi

wait_to_close
exit "$SYNC_EXIT_CODE"
