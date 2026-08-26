#!/bin/zsh

set -u

SCRIPT_DIR="${0:A:h}"
cd "$SCRIPT_DIR" || exit 1

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

if [[ -z "$PYTHON_CMD" ]]; then
  printf '[失败] 未找到 Python 3.10 或更高版本。\n'
  printf '按回车键关闭窗口...'
  read -r _
  exit 1
fi

"$PYTHON_CMD" "$SCRIPT_DIR/view_erp.py"
VIEW_EXIT_CODE=$?

if [[ "$VIEW_EXIT_CODE" -ne 0 ]]; then
  printf '\n[失败] ERP专用Chrome打开失败。\n'
  printf '按回车键关闭窗口...'
  read -r _
fi

exit "$VIEW_EXIT_CODE"
