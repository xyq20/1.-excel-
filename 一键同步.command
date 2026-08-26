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

KEYCHAIN_SERVICE='ERP Excel Sync'
KEYCHAIN_ACCOUNT="${USER:-erp-sync-user}"
COOKIE_SOURCE='environment'

if [[ -z "${ERP_COOKIE:-}" ]]; then
  ERP_COOKIE="$(security find-generic-password \
    -s "$KEYCHAIN_SERVICE" \
    -a "$KEYCHAIN_ACCOUNT" \
    -w 2>/dev/null || true)"
  [[ -n "$ERP_COOKIE" ]] && COOKIE_SOURCE='keychain'
fi

if [[ -z "${ERP_COOKIE:-}" ]]; then
  printf '首次运行，请粘贴 ERP Cookie（输入不会显示）：'
  IFS= read -r -s ERP_COOKIE
  printf '\n'
  [[ -n "$ERP_COOKIE" ]] || fail 'ERP Cookie 不能为空。'
  security add-generic-password \
    -U \
    -s "$KEYCHAIN_SERVICE" \
    -a "$KEYCHAIN_ACCOUNT" \
    -w "$ERP_COOKIE" >/dev/null 2>&1 || \
    fail '无法将 ERP Cookie 保存到 macOS 钥匙串。'
  printf '[已保存] ERP Cookie 已安全存入 macOS 钥匙串。\n'
elif [[ "$COOKIE_SOURCE" == 'keychain' ]]; then
  printf '[已读取] 已从 macOS 钥匙串读取 ERP Cookie。\n'
else
  printf '[已读取] 已从环境变量读取 ERP Cookie。\n'
fi
export ERP_COOKIE

printf '使用 Python：%s\n\n' "$($PYTHON_CMD --version 2>&1)"
"$PYTHON_CMD" "$SCRIPT_DIR/erp_excel_sync.py" --config "$SCRIPT_DIR/config.json"
SYNC_EXIT_CODE=$?

if [[ "$SYNC_EXIT_CODE" -eq 0 ]]; then
  printf '\n[成功] 月度同步已完成，原工作簿已保留为 .bak 备份。\n'
elif [[ "$SYNC_EXIT_CODE" -eq 2 ]]; then
  printf '\n[未替换] 关键货号校验未通过，请查看 reports 目录。\n'
else
  printf '\n[失败] 同步程序退出，错误码：%s\n' "$SYNC_EXIT_CODE"
  printf '如果报错提示 Cookie 或登录已失效，请在“钥匙串访问”中删除“ERP Excel Sync”后重新运行。\n'
fi

wait_to_close
exit "$SYNC_EXIT_CODE"
