#!/bin/zsh

cd "/Users/blue/Documents/跨境供应链跟单" || exit 1

PYTHON="/Users/blue/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
SERVER="tools/scm_dashboard_server.py"
URL="http://127.0.0.1:8765/dashboard"

if pgrep -f "$SERVER" >/dev/null 2>&1; then
  open "$URL"
  exit 0
fi

if curl -fsS "$URL" >/dev/null 2>&1; then
  open "$URL"
  exit 0
fi

"$PYTHON" "$SERVER" &
SERVER_PID=$!

for i in {1..20}; do
  if curl -fsS "$URL" >/dev/null 2>&1; then
    open "$URL"
    echo "SCM 看板已启动：$URL"
    echo "这个窗口可以最小化；关闭窗口会停止看板服务。"
    wait "$SERVER_PID"
    exit 0
  fi
  sleep 0.5
done

echo "看板启动失败，请把这个窗口内容发给 Codex。"
wait "$SERVER_PID"
