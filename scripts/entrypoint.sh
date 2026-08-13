#!/bin/sh
set -e

# 如果定义了 SERVER_URL，则在 HTML 目录中执行全局查找和替换
if [ -n "$SERVER_URL" ]; then
    echo "正在注入 SERVER_URL: $SERVER_URL"
    # 使用 envsubst 安全替换，仅处理文本文件避免损坏二进制
    export SERVER_URL
    find /usr/share/nginx/html -type f \( -name "*.html" -o -name "*.js" -o -name "*.css" -o -name "*.xml" \) -exec sh -c '
      for f do
        tmp=$(mktemp)
        envsubst '\''$SERVER_URL'\'' < "$f" > "$tmp" && mv "$tmp" "$f"
      done
    ' sh {} +
fi

exec "$@"
