#!/bin/sh
set -e

# 如果定义了 SERVER_URL，则在 HTML 目录中执行全局查找和替换
if [ -n "$SERVER_URL" ]; then
    echo "正在注入 SERVER_URL: $SERVER_URL"
    # 使用 @ 作为定界符以处理包含 / 的 URL
    find /usr/share/nginx/html -type f -exec sed -i "s|__SERVER_URL__|$SERVER_URL|g" {} +
fi

exec "$@"
