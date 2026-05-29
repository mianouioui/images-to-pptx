#!/bin/zsh
# 图片转PPTX v3.0.2 - macOS 入口脚本
chmod +x "$0" >/dev/null 2>&1 || true
SCRIPT_DIR="${0:A:h}"
exec "$SCRIPT_DIR/图片转PPTX.command" "$@"
