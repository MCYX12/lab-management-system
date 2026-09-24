#!/usr/bin/env bash
set -e

cd "$(dirname "$0")"

if [ ! -x ".venv/bin/python" ]; then
  echo "未找到 .venv。请先运行 ./install_ubuntu.sh 完成首次安装。"
  exit 1
fi

exec .venv/bin/python main.py "$@"
