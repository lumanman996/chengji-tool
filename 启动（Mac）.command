#!/bin/bash
# 双击运行。第一次会自动准备环境（需要联网），以后直接启动。
cd "$(dirname "$0")" || exit 1
if [ ! -x .venv/bin/python ]; then
  echo "第一次使用，正在准备环境（需要联网，大约一两分钟）…"
  if ! command -v python3 >/dev/null 2>&1; then
    echo "没有找到 Python 3。请先到 https://www.python.org/downloads/ 下载安装，再双击本文件。"
    read -r -p "按回车键关闭窗口…"; exit 1
  fi
  python3 -m venv .venv && .venv/bin/pip install -q --upgrade pip && .venv/bin/pip install -q -r requirements.txt
  if [ $? -ne 0 ]; then
    rm -rf .venv
    echo "环境没有准备好（多半是网络问题），请检查网络后再双击一次。"
    read -r -p "按回车键关闭窗口…"; exit 1
  fi
fi
exec .venv/bin/python -m chengji "$@"
