#!/bin/bash
# 通用校验引擎 - 一键运行（Mac/Linux）
# 用法: ./run.sh [数据文件.xlsx] [规则.json]
cd "$(dirname "$0")"
if [ $# -ge 1 ]; then
    python3 run.py "$1" "${2:-rules/example_parking_rules.json}"
else
    python3 run.py
fi
