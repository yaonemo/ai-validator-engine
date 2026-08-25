#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用配置驱动校验引擎 - 一键运行入口
用法:
    python run.py 数据文件.xlsx [规则文件.json]
    python run.py                      # 用示例数据演示
示例:
    python run.py data/example_parking_data.xlsx rules/example_parking_rules.json
"""
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def main():
    if len(sys.argv) >= 2:
        data_file = sys.argv[1]
    else:
        data_file = os.path.join(BASE_DIR, "data", "example_parking_data.xlsx")
        print("未指定数据文件，使用示例数据演示\n")

    if len(sys.argv) >= 3:
        rules_file = sys.argv[2]
    else:
        rules_file = os.path.join(BASE_DIR, "rules", "example_parking_rules.json")

    if not os.path.exists(data_file):
        print(f"❌ 找不到数据文件: {data_file}")
        sys.exit(1)
    if not os.path.exists(rules_file):
        print(f"❌ 找不到规则文件: {rules_file}")
        sys.exit(1)

    engine = os.path.join(BASE_DIR, "engine", "validate_config_engine.py")
    os.system(f"{sys.executable} \"{engine}\" \"{data_file}\" \"{rules_file}\"")

if __name__ == "__main__":
    main()
