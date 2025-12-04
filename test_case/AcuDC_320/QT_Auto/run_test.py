#!/usr/bin/env python3
"""
测试运行脚本（完全依赖conftest.py配置）
"""

import pytest
import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    # 直接运行pytest，所有配置都在conftest.py中
    sys.exit(pytest.main([
        "test_transaction/test_transaction.py",
        "test_echilog/test_echilog.py",
        "-v",
        "--tb=short",
        "-s"
    ]))