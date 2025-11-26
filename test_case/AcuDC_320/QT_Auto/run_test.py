#!/usr/bin/env python3
"""
测试运行脚本
"""

import pytest
import os
import sys
from datetime import datetime

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def run_tests():
    """运行测试并生成报告"""

    # 创建报告目录
    report_dir = "reports"
    if not os.path.exists(report_dir):
        os.makedirs(report_dir)

    # 生成带时间戳的报告文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_report = os.path.join(report_dir, f"test_report_{timestamp}.html")

    # 简化的pytest参数 - 移除有问题的依赖参数
    pytest_args = [
        "test_transaction/test_transaction.py",
        "-v",
        f"--html={html_report}",
        "--self-contained-html",
        "--capture=sys",
        "-s",
        # "-x",
    ]

    print("🚀 开始运行自动化测试...")
    print(f"📊 HTML报告将生成到: {html_report}")

    # 运行测试
    exit_code = pytest.main(pytest_args)

    if exit_code == 0:
        print("✅ 所有测试通过!")
    else:
        print(f"❌ 有测试失败，退出码: {exit_code}")

    return exit_code


if __name__ == "__main__":
    run_tests()