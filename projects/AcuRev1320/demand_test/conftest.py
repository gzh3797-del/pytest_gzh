"""AcuRev-1320 需量测试的 pytest 配置。

新增 --measure 命令行选项, 让测量模式(mV/mA/rct)可在一条 pytest 命令内选择,
不再依赖 DEMAND_MEASURE_MODE 环境变量。接线方式 / 触发方式仍用 pytest 原生 -k
选参数化 id(如 -k "3e4wy-fixed"), 三项配置同命令搞定。
"""
import pytest


def pytest_addoption(parser):
    group = parser.getgroup("demand", "AcuRev1320 需量测试")
    group.addoption(
        "--measure",
        action="store",
        default=None,
        help="测量模式: mv|ma|rct 或 0/1/2/3 (缺省 mv, 读 test_case_mV)。",
    )
