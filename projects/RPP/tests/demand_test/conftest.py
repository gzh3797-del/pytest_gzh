"""RPP 需量测试的 pytest 配置(由 AcuRev1320/demand_test 复制迁移)。

新增 --measure 命令行选项, 让测量模式(mV/mA/rct)可在一条 pytest 命令内选择,
不再依赖 DEMAND_MEASURE_MODE 环境变量。接线方式 / 触发方式仍用 pytest 原生 -k
选参数化 id(如 -k "3e4wy-fixed"), 三项配置同命令搞定。
"""


def pytest_addoption(parser):
    group = parser.getgroup("demand", "RPP 需量测试")
    try:
        group.addoption(
            "--measure",
            action="store",
            default=None,
            help="测量模式: mv|ma|rct 或 0/1/2/3 (缺省 mv, 读 test_case_mV)。",
        )
    except ValueError:
        # 整仓库收集时 AcuRev1320/demand_test 的 conftest 可能已注册过 --measure,
        # 两模块共用同一选项, 重复注册直接跳过。
        pass
