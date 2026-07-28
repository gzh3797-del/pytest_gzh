# -*- coding: utf-8 -*-
"""Pass Through 协议本地 conftest：默认在本目录生成「用例执行结果.html」。

无需每次手动加 --html；直接 `pytest "Pass Through"` 即可。
如显式传 --html 则以命令行为准。
"""
import pathlib

import pytest

_HERE = pathlib.Path(__file__).resolve().parent


# page_factory 复用项目级 browser（不再调用 sync_playwright）。
# 覆盖根 conftest 的同名 autouse fixture，去掉对 pytest-playwright `page` 的依赖，
# 避免每条用例额外创建无用 context。
@pytest.fixture(autouse=True)
def _expose_page_for_report():
    yield


def pytest_configure(config):
    # 默认（未显式指定，或仅是 pytest.ini 里的 report.html）→ 重定向到本套件目录，
    # 保证每个套件只产出一个「用例执行结果.html」，不再额外生成根目录 report.html。
    hp = getattr(config.option, "htmlpath", None)
    if hp is None or pathlib.Path(hp).name == "report.html":
        config.option.htmlpath = str(_HERE / "用例执行结果.html")
        config.option.self_contained_html = True

# 关闭 pymodbus 冗长 DEBUG 日志（大幅减少输出/上下文）
import logging as _logging
_logging.getLogger("pymodbus").setLevel(_logging.WARNING)


# ── HTML 报告：把关联TC编号并入用例标识（Test 列显示为 函数名[TC编号]，与其他套件参数化风格一致）──
# 编号取自测试模块的 CASE_ID_MAP（函数名 → 关联TC编号），便于按编号检索对应用例。
def _case_id_of(item):
    cmap = getattr(getattr(item, "module", None), "CASE_ID_MAP", {}) or {}
    name = getattr(item, "originalname", None) or item.name.split("[")[0]
    return cmap.get(name, "")


# 前置检查用例函数名（来自测试模块 PRECHECK_CASES），收集阶段填充。
_PRECHECK = set()


def pytest_collection_modifyitems(items):
    for it in items:
        _PRECHECK.update(getattr(getattr(it, "module", None), "PRECHECK_CASES", ()) or ())
        cid = _case_id_of(it)
        if cid and f"[{cid}]" not in it.nodeid:
            it._nodeid = f"{it._nodeid}[{cid}]"


# ── HTML 报告：前置检查用例（无对应手工用例编号）通过/跳过时不在结果表展示；
#    失败时仍展示，便于排查环境/配置问题。──
@pytest.hookimpl(optionalhook=True)
def pytest_html_results_table_row(report, cells):
    name = report.nodeid.split("::")[-1].split("[")[0]
    if name in _PRECHECK and not report.failed:
        del cells[:]
