# -*- coding: utf-8 -*-
import pathlib, sys

import pytest

_PROTO = pathlib.Path(__file__).resolve().parent.parent
_DC    = pathlib.Path(__file__).resolve().parent
for p in (_PROTO, _DC):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))


# DataCollect 用例通过 dc_page（session 级）管理自己的 context/page，不需要
# pytest-playwright 的 per-test `page` fixture。覆盖项目级同名 autouse fixture，
# 断开对 `page` 的依赖，避免每条用例额外创建一个无用浏览器 context。
@pytest.fixture(autouse=True)
def _expose_page_for_report():
    yield


def pytest_configure(config):
    """每个套件只产出一个「用例执行结果.html」（默认或 pytest.ini 的 report.html 都重定向到本目录）。"""
    hp = getattr(config.option, "htmlpath", None)
    if hp is None or pathlib.Path(hp).name == "report.html":
        config.option.htmlpath = str(_DC / "用例执行结果.html")
        config.option.self_contained_html = True

# 关闭 pymodbus 冗长 DEBUG 日志（大幅减少输出/上下文）
import logging as _logging
_logging.getLogger("pymodbus").setLevel(_logging.WARNING)


# ── HTML 报告：把用例编号并入用例标识（Test 列显示为 方法名[用例编号]，与其他套件参数化风格一致）──
# DataCollect 映射无「关联TC编号」，用其 DC 用例编号；编号取自测试模块的 CASE_ID_MAP。
def _case_id_of(item):
    cmap = getattr(getattr(item, "module", None), "CASE_ID_MAP", {}) or {}
    name = getattr(item, "originalname", None) or item.name.split("[")[0]
    return cmap.get(name, "")


def pytest_collection_modifyitems(items):
    for it in items:
        cid = _case_id_of(it)
        if cid and f"[{cid}]" not in it.nodeid:
            it._nodeid = f"{it._nodeid}[{cid}]"
