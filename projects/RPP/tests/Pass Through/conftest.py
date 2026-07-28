# -*- coding: utf-8 -*-
"""RPP Pass Through 套件本地 conftest：默认在本目录生成「用例执行结果.html」。

无需每次手动加 --html；直接 `pytest "projects/RPP/tests/Pass Through"` 即可。
如显式传 --html 则以命令行为准。
browser 使用 pytest-playwright 内置共享实例；HEADED=1 时有头模式运行。
"""
import os
import pathlib

import pytest

_HERE = pathlib.Path(__file__).resolve().parent


def pytest_configure(config):
    hp = getattr(config.option, "htmlpath", None)
    if hp is None or pathlib.Path(hp).name == "report.html":
        config.option.htmlpath = str(_HERE / "用例执行结果.html")
        config.option.self_contained_html = True


@pytest.fixture(scope="session")
def browser_type_launch_args(browser_type_launch_args):
    """HEADED=1 → 有头模式（与 1.7 的使用习惯一致，便于观察操作过程）。"""
    if os.getenv("HEADED", "0").strip().lower() in ("1", "true", "yes", "on"):
        return {**browser_type_launch_args, "headless": False}
    return browser_type_launch_args


# 关闭 pymodbus 冗长 DEBUG 日志
import logging as _logging
_logging.getLogger("pymodbus").setLevel(_logging.WARNING)


def _case_id_of(item):
    cmap = getattr(getattr(item, "module", None), "CASE_ID_MAP", {}) or {}
    name = getattr(item, "originalname", None) or item.name.split("[")[0]
    return cmap.get(name, "")


def pytest_collection_modifyitems(items):
    for it in items:
        cid = _case_id_of(it)
        if cid and f"[{cid}]" not in it.nodeid:
            it._nodeid = f"{it._nodeid}[{cid}]"
