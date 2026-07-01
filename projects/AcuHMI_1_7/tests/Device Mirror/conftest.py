# -*- coding: utf-8 -*-
"""Device Mirror 协议本地 conftest：默认在本目录生成「用例执行结果.html」。

无需每次手动加 --html；直接 `pytest "Device Mirror"` 即可。
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
