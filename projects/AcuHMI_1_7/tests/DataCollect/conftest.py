# -*- coding: utf-8 -*-
import pathlib, sys

_PROTO = pathlib.Path(__file__).resolve().parent.parent
_DC    = pathlib.Path(__file__).resolve().parent
for p in (_PROTO, _DC):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))


def pytest_configure(config):
    """默认在本目录生成「用例执行结果.html」，无需手动加 --html；显式传 --html 时以命令行为准。"""
    if getattr(config.option, "htmlpath", None) is None:
        config.option.htmlpath = str(_DC / "用例执行结果.html")
        config.option.self_contained_html = True

# 关闭 pymodbus 冗长 DEBUG 日志（大幅减少输出/上下文）
import logging as _logging
_logging.getLogger("pymodbus").setLevel(_logging.WARNING)
