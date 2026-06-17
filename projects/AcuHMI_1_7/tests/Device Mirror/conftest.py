# -*- coding: utf-8 -*-
"""Device Mirror 协议本地 conftest：默认在本目录生成「用例执行结果.html」。

无需每次手动加 --html；直接 `python -m pytest "Device Mirror"` 即可。
如显式传 --html 则以命令行为准。
"""
import pathlib

_HERE = pathlib.Path(__file__).resolve().parent


def pytest_configure(config):
    if getattr(config.option, "htmlpath", None) is None:
        config.option.htmlpath = str(_HERE / "用例执行结果.html")
        config.option.self_contained_html = True

# 关闭 pymodbus 冗长 DEBUG 日志（大幅减少输出/上下文）
import logging as _logging
_logging.getLogger("pymodbus").setLevel(_logging.WARNING)
