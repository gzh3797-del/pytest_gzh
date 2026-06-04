# -*- coding: utf-8 -*-
"""
helpers/template_matcher.py — 封装模板参数读取，供 BACnet/IP 参数列表一致性用例使用。

对外接口：
    get_bacnet_descriptions_4100()   -> set[str]
    get_bacnet_descriptions_2100()   -> set[str]
"""
from __future__ import annotations

import sys
from pathlib import Path

# 把仓库根目录加入 path，以便 import Protocols 模块
_REPO_ROOT = str(Path(__file__).parents[4])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from Protocols.template_reader import find_template_file, get_bacnet_params
from Protocols.config import TEMPLATE_DIR


def _clean_description(desc: str) -> str:
    """取 description 第一行并去除首尾空白（模板部分单元格含多行副标题）。"""
    return desc.split("\n")[0].strip()


def get_bacnet_descriptions_4100() -> set[str]:
    """返回 AcuRev-4100 BACnet/IP 参数的 description 集合（模板基准）。"""
    path = find_template_file(TEMPLATE_DIR, "AcuRev4100")
    params = get_bacnet_params(path)
    return {_clean_description(p.description) for p in params if p.description}


def get_bacnet_descriptions_2100() -> set[str]:
    """返回 AcuRev-2100 BACnet/IP 参数的 description 集合（模板基准）。"""
    path = find_template_file(TEMPLATE_DIR, "AcuRev2100")
    params = get_bacnet_params(path)
    return {_clean_description(p.description) for p in params if p.description}
