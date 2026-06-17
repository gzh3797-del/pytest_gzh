# -*- coding: utf-8 -*-
"""
helpers/template_matcher.py — 封装模板参数读取，供 BACnet/IP 参数列表一致性用例使用。

对外接口：
    get_bacnet_descriptions_4100()        -> set[str]
    get_bacnet_descriptions_2100()        -> set[str]
    get_bacnet_param_keys(template_name)  -> set[str]
    get_bacnet_template_map(template_name)-> dict[str, TemplateParam]
"""
from __future__ import annotations

import sys
from pathlib import Path

# 把仓库根目录加入 path，以便 import tools.Protocols 模块
_REPO_ROOT = str(Path(__file__).parents[3])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from tools.Protocols.template_reader import (
    TemplateParam,
    find_template_file,
    get_bacnet_params,
)
from tools.Protocols.config import TEMPLATE_DIR

# 协议层用例的设备关键词 → 主模板文件设备名（find_template_file 大小写/连字符不敏感）。
# PXM350 在模板中命名为 AcuRev-1300；AcuVIM3 命名为 Acuvim3。
DEVICE_TEMPLATE_NAMES: dict[str, str] = {
    "AcuRev4100": "AcuRev4100",
    "AcuRev2100": "AcuRev2100",
    "AcuvimIIR":  "AcuvimIIR",   # PXE1
    "AcuvimIIW":  "AcuvimIIW",   # PXE2
    "AcuVIM3":    "AcuVIM3",
    "AcuRev1300": "AcuRev1300",  # PXM350
}

# 设备 → Protocols Modbus 地址映射模块（用于 BACnet 上传值 vs Modbus 实时值比对）。
# 注意 PXM350 的 Modbus 模块名为 devices.pxm350（非 acurev1300）。
DEVICE_MODBUS_MODULES: dict[str, str] = {
    "AcuRev4100": "devices.acurev4100",
    "AcuRev2100": "devices.acurev2100",
    "AcuvimIIR":  "devices.acuvimiir",
    "AcuvimIIW":  "devices.acuvimiiw",
    "AcuVIM3":    "devices.acuvim3",
    "AcuRev1300": "devices.pxm350",
}


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


def get_bacnet_template_map(template_name: str) -> dict[str, TemplateParam]:
    """
    返回指定设备模板的 BACnet/IP 参数 param_key → TemplateParam 映射（模板基准）。

    template_name 取 DEVICE_TEMPLATE_NAMES 的值（如 "AcuvimIIR"、"AcuRev1300"）。
    模板文件不存在时由 find_template_file 抛 FileNotFoundError，调用方自行处理（skip）。
    """
    path = find_template_file(TEMPLATE_DIR, template_name)
    return {p.param_key: p for p in get_bacnet_params(path) if p.param_key}


def get_bacnet_param_keys(template_name: str) -> set[str]:
    """
    返回指定设备模板 BACnet/IP 发布参数的 param_key 集合（全量范围基准）。

    与网关 objectName 解析出的 param_key（name.partition("-") 后半段）对齐，
    用于严格比对「网关实际发布 vs 模板应发布」。
    """
    return set(get_bacnet_template_map(template_name))
