# -*- coding: utf-8 -*-
"""
template_reader.py — 读取设备参数模板文件（Template/*.xlsx）

blockParams sheet 关键列：
  paramType  — 参数标识（如 FREQ_Hz）
  descrption — 参数描述（模板原始拼写，如 System Frequency）
  unit       — 工程单位（如 Hz）
  BACnetIP   — 非空 → 该参数在 BACnet/IP 北向发布
  AcuCloud   — 非空 → 该参数由 AcuCloud 同步
"""
from __future__ import annotations

import warnings
from dataclasses import dataclass
from pathlib import Path

import openpyxl


@dataclass
class TemplateParam:
    param_key:   str   # paramType 列
    description: str   # descrption 列（模板原文拼写）
    unit:        str   # unit 列（空字符串表示无单位）
    bacnet_ip:   str   # BACnetIP 列原始值（非空 → BACnet 发布）
    acucloud:    str   # AcuCloud 列原始值（非空 → AcuCloud 同步）
    mqtt:        str = ""     # MQTT 列原始值（非空 → MQTT 发布）
    datalog:     str = ""     # DataLog 列原始值（非空 → Datalog 记录）


def _norm(s: str) -> str:
    """小写并去除连字符/下划线/空格，用于模糊匹配文件名。"""
    return s.lower().replace("-", "").replace("_", "").replace(" ", "")


def find_template_file(template_dir: str, device_name: str) -> str:
    """
    在 template_dir 下递归查找设备模板文件（大小写/连字符/下划线不敏感）。
    例：device_name='AcuRev4100' → 匹配 'AcuRev-4100_v1.01_20260427.xlsx'
    """
    needle = _norm(device_name)
    for p in Path(template_dir).rglob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        if needle in _norm(p.stem):
            return str(p)
    raise FileNotFoundError(
        f"在 {template_dir} 中未找到设备 '{device_name}' 的模板文件"
    )


def load_template(xlsx_path: str) -> list[TemplateParam]:
    """解析模板 blockParams sheet，返回全部参数列表。"""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    ws = wb["blockParams"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        return []

    headers = [str(h).strip() if h is not None else "" for h in rows[0]]

    def _col(name: str) -> int:
        try:
            return headers.index(name)
        except ValueError:
            raise ValueError(f"模板缺少列 '{name}'，当前列：{headers}")

    ci_key     = _col("paramType")
    ci_desc    = _col("descrption")   # 模板原始拼写（少一个 i）
    ci_unit    = _col("unit")
    ci_bacnet  = _col("BACnetIP")
    ci_cloud   = _col("AcuCloud")
    ci_mqtt    = headers.index("MQTT")    if "MQTT"    in headers else -1
    ci_datalog = headers.index("DataLog") if "DataLog" in headers else -1

    def _str(row: tuple, ci: int) -> str:
        return str(row[ci]).strip() if ci >= 0 and row[ci] is not None else ""

    params: list[TemplateParam] = []
    for row in rows[1:]:
        key = row[ci_key]
        if not key:
            continue
        params.append(TemplateParam(
            param_key   = str(key).strip(),
            description = _str(row, ci_desc),
            unit        = _str(row, ci_unit),
            bacnet_ip   = _str(row, ci_bacnet),
            acucloud    = _str(row, ci_cloud),
            mqtt        = _str(row, ci_mqtt),
            datalog     = _str(row, ci_datalog),
        ))

    return params


def get_bacnet_params(xlsx_path: str) -> list[TemplateParam]:
    """返回 BACnetIP 列非空的参数（BACnet/IP 发布范围）。"""
    return [p for p in load_template(xlsx_path) if p.bacnet_ip]


def get_cloud_params(xlsx_path: str) -> list[TemplateParam]:
    """返回 AcuCloud 列非空的参数（AcuCloud 同步范围）。"""
    return [p for p in load_template(xlsx_path) if p.acucloud]


def get_mqtt_params(xlsx_path: str) -> list[TemplateParam]:
    """返回 MQTT 列非空的参数（MQTT 发布范围）。"""
    return [p for p in load_template(xlsx_path) if p.mqtt]


def get_datalog_params(xlsx_path: str) -> list[TemplateParam]:
    """返回 DataLog 列非空的参数（Datalog 记录范围）。"""
    return [p for p in load_template(xlsx_path) if p.datalog]
