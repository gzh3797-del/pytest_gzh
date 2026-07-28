"""AcuRev-100（ACmeter）项目级 pytest fixture。

本项目无 Web UI（仅 RS-485/USB Modbus RTU + Acuview 2 上位机），conftest 保持精简：
- acuview 自动化用例直接调 comm.ctl_acuview 引擎并自带 config_path，无需额外 fixture；
- 这里仅提供一个读取本项目 config.yaml 的 project_config fixture，供将来直连 Modbus 的 pytest 用例复用。
仓库根 conftest.py 已把仓库根加入 sys.path（import comm.ctl_acuview 可用）。
"""
from __future__ import annotations

import logging
import os
import re
import sys
from pathlib import Path

import pytest
import yaml

# ── 跑单条用例不必再带 $env:PYTHONUTF8=1 / --log-cli-level(2026-07-27) ──────────
# 1) 控制台编码兜底: helper 的部分打印含 ⚠️ 等 GBK 编不了的字符(且多在失败/退场分支),
#    GBK 控制台下会 UnicodeEncodeError——报错信息自己先崩掉。errors=replace 保留控制台
#    原编码(中文正常), 编不了的字符退化成 ?, 打印永不抛异常。
# 2) pymodbus DEBUG 刷屏: 每次寄存器读写两行 DEBUG, 单条用例上千行, 控制台和 HTML 报告
#    的 Captured log 全被淹没。默认压到 WARNING; 要看协议帧时置 PYMODBUS_DEBUG=1。
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")
if not os.environ.get("PYMODBUS_DEBUG"):
    for _name in ("pymodbus", "pymodbus.logging"):
        logging.getLogger(_name).setLevel(logging.WARNING)

PROJECT_ROOT = Path(__file__).resolve().parent
CONFIG_PATH = PROJECT_ROOT / "config.yaml"
CASE_MAP_PATH = PROJECT_ROOT / "tests" / "case_map.yaml"

# CT 类型 -> 表型标记: 100mA/80mA=mA 表, 333mV=mV 表, RCT=独立表型(暂不纳入 mA/mV)。
# 无 ct_type(如电压/相角/计时等 CT 无关用例, 或 case_map 外的 014~017)不打标, -m ma/-m mv 均不选中。
_CT_TO_MARK = {"100mA": "ma", "80mA": "ma", "333mV": "mv", "RCT": "rct"}
_CASE_TOKEN_RE = re.compile(r"\d{3}_\d{2}_case\d+")


def _build_meter_type_map() -> "dict[str, str]":
    """从 case_map.yaml 建 用例token(如 004_01_case1) -> 标记(ma/mv/rct) 映射。

    同一模块内 mA/mV 的 case 编号不重叠(mV 占 case1~7, mA 占 case8+), 故 token 唯一。
    """
    mapping: "dict[str, str]" = {}
    try:
        with open(CASE_MAP_PATH, "r", encoding="utf-8") as fh:
            data = yaml.safe_load(fh) or {}
    except FileNotFoundError:
        return mapping
    cases = data.get("cases", []) if isinstance(data, dict) else data
    for case in cases:
        m = _CASE_TOKEN_RE.search(str(case.get("case_id", "")))
        if not m:
            continue
        ct = (case.get("meter_cfg") or {}).get("ct_type")
        mark = _CT_TO_MARK.get(str(ct))
        if mark:
            mapping[m.group(0)] = mark
    return mapping


_METER_TYPE_MAP = _build_meter_type_map()


def _build_angle_map() -> "dict[str, tuple]":
    """从 case_map.yaml 建 用例token -> 首测点角度签名(6角+freq) 映射。

    用于按角度分组排序, 见 pytest_collection_modifyitems。多测点用例取首点: 用例内部的
    角度切换避不掉, 能省的只有"用例之间"的来回切。
    """
    mapping: "dict[str, tuple]" = {}
    try:
        with open(CASE_MAP_PATH, "r", encoding="utf-8") as fh:
            data = yaml.safe_load(fh) or {}
    except FileNotFoundError:
        return mapping
    cases = data.get("cases", []) if isinstance(data, dict) else data
    for case in cases:
        m = _CASE_TOKEN_RE.search(str(case.get("case_id", "")))
        pts = case.get("points") or []
        if not m or not pts:
            continue
        src = (pts[0] or {}).get("source") or {}
        try:
            mapping[m.group(0)] = tuple(float(src[k]) for k in
                                        ("qua", "qub", "quc", "qia", "qib", "qic", "freq"))
        except (KeyError, TypeError, ValueError):
            continue
    return mapping


_ANGLE_MAP = _build_angle_map()


def _reorder_by_angle(items: list) -> list:
    """同目录内按"角度签名"分组排序: 相同角度的用例挨着跑。

    🔴 为什么值得排(2026-07-28 实测): CL3021 的角度帧一发, 自供电电表必重启一次。字母序会把
    非标角度用例夹在标准角度用例中间 —— 实测 005 批的 case7(电流角 60/300/180)前后各切一次:
    进 case7 掉电 1.0s, 进 case8 又掉电 1.0s, 而 case8 自己就是标准角度、纯属被连累。
    按角度分组后整批只在角度真变时切一次。

    只在"该目录下每条用例都能查到角度签名"时才动它, 否则原样返回 —— 不给 010/011 这类
    自带执行顺序讲究的模块添乱。同签名内部用稳定排序保持原有相对次序。
    """
    def _sig(item) -> "tuple | None":
        m = _CASE_TOKEN_RE.search(item.name)
        return _ANGLE_MAP.get(m.group(0)) if m else None

    by_dir: "dict[str, list]" = {}
    for it in items:
        by_dir.setdefault(str(Path(str(it.fspath)).parent), []).append(it)
    out = []
    for group in by_dir.values():
        sigs = [_sig(it) for it in group]
        if all(s is not None for s in sigs) and len(set(sigs)) > 1:
            group = [it for _, it in sorted(zip(sigs, group), key=lambda pair: pair[0])]
        out.extend(group)
    return out


def pytest_collection_modifyitems(items):
    """① 按 ct_type 补 ma/mv/rct 标记(供 -m 选跑表型); ② 按角度分组排序(省电表重启)。"""
    for item in items:
        m = _CASE_TOKEN_RE.search(item.name)
        if not m:
            continue
        mark = _METER_TYPE_MAP.get(m.group(0))
        if mark:
            item.add_marker(getattr(pytest.mark, mark))
    items[:] = _reorder_by_angle(items)


@pytest.fixture(scope="session")
def project_config() -> dict:
    """加载本项目 config.yaml 为 dict。"""
    with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
        return yaml.safe_load(fh)
