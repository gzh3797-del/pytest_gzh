"""AcuRev-1320 项目配置适配层。

从分层配置（configs/global.yaml ← projects/AcuRev1320/config.yaml）+ .env 暴露常量，
供 tests/ 下各模块引用，避免在测试/引擎代码里硬编码连接参数与魔法值。
范式与 projects/AcuHMI_1_7/settings.py 一致。
"""
from __future__ import annotations

from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - 未装 python-dotenv 时跳过 .env 加载
    load_dotenv = None

from framework.config.loader import load_config

_REPO_ROOT = Path(__file__).resolve().parents[2]
if load_dotenv is not None:
    load_dotenv(_REPO_ROOT / "configs" / ".env")

_cfg = load_config("AcuRev1320")
_demand = _cfg.get("demand", {}) or {}

# ── 通用连接（来自 global.yaml，可被本项目 config.yaml 覆盖）──
CONN_MODE = _cfg.get("conn_mode", "rtu")
# demand 引擎据此把项目级连接参数同步进 comm 层共享的 modbus_config（见 demand_engine）
DEMAND_CONN_MODE = _cfg.get("conn_mode", "rtu")
DEMAND_RTU = dict(_cfg.get("rtu", {}) or {})   # 合并后的 rtu（含本项目 config.yaml 覆盖的 port）
DEMAND_TCP = dict(_cfg.get("tcp", {}) or {})

# ── 需量测试（tests/demand）配置 ──
DEMAND_SLAVE_ID = int(_demand.get("slave_id", 1))
DEMAND_TEST_TYPE = int(_demand.get("test_type", 1))
DEMAND_STABILIZE_SECONDS = int(_demand.get("stabilize_seconds", 300))
DEMAND_CASE_FILE = _demand.get("case_file", "demand_test_case.xlsx")
