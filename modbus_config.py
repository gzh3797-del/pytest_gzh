"""台架 / Modbus 环境配置。

`comm/` 下各模块通过 `from modbus_config import modbus_config` 取连接参数；本模块
从分层配置的全局层 `configs/global.yaml` 读取（与 AcuHMI_1_7 等新结构共用同一份配置），
敏感值（ssh 密码）走 `configs/.env` 的 `SSH_PASSWORD`，不写入 yaml、不入库。
"""
import json
import os
from pathlib import Path

import yaml

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - 未装 python-dotenv 时跳过 .env 加载
    load_dotenv = None

_REPO_ROOT = Path(__file__).resolve().parent
_GLOBAL_YAML = _REPO_ROOT / "configs" / "global.yaml"
_ENV_FILE = _REPO_ROOT / "configs" / ".env"
# write_json 的运行期快照（产物，git 忽略）
_RUNTIME_SNAPSHOT = _REPO_ROOT / "reports" / "modbus_config_runtime.json"


def _load() -> dict:
    if load_dotenv is not None and _ENV_FILE.exists():
        load_dotenv(_ENV_FILE)
    with _GLOBAL_YAML.open("r", encoding="utf-8") as handle:
        cfg = yaml.safe_load(handle) or {}
    ssh = cfg.get("ssh")
    if isinstance(ssh, dict) and not ssh.get("password"):
        env_pwd = os.getenv("SSH_PASSWORD")
        if env_pwd:
            ssh["password"] = env_pwd
    return cfg


modbus_config = _load()


def write_json(key, value):
    """运行期改写部分连接参数并落盘快照（兼容旧接口；不回写 yaml）。"""
    if key in ("baudrate", "parity", "slaveid"):
        modbus_config["rtu"][key] = value
    if key in ("ip", "port"):
        modbus_config["tcp"][key] = value
    _RUNTIME_SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
    with _RUNTIME_SNAPSHOT.open("w", encoding="utf-8") as handle:
        json.dump(modbus_config, handle, indent=4, ensure_ascii=False)
