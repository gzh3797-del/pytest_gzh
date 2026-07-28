"""Alarm 模块（Alarm Config 用例组）专有配置。

RPP 固件尚未提供 Alarm 页面，本模块当前按 AcuHMI-1-7 真机（192.168.3.71）执行；
RPP 真机就绪后仅需改 ALARM_BASE_URL / ALARM_TRIGGER_DEVICE（环境变量或本文件默认值）。
账号密码只从 configs/.env 读取（WEB_USERNAME / WEB_PASSWORD），不硬编码入库。
"""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

# projects/RPP/tests/Alarm/config_alarm.py → 仓库根需上溯 4 级
_REPO_ROOT = Path(__file__).resolve().parents[4]
load_dotenv(_REPO_ROOT / "configs" / ".env")

# 被测网关 Web 地址（默认 AcuHMI-1-7 真机，自签名证书 → 需 ignore_https_errors）
BASE_URL = os.getenv("ALARM_BASE_URL", "https://192.168.3.71")
USERNAME = os.getenv("WEB_USERNAME", "admin")
PASSWORD = os.getenv("WEB_PASSWORD", "")

# 触发告警用的下挂设备：必须是 Physical Devices 列表中轮询 Status=ON 的设备，
# 否则告警监控不评估、用例全部等待超时。
TRIGGER_DEVICE = os.getenv("ALARM_TRIGGER_DEVICE", "AcuRev4100_392")

# 网关约 60s 轮询一次下挂设备；告警触发/消除最多等 POLL_ROUNDS * POLL_STEP_MS。
POLL_STEP_MS = 10_000
POLL_ROUNDS = 15

# 本模块创建的告警规则统一使用该前缀，便于识别与兜底清理（不误删他人规则）。
RULE_PREFIX = "at_"
