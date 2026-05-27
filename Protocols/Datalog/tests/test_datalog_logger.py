# -*- coding: utf-8 -*-
"""
test_datalog_logger.py — DataLog Logger 自动化测试套件

对应 Excel：DataLog测试用例.xlsx（Sheet1，47 条）

用例分类：
  A  Disable 验证      —— Logger disable，验证服务器目录无文件（6 条）
  B  None PostChannel  —— PostChannel=None，validate 不推送（6 条）
  C  正向推送验证       —— enable + FTP/SFTP/HTTP，LV0+LV1（12 条）
       case04 Logger1 (FTP+CSV+1min) 执行完整三段验证（范围+单位+Modbus数值）
       其余 C 类只验证文件存在 + 扩展名正确（避免重复 Modbus 轮询）
  D  Prefix 超长负面    —— prefix 超长，保存失败且提示准确（3 条）
  E  联动验证           —— LogFileLength 变化时 LogInterval 可选项正确（3 条）

运行方式（从仓库根）：
  pytest Protocols/Datalog/tests/test_datalog_logger.py -v
  pytest Protocols/Datalog/tests/test_datalog_logger.py -k "case04"
  pytest Protocols/Datalog/tests/test_datalog_logger.py -m "lv0"
"""
from __future__ import annotations

import asyncio
import csv
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))

import config
from helpers import collect_files
from datalog_page import DataLoggerConfig, DataLoggerPage, PostChannelPage
from datalog_server_verifier import (
    ServerInfo,
    _select_latest_per_device,
    verify_files,
    wait_for_files,
)

# ─── pytest marks ─────────────────────────────────────────────────────────────

def pytest_configure(config_obj):
    config_obj.addinivalue_line("markers", "lv0: 最高优先级用例")
    config_obj.addinivalue_line("markers", "lv1: 高优先级用例")
    config_obj.addinivalue_line("markers", "lv4: 负面/边界用例")


# ─── 工具函数 ─────────────────────────────────────────────────────────────────

# LogFileLength UI 文字 → 推送等待超时（秒）
# 实际等待 = LogFileLength × 1.5，最少 90s
_LENGTH_TIMEOUT: dict[str, float] = {
    "1 minute":   120,
    "5 minute":   450,
    "10 minute":  720,
}

# Post Channel 编号映射（与 conftest 中配置的 session 级 channel 一致）
_PROTO_TO_CHANNEL: dict[str, int] = {
    "FTP":   1,
    "SFTP":  2,
    "HTTP":  3,
    "HTTPS": 3,
}

# LogFileLength → 该 length 下 LogInterval 可选项（UI 文本，可按实际调整）
# 来源：test case 21/13/13 预期结果
EXPECTED_INTERVALS: dict[str, list[str]] = {
    "1 minute":  ["1 minute"],
    "5 minute":  ["1 minute", "5 minute"],
    "10 minute": ["1 minute", "5 minute", "10 minute"],
    "15 minute": ["1 minute", "5 minute", "10 minute", "15 minute"],
    "30 minute": ["1 minute", "5 minute", "10 minute", "15 minute", "30 minute"],
    "1 hour":    ["1 minute", "5 minute", "10 minute", "15 minute", "30 minute", "1 hour"],
    "6 hour":    ["1 minute", "5 minute", "10 minute", "15 minute", "30 minute", "1 hour", "6 hour"],
    "12 hour":   ["1 minute", "5 minute", "10 minute", "15 minute", "30 minute", "1 hour", "6 hour", "12 hour"],
    "24 hour":   ["5 minute", "10 minute", "15 minute", "30 minute", "1 hour", "6 hour", "12 hour", "1 day"],
    "7 day":     ["15 minute", "30 minute", "1 hour", "6 hour", "12 hour", "1 day", "7 day"],
    "1 month":   ["1 hour", "6 hour", "12 hour", "1 day", "7 day", "1 month"],
}

# LogInterval / LogFileLength UI 文字 → 秒数（共用同一张表，UI 选项文字相同）
_INTERVAL_SECONDS: dict[str, float] = {
    "1 minute":  60,
    "5 minute":  300,
    "10 minute": 600,
    "15 minute": 900,
    "30 minute": 1800,
    "1 hour":    3600,
    "6 hour":    21600,
    "12 hour":   43200,
    "1 day":     86400,
    "7 day":     604800,
    "1 month":   2592000,
}
_LENGTH_SECONDS = _INTERVAL_SECONDS

# Timestamp Format UI 文字 → 验证文件内容时间戳的正则（按优先级）
_TS_FORMAT_RE: dict[str, list] = {
    "Local Time String": [
        re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}"),   # YYYY-MM-DD HH:MM[:SS]
        re.compile(r"^\d{2}/\d{2}/\d{4} \d{2}:\d{2}"),   # MM/DD/YYYY HH:MM[:SS]
    ],
    "UTC Seconds": [
        re.compile(r"^\d{10,13}$"),   # Unix 秒 / 毫秒
    ],
    "ISO8601 Format": [
        re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}"),
    ],
}

# Timestamp Format UI 文字 → datetime 解析器（失败则尝试下一个）
_TS_PARSERS: dict[str, list] = {
    "Local Time String": [
        lambda s: datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S"),
        lambda s: datetime.strptime(s[:16], "%Y-%m-%d %H:%M"),
        lambda s: datetime.strptime(s[:19], "%m/%d/%Y %H:%M:%S"),
        lambda s: datetime.strptime(s[:14], "%m/%d/%Y %H:%M"),
    ],
    "UTC Seconds": [
        lambda s: datetime.utcfromtimestamp(float(s)),
    ],
    "ISO8601 Format": [
        lambda s: datetime.fromisoformat(s.replace("Z", "+00:00")),
        lambda s: datetime.strptime(s[:19], "%Y-%m-%dT%H:%M:%S"),
    ],
}


def _verify_filename_format(filepath: str, prefix: str, name_fmt: str) -> str:
    """
    验证文件名前缀和时间戳格式是否与 Log File Name Prefix / Log File Name Format 配置一致。
    返回错误描述；无误则返回空字符串。

    网关实际命名格式（per-device 文件）：
      {prefix}{serial}-{model}-{YYYY-MM-DDTHH-MM-SS±HHMM}-{interval}.ext
    e.g. meter0_logger1AHI260110001-AcuvimIIW-2026-05-24T22-01-00-0400-1min.csv
    """
    stem = Path(filepath).stem
    if not stem.startswith(prefix):
        return f"文件名前缀不匹配：期望以 '{prefix}' 开头，实际 '{stem}'"
    rest = stem[len(prefix):]
    # Gateway actual filename patterns (per-device file):
    #   {serial}-{model}-{ISO8601_local_time}-{interval}   (LocalTime / TimeInterval format)
    #   {serial}-{model}-{unix_epoch_10digits}-{interval}  (UTC Timestamp format)
    _ISO_TS  = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}")
    _UNIX_TS = re.compile(r"(?<![A-Za-z\d])\d{10}(?!\d)")  # exactly 10-digit Unix epoch
    if name_fmt == "UTC Timestamp":
        if not (_UNIX_TS.search(rest) or _ISO_TS.search(rest) or
                re.match(r"^[_-]?\d{14}$", rest)):
            return (f"文件名时间戳格式不符"
                    f"（期望 UTC Timestamp：prefix + YYYYMMDDHHmmss），实际后缀：'{rest}'")
    elif name_fmt == "Time Interval Format":
        if not (_ISO_TS.search(rest) or re.match(r"^[_-]?\d{14}[_-]\d{14}$", rest)):
            return (f"文件名区间格式不符"
                    f"（期望 Time Interval Format：prefix + start + end），实际后缀：'{rest}'")
    return ""


def _read_timestamps(filepath: str, file_format: str) -> list[str]:
    """读取文件各数据行的时间戳字符串（CSV 取第一列；JSON 取 'time'/'Time' 等字段）。"""
    timestamps: list[str] = []
    try:
        if file_format.lower() == "csv":
            with open(filepath, encoding="utf-8", errors="replace") as f:
                reader = csv.reader(f)
                next(reader, None)   # 跳过表头行
                for row in reader:
                    if row and row[0].strip():
                        timestamps.append(row[0].strip())
        else:
            with open(filepath, encoding="utf-8", errors="replace") as f:
                data = json.load(f)
            if isinstance(data, list):
                for rec in data:
                    if isinstance(rec, dict):
                        for key in ("time", "Time", "timestamp", "Timestamp",
                                    "datetime", "DateTime"):
                            if key in rec:
                                timestamps.append(str(rec[key]))
                                break
    except Exception:
        pass
    return timestamps


def _parse_timestamps_to_dt(timestamps: list[str], ts_fmt: str) -> list[datetime]:
    """将时间戳字符串列表解析为 datetime 列表（解析失败的跳过）。"""
    parsers = _TS_PARSERS.get(ts_fmt, [])
    result = []
    for s in timestamps:
        for parser in parsers:
            try:
                dt = parser(s)
                if dt.tzinfo is not None:
                    dt = dt.replace(tzinfo=None)
                result.append(dt)
                break
            except Exception:
                continue
    return result


def _verify_timestamp_format(timestamps: list[str], ts_fmt: str) -> str:
    """
    验证文件中至少前 3 行时间戳格式与 Timestamp Format 配置一致。
    返回错误描述；无误则返回空字符串。
    """
    if not timestamps:
        return "文件中无数据行，无法验证时间戳格式"
    patterns = _TS_FORMAT_RE.get(ts_fmt, [])
    if not patterns:
        return ""
    for ts in timestamps[:3]:
        if not any(p.match(ts) for p in patterns):
            return f"时间戳格式不符（期望 {ts_fmt}），实际值：'{ts}'"
    return ""


def _verify_log_interval(timestamps: list[str], ts_fmt: str,
                          interval_str: str) -> str:
    """
    验证相邻行时间间隔是否接近配置的 Log Interval。
    行数少于 2 时跳过。允许 ±10%（最少 ±30s）的容差。
    """
    if len(timestamps) < 2:
        return ""
    expected_sec = _INTERVAL_SECONDS.get(interval_str, 60)
    dt_list = _parse_timestamps_to_dt(timestamps, ts_fmt)
    if len(dt_list) < 2:
        return ""
    tolerance = max(expected_sec * 0.1, 30)
    bad: list[str] = []
    for i in range(min(len(dt_list) - 1, 5)):
        diff = abs((dt_list[i + 1] - dt_list[i]).total_seconds())
        if abs(diff - expected_sec) > tolerance:
            bad.append(f"{diff:.0f}s")
    if bad:
        return (f"Log Interval 不符（期望 {expected_sec:.0f}s±{tolerance:.0f}s），"
                f"实际间隔：{bad}")
    return ""


def _verify_file_length(timestamps: list[str], ts_fmt: str,
                         length_str: str, interval_str: str) -> str:
    """
    验证文件行数是否与 Log File Length / Log Interval 配置一致
    （期望行数 ≈ length_sec / interval_sec，容差 ±2 行）。
    同时校验首尾时间戳跨度不超出 length_sec 的 130%+30s。
    """
    if not timestamps:
        return ""
    expected_length_sec   = _LENGTH_SECONDS.get(length_str,   60)
    expected_interval_sec = _INTERVAL_SECONDS.get(interval_str, 60)
    expected_rows = round(expected_length_sec / expected_interval_sec)
    actual_rows   = len(timestamps)
    if abs(actual_rows - expected_rows) > 2:
        return (f"文件行数与配置不符（期望约 {expected_rows} 行 "
                f"[{length_str} / {interval_str}]，实际 {actual_rows} 行）")
    dt_list = _parse_timestamps_to_dt([timestamps[0], timestamps[-1]], ts_fmt)
    if len(dt_list) == 2:
        span = abs((dt_list[1] - dt_list[0]).total_seconds())
        max_allowed = expected_length_sec * 1.3 + 30
        if span > max_allowed:
            return (f"文件时间跨度超出预期（期望 ≤{max_allowed:.0f}s，实际 {span:.0f}s）")
    return ""


def _data_dirs_for_protos(pool: dict[str, ServerInfo],
                           protos: list[str]) -> list[str]:
    return [os.path.normpath(pool[p].data_dir)
            for p in protos if p in pool]


def _configure_logger(dl_page: DataLoggerPage, logger_n: int,
                      cfg: DataLoggerConfig):
    """包装 configure_logger，并在完成后记录日志。"""
    dl_page.configure_logger(logger_n, cfg)


def _configure_logger_none_channel(dl_page: DataLoggerPage, logger_n: int,
                                   file_format: str, file_length: str,
                                   timestamp_fmt: str, name_fmt: str,
                                   prefix: str, interval: str):
    """
    配置 Data Logger 为 PostChannel=None（本地存储，不推送）。
    因 DataLoggerConfig.channel_index 不支持 none，改用此函数直接操作页面元素。
    注：UI 下拉中 None 选项的实际文本请按网关版本核实（可能是 "None" 或 "-- None --"）。
    """
    dl_page.navigate_to_logger(logger_n)
    dl_page._set_enable(True)
    dl_page._select_el_by_text(dl_page.POST_CHANNEL_SELECT, "None", "Post Channel")
    dl_page._set_timestamp_format(timestamp_fmt)
    dl_page._set_log_file_name_format(name_fmt)
    dl_page._select_el_by_text(dl_page.LOG_FILE_FORMAT_SELECT, file_format, "Log File Format")
    if prefix:
        dl_page._fill(dl_page.LOG_FILE_NAME_PREFIX_INPUT, prefix, "Log File Name Prefix")
    dl_page._select_el_by_text(dl_page.LOG_FILE_LENGTH_SELECT, file_length, "Log File Length")
    dl_page._select_el_by_text(dl_page.LOG_INTERVAL_SELECT, interval, "Log Interval")
    dl_page._select_devices([])   # 全选设备
    dl_page._safe_click(dl_page.SAVE_BTN, "Save")
    time.sleep(2)


def _get_interval_options(dl_page: DataLoggerPage, logger_n: int,
                          file_length: str) -> list[str]:
    """
    导航到 Data Logger N 页面，选定 LogFileLength 后读取 LogInterval 下拉所有可用选项。
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    dl_page.navigate_to_logger(logger_n)
    # 先选 file length（触发联动更新）
    dl_page._select_el_by_text(
        dl_page.LOG_FILE_LENGTH_SELECT, file_length, "Log File Length")
    time.sleep(0.8)

    # 展开 interval 下拉，读取选项
    options: list[str] = []
    try:
        container = dl_page.driver.find_element(
            *dl_page.LOG_INTERVAL_SELECT)
        inner = container.find_elements(
            By.XPATH, ".//div[contains(@class,'el-select__wrapper')]")
        dl_page._js_click(inner[0] if inner else container)
        time.sleep(0.8)
        items = dl_page.driver.find_elements(
            By.XPATH,
            "//li[contains(@class,'el-select-dropdown__item')"
            " and not(contains(@class,'disabled'))]")
        options = [el.text.strip() for el in items if el.text.strip()]
    except Exception:
        pass
    finally:
        try:
            dl_page.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(0.2)
        except Exception:
            pass
    return options


def _check_save_error(driver) -> str:
    """
    保存后检测页面是否出现错误提示。
    返回错误文本（若无则返回空字符串）。
    """
    from selenium.webdriver.common.by import By
    time.sleep(1.5)
    for xpath in [
        "//div[contains(@class,'el-message') and contains(@class,'el-message--error')]",
        "//div[contains(@class,'el-form-item__error')]",
        "//div[contains(@class,'toast') and contains(translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'error')]",
        "//span[contains(@class,'error')]",
        "//*[contains(@class,'is-error')]//div[contains(@class,'el-form-item__error')]",
    ]:
        els = driver.find_elements(By.XPATH, xpath)
        for el in els:
            txt = el.text.strip()
            if txt:
                return txt
    return ""


# ─── A 类：Disable 验证 ───────────────────────────────────────────────────────

# (case_id, logger_n, check_protocols)
_DISABLE_CASES = [
    ("TestCase_AcuHMI_003_06_case03", 1, ["FTP", "SFTP"]),
    ("TestCase_AcuHMI_003_01_case02", 1, ["HTTP", "HTTPS"]),
    ("TestCase_AcuHMI_003_02_case01", 2, ["FTP", "SFTP"]),
    ("TestCase_AcuHMI_003_02_case02", 2, ["HTTP", "HTTPS"]),
    ("TestCase_AcuHMI_003_03_case01", 3, ["FTP", "SFTP"]),
    ("TestCase_AcuHMI_003_03_case02", 3, ["HTTP", "HTTPS"]),
]


@pytest.mark.lv1
@pytest.mark.parametrize(
    "case_id, logger_n, check_protos",
    _DISABLE_CASES,
    ids=[c[0] for c in _DISABLE_CASES],
)
def test_disable_no_push(case_id, logger_n, check_protos, pool, driver):
    """
    A 类：Logger N 置为 disable，验证远端目录下不存在该 Logger 的日志文件。
    对应用例：TestCase_AcuHMI_003_0{1/2/3}_case01 / case02
    """
    dl_page = DataLoggerPage(driver)
    dl_page.disable_logger(logger_n)

    # 等待一个 1-minute LogFileLength 周期后确认无文件
    time.sleep(90)

    dirs = _data_dirs_for_protos(pool, check_protos)
    found = collect_files(dirs)
    assert not found, (
        f"[{case_id}] Logger{logger_n} 已 disable，但在目录中发现文件：{found}"
    )


# ─── B 类：None PostChannel 验证 ──────────────────────────────────────────────

# (case_id, logger_n, file_format)
_NONE_CHANNEL_CASES = [
    ("TestCase_AcuHMI_003_01_case03", 1, "csv",  "Local Time String", "UTC Timestamp",     "meter0_logger1"),
    ("TestCase_AcuHMI_003_01_case15", 1, "json", "Local Time String", "UTC Timestamp",     "meter0_logger1"),
    ("TestCase_AcuHMI_003_02_case03", 2, "csv",  "Local Time String", "UTC Timestamp",     "meter0_Logger2"),
    ("TestCase_AcuHMI_003_02_case07", 2, "json", "Local Time String", "UTC Timestamp",     "meter0_Logger2"),
    ("TestCase_AcuHMI_003_03_case03", 3, "csv",  "Local Time String", "UTC Timestamp",     "meter0_Logger3"),
    ("TestCase_AcuHMI_003_03_case07", 3, "json", "Local Time String", "UTC Timestamp",     "meter0_Logger3"),
]


@pytest.mark.lv1
@pytest.mark.parametrize(
    "case_id, logger_n, fmt, ts_fmt, name_fmt, prefix",
    _NONE_CHANNEL_CASES,
    ids=[c[0] for c in _NONE_CHANNEL_CASES],
)
def test_none_channel_no_push(case_id, logger_n, fmt, ts_fmt, name_fmt, prefix,
                               pool, driver):
    """
    B 类：PostChannel = None，Logger enable 后远端目录下不存在该 Logger 的日志文件。
    对应用例：case03 / case15 / Logger2-case03 / Logger2-case07 / Logger3-case03 / Logger3-case07
    """
    dl_page = DataLoggerPage(driver)
    _configure_logger_none_channel(
        dl_page, logger_n,
        file_format=fmt,
        file_length="1 minute",
        timestamp_fmt=ts_fmt,
        name_fmt=name_fmt,
        prefix=prefix,
        interval="1 minute",
    )

    time.sleep(90)

    all_dirs = _data_dirs_for_protos(pool, ["FTP", "SFTP", "HTTP", "HTTPS"])
    found = collect_files(all_dirs)
    assert not found, (
        f"[{case_id}] PostChannel=None，但在目录中发现文件：{found}"
    )

    # 测试结束后禁用该 Logger
    dl_page.disable_logger(logger_n)


# ─── C 类：正向推送验证 ────────────────────────────────────────────────────────

@dataclass
class PushCase:
    case_id: str
    logger_n: int
    protocol: str          # "FTP" | "SFTP" | "HTTP"
    file_format: str       # "csv" | "json"
    file_length: str       # UI 文本，如 "1 minute"
    timestamp_fmt: str
    name_fmt: str
    prefix: str
    interval: str
    level: str
    full_verify: bool = False   # True = 完整三段验证（仅 case04 Logger1）


_PUSH_CASES: list[PushCase] = [
    # ── Logger 1 ──────────────────────────────────────────────────────────────
    PushCase("TestCase_AcuHMI_003_01_case04", 1, "FTP",  "csv",  "1 minute",
             "Local Time String", "UTC Timestamp",     "meter0_logger1", "1 minute",
             "LV0", full_verify=True),   # ← 唯一执行完整三段验证的用例
    PushCase("TestCase_AcuHMI_003_01_case05", 1, "SFTP", "csv",  "5 minute",
             "UTC Seconds",       "Time Interval Format", "meter1_logger1", "1 minute",
             "LV0"),
    PushCase("TestCase_AcuHMI_003_01_case06", 1, "HTTP", "csv",  "10 minute",
             "ISO8601 Format",    "Time Interval Format", "meter2_logger1", "1 minute",
             "LV0"),
    PushCase("TestCase_AcuHMI_003_01_case16", 1, "FTP",  "json", "1 minute",
             "UTC Seconds",       "UTC Timestamp",     "meter0_logger1", "1 minute",
             "LV2"),
    PushCase("TestCase_AcuHMI_003_01_case17", 1, "SFTP", "json", "5 minute",
             "UTC Seconds",       "Time Interval Format", "meter1_logger1", "1 minute",
             "LV0"),
    PushCase("TestCase_AcuHMI_003_01_case18", 1, "HTTP", "json", "10 minute",
             "ISO8601 Format",    "Time Interval Format", "meter2_logger1", "1 minute",
             "LV2"),
    # ── Logger 2 ──────────────────────────────────────────────────────────────
    PushCase("TestCase_AcuHMI_003_02_case04", 2, "FTP",  "csv",  "1 minute",
             "Local Time String", "UTC Timestamp",     "meter0_Logger2", "1 minute",
             "LV2"),
    PushCase("TestCase_AcuHMI_003_02_case05", 2, "SFTP", "csv",  "5 minute",
             "UTC Seconds",       "Time Interval Format", "meter1_Logger2", "1 minute",
             "LV1"),
    PushCase("TestCase_AcuHMI_003_02_case06", 2, "HTTP", "csv",  "10 minute",
             "ISO8601 Format",    "Time Interval Format", "meter2_Logger2", "1 minute",
             "LV1"),
    PushCase("TestCase_AcuHMI_003_02_case08", 2, "FTP",  "json", "1 minute",
             "UTC Seconds",       "UTC Timestamp",     "meter0_Logger2", "1 minute",
             "LV1"),
    # ── Logger 3 ──────────────────────────────────────────────────────────────
    PushCase("TestCase_AcuHMI_003_03_case04", 3, "FTP",  "csv",  "1 minute",
             "Local Time String", "UTC Timestamp",     "meter0_Logger3", "1 minute",
             "LV1"),
    PushCase("TestCase_AcuHMI_003_03_case05", 3, "SFTP", "csv",  "5 minute",
             "UTC Seconds",       "Time Interval Format", "meter1_Logger3", "1 minute",
             "LV1"),
]


def _make_push_case_id(c: PushCase) -> str:
    return f"{c.case_id}[L{c.logger_n}-{c.protocol}-{c.file_format}-{c.file_length.replace(' ', '')}]"


@pytest.mark.parametrize(
    "case",
    _PUSH_CASES,
    ids=[_make_push_case_id(c) for c in _PUSH_CASES],
)
def test_push_verify(case: PushCase, pool, driver):
    """
    C 类：Logger N enable，配置 PostChannel + 格式 + 时间参数，验证远端服务器收到日志文件。

    case04 Logger1（FTP+CSV+1min）：执行完整三段验证（范围 + 单位 + Modbus 数值比对）。
    其余用例：只验证文件存在 + 扩展名与格式配置一致。
    """
    # 1. 根据协议确定 channel 编号，并启用对应的 Post Channel
    channel_n = _PROTO_TO_CHANNEL[case.protocol]
    PostChannelPage(driver).enable_channel(channel_n)

    # 2. 配置 Data Logger
    dl_page = DataLoggerPage(driver)
    _configure_logger(
        dl_page, case.logger_n,
        DataLoggerConfig(
            channel_index=channel_n,
            enabled=True,
            log_file_format=case.file_format,
            log_file_length=case.file_length,
            timestamp_format=case.timestamp_fmt,
            log_file_name_format=case.name_fmt,
            log_file_name_prefix=case.prefix,
            log_interval=case.interval,
            device_names=[],   # 全选设备
        ),
    )

    # 3. 等待文件到达
    target_dir = os.path.normpath(pool[case.protocol].data_dir)
    timeout = _LENGTH_TIMEOUT.get(case.file_length, 120)
    file_paths = wait_for_files([target_dir], timeout=timeout)

    # 收到文件后立即禁用该 Logger，防止继续推送
    if file_paths:
        dl_page.disable_logger(case.logger_n)

    # 4. 基础断言：目录中有文件
    assert file_paths, (
        f"[{case.case_id}] 超时 {timeout}s 内未收到任何文件 "
        f"（Logger{case.logger_n} {case.protocol} {case.file_format} {case.file_length}）"
    )

    # 5. 扩展名与 Log File Format 配置一致
    expected_ext = f".{case.file_format.lower()}"
    for fp in file_paths:
        assert Path(fp).suffix.lower() == expected_ext, (
            f"[{case.case_id}] 文件扩展名不匹配：{Path(fp).name}，期望 {expected_ext}"
        )

    # 6. 文件名格式：前缀 + Log File Name Format（UTC Timestamp / Time Interval Format）
    for fp in file_paths:
        err = _verify_filename_format(fp, case.prefix, case.name_fmt)
        assert not err, (
            f"[{case.case_id}] 文件名格式错误（{Path(fp).name}）：{err}"
        )

    # 7. 文件内容校验（对每个收到的文件）
    for fp in file_paths:
        timestamps = _read_timestamps(fp, case.file_format)

        # 7a. Timestamp Format：首列/time 字段格式与配置一致
        ts_err = _verify_timestamp_format(timestamps, case.timestamp_fmt)
        assert not ts_err, (
            f"[{case.case_id}] Timestamp Format 不符（{Path(fp).name}）：{ts_err}"
        )

        # 7b. Log Interval：相邻行时间间隔与配置一致（行数 <2 时自动跳过）
        interval_err = _verify_log_interval(timestamps, case.timestamp_fmt, case.interval)
        assert not interval_err, (
            f"[{case.case_id}] Log Interval 错误（{Path(fp).name}）：{interval_err}"
        )

        # 7c. Log File Length：文件行数与 length/interval 比例一致
        length_err = _verify_file_length(
            timestamps, case.timestamp_fmt, case.file_length, case.interval)
        assert not length_err, (
            f"[{case.case_id}] Log File Length 错误（{Path(fp).name}）：{length_err}"
        )

    # 8. 完整三段验证（范围 + 单位 + Modbus 数值比对，仅 case04 Logger1）
    if case.full_verify:
        _run_full_verify(case, file_paths, pool)


def _run_full_verify(case: PushCase, file_paths: list[str],
                     pool: dict[str, ServerInfo]):
    """
    执行完整三段验证：范围检查 + 单位检查（JSON）+ Modbus 数值比对。
    复用 verify_files() 中的逻辑。
    """
    file_paths = _select_latest_per_device(file_paths)

    target_dir = os.path.normpath(pool[case.protocol].data_dir)
    dir_to_protocol = {target_dir: case.protocol.upper()}
    dir_to_channel  = {target_dir: _PROTO_TO_CHANNEL[case.protocol]}

    results = asyncio.run(
        verify_files(file_paths, dir_to_protocol, dir_to_channel)
    )

    # 汇总断言
    failures: list[str] = []
    for file_result in results:
        fname = Path(file_result.file_path).name
        if file_result.error:
            failures.append(f"文件比对错误（{fname}）：{file_result.error}")
            continue
        # 范围检查
        if file_result.scope and file_result.scope.missing_from_file:
            failures.append(
                f"范围缺失参数（{fname}）：{file_result.scope.missing_from_file[:5]}…"
            )
        # 数值比对
        for param in file_result.compare_results:
            if param.status == "FAIL":
                failures.append(
                    f"数值比对 FAIL（{fname}）"
                    f" {param.param_key}："
                    f"file={param.file_value} modbus={param.modbus_value}"
                )

    assert not failures, (
        f"[{case.case_id}] 完整三段验证发现 {len(failures)} 个问题：\n"
        + "\n".join(failures[:10])
    )


# ─── D 类：Prefix 超长，保存失败 ──────────────────────────────────────────────

# (case_id, logger_n, long_prefix)
_PREFIX_CASES = [
    ("TestCase_AcuHMI_003_01_case19", 1, "meter2_logger1_12345678910"),
    ("TestCase_AcuHMI_003_02_case11", 2, "meter2_logger1_12345678910"),
    ("TestCase_AcuHMI_003_03_case11", 3, "meter2_logger1_12345678910"),
]


@pytest.mark.lv4
@pytest.mark.parametrize(
    "case_id, logger_n, long_prefix",
    _PREFIX_CASES,
    ids=[c[0] for c in _PREFIX_CASES],
)
def test_prefix_too_long_save_fail(case_id, logger_n, long_prefix, driver):
    """
    D 类：Log File Name Prefix 超过长度限制，保存应失败且显示准确错误信息。
    对应用例：case19 / case11(Logger2) / case11(Logger3)
    注：prefix 字段长度限制请以网关实际行为为准（通常 ≤ 20 字符）。
    """
    dl_page = DataLoggerPage(driver)
    dl_page.navigate_to_logger(logger_n)
    dl_page._set_enable(True)

    # 填写超长前缀
    dl_page._fill(
        dl_page.LOG_FILE_NAME_PREFIX_INPUT,
        long_prefix,
        "Log File Name Prefix (超长)",
    )

    # 点击保存
    dl_page._safe_click(dl_page.SAVE_BTN, "Save")

    # 检测错误提示
    error_msg = _check_save_error(driver)

    assert error_msg, (
        f"[{case_id}] 超长 prefix 保存后，页面未显示错误信息（prefix='{long_prefix}'）"
    )


# ─── E 类：LogFileLength → LogInterval 联动验证 ──────────────────────────────

@pytest.mark.lv1
@pytest.mark.parametrize(
    "logger_n, case_id",
    [
        (1, "TestCase_AcuHMI_003_04_case21"),
        (2, "TestCase_AcuHMI_003_02_case13"),
        (3, "TestCase_AcuHMI_003_03_case13"),
    ],
    ids=["Logger1", "Logger2", "Logger3"],
)
def test_interval_linkage(logger_n, case_id, driver):
    """
    E 类：选不同 LogFileLength 后，LogInterval 下拉可选项必须与预期联动规则一致。
    对应用例：case21 / Logger2-case13 / Logger3-case13

    EXPECTED_INTERVALS 中的 UI 文本须与网关实际下拉选项一致（可按版本调整）。
    """
    dl_page = DataLoggerPage(driver)
    dl_page.navigate_to_logger(logger_n)

    mismatches: list[str] = []
    for length_text, expected in EXPECTED_INTERVALS.items():
        actual = _get_interval_options(dl_page, logger_n, length_text)
        if not actual:
            mismatches.append(
                f"  [{length_text}] 无法读取 LogInterval 下拉选项（请检查定位器）"
            )
            continue

        # 比较可选项集合（忽略顺序差异）
        actual_set   = set(actual)
        expected_set = set(expected)
        missing  = expected_set - actual_set
        extra    = actual_set - expected_set
        if missing or extra:
            mismatches.append(
                f"  [{length_text}] 缺失：{missing}  多余：{extra}"
            )

    assert not mismatches, (
        f"[{case_id}] Logger{logger_n} 联动验证失败：\n" + "\n".join(mismatches)
    )
