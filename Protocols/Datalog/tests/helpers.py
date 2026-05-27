# -*- coding: utf-8 -*-
"""
helpers.py — DataLog 测试共享工具

各用例文件通过 `from helpers import ...` 引用本模块。
包含：协议映射、超时表、验证函数、公共流程封装。
"""
from __future__ import annotations

import asyncio
import csv
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))            # Protocols/Datalog/
sys.path.insert(0, str(Path(__file__).parent.parent.parent))     # Protocols/
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))  # 仓库根

# ─── 协议 → Post Channel 编号 ─────────────────────────────────────────────────
PROTO_TO_CHANNEL: dict[str, int] = {
    "FTP":   1,
    "SFTP":  2,
    "HTTP":  3,
    "HTTPS": 3,
}

# ─── LogFileLength → 等待文件超时（秒）────────────────────────────────────────
LENGTH_TIMEOUT: dict[str, float] = {
    "1 minute":   120,
    "5 minute":   450,
    "10 minute":  720,
    "15 minute":  1080,
    "30 minute":  2100,
    "1 hour":     3780,
    "6 hour":     21960,
    "12 hour":    43560,
    "24 hour":    87000,
    "7 day":      604800 + 300,
    "1 month":    2592300,
}

# ─── LogInterval / LogFileLength UI 文字 → 秒数 ───────────────────────────────
INTERVAL_SECONDS: dict[str, float] = {
    "1 minute":   60,
    "5 minute":   300,
    "10 minute":  600,
    "15 minute":  900,
    "30 minute":  1800,
    "1 hour":     3600,
    "6 hour":     21600,
    "12 hour":    43200,
    "1 day":      86400,
    "24 hour":    86400,
    "7 day":      604800,
    "1 month":    2592000,
}
LENGTH_SECONDS = INTERVAL_SECONDS

# ─── Timestamp Format → 正则表 ────────────────────────────────────────────────
TS_FORMAT_RE: dict[str, list] = {
    "Local Time String": [
        re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}"),
        re.compile(r"^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}"),
    ],
    "UTC Seconds": [
        re.compile(r"^\d{10,13}$"),
    ],
    "ISO8601 Format": [
        re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}"),
    ],
}

# ─── Timestamp Format → datetime 解析器 ───────────────────────────────────────
TS_PARSERS: dict[str, list] = {
    "Local Time String": [
        lambda s: datetime.strptime(s, "%Y-%m-%d %H:%M:%S"),
        lambda s: datetime.strptime(s, "%m/%d/%Y %H:%M:%S"),
    ],
    "UTC Seconds": [
        lambda s: datetime.utcfromtimestamp(float(s)),
    ],
    "ISO8601 Format": [
        lambda s: datetime.fromisoformat(s.replace("Z", "+00:00")),
        lambda s: datetime.strptime(s[:19], "%Y-%m-%dT%H:%M:%S"),
    ],
}

# ─── LogFileLength → LogInterval 联动期望表 ───────────────────────────────────
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


# ─── 文件收集 ─────────────────────────────────────────────────────────────────

def collect_files(dirs: list[str], exts=(".json", ".csv")) -> list[str]:
    result = []
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for fn in os.listdir(d):
            if fn.lower().endswith(exts):
                result.append(os.path.join(d, fn))
    return result


# ─── 文件名格式验证 ────────────────────────────────────────────────────────────

def verify_filename_format(filepath: str, prefix: str, name_fmt: str) -> str:
    """
    验证文件名是否符合 Log File Name Prefix + Log File Name Format 配置。
    返回错误描述；无误则返回空字符串。
    """
    stem = Path(filepath).stem
    if not stem.startswith(prefix):
        return f"前缀不匹配：期望 '{prefix}'，实际 '{stem}'"
    rest = stem[len(prefix):]
    if name_fmt == "UTC Timestamp":
        if not re.match(r"^[_-]?\d{14}$", rest):
            return f"UTC Timestamp 格式不符（期望 prefix+YYYYMMDDHHmmss），实际后缀：'{rest}'"
    elif name_fmt == "Time Interval Format":
        if not re.match(r"^[_-]?\d{14}[_-]\d{14}$", rest):
            return f"Time Interval Format 不符（期望 prefix+start+end），实际后缀：'{rest}'"
    return ""


# ─── 文件内容时间戳读取 ────────────────────────────────────────────────────────

def read_timestamps(filepath: str, file_format: str) -> list[str]:
    """CSV 取第一列；JSON 取 time/Time/timestamp 等字段。"""
    timestamps: list[str] = []
    try:
        if file_format.lower() == "csv":
            with open(filepath, encoding="utf-8", errors="replace") as f:
                reader = csv.reader(f)
                next(reader, None)
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


def _parse_ts_to_dt(timestamps: list[str], ts_fmt: str) -> list[datetime]:
    parsers = TS_PARSERS.get(ts_fmt, [])
    result = []
    for s in timestamps:
        for p in parsers:
            try:
                dt = p(s)
                if dt.tzinfo is not None:
                    dt = dt.replace(tzinfo=None)
                result.append(dt)
                break
            except Exception:
                continue
    return result


# ─── 内容验证 ─────────────────────────────────────────────────────────────────

def verify_timestamp_format(timestamps: list[str], ts_fmt: str) -> str:
    if not timestamps:
        return "文件无数据行，无法验证时间戳格式"
    patterns = TS_FORMAT_RE.get(ts_fmt, [])
    if not patterns:
        return ""
    for ts in timestamps[:3]:
        if not any(p.match(ts) for p in patterns):
            return f"格式不符（期望 {ts_fmt}），实际：'{ts}'"
    return ""


def verify_log_interval(timestamps: list[str], ts_fmt: str,
                         interval_str: str) -> str:
    if len(timestamps) < 2:
        return ""
    expected = INTERVAL_SECONDS.get(interval_str, 60)
    dts = _parse_ts_to_dt(timestamps, ts_fmt)
    if len(dts) < 2:
        return ""
    tol = max(expected * 0.1, 30)
    bad = []
    for i in range(min(len(dts) - 1, 5)):
        diff = abs((dts[i + 1] - dts[i]).total_seconds())
        if abs(diff - expected) > tol:
            bad.append(f"{diff:.0f}s")
    if bad:
        return f"期望 {expected:.0f}s±{tol:.0f}s，实际间隔：{bad}"
    return ""


def verify_file_length(timestamps: list[str], ts_fmt: str,
                        length_str: str, interval_str: str) -> str:
    if not timestamps:
        return ""
    exp_len = LENGTH_SECONDS.get(length_str, 60)
    exp_ivl = INTERVAL_SECONDS.get(interval_str, 60)
    exp_rows = round(exp_len / exp_ivl)
    if abs(len(timestamps) - exp_rows) > 2:
        return (f"行数不符（期望约 {exp_rows} 行 [{length_str}/{interval_str}]，"
                f"实际 {len(timestamps)} 行）")
    dts = _parse_ts_to_dt([timestamps[0], timestamps[-1]], ts_fmt)
    if len(dts) == 2:
        span = abs((dts[1] - dts[0]).total_seconds())
        if span > exp_len * 1.3 + 30:
            return f"时间跨度超出预期（≤{exp_len * 1.3 + 30:.0f}s，实际 {span:.0f}s）"
    return ""


# ─── 完整三段验证（范围 + Modbus 数值） ───────────────────────────────────────

def run_full_3stage_verify(case_id: str, file_paths: list[str],
                            pool, protocol: str):
    from datalog_server_verifier import _select_latest_per_device, verify_files

    fps = _select_latest_per_device(list(file_paths))
    tdir = os.path.normpath(pool[protocol].data_dir)
    d2p = {tdir: protocol.upper()}
    d2c = {tdir: PROTO_TO_CHANNEL[protocol]}

    results = asyncio.run(verify_files(fps, d2p, d2c))

    failures: list[str] = []
    for fr in results:
        fname = Path(fr.file_path).name
        if fr.error:
            failures.append(f"比对错误（{fname}）：{fr.error}")
            continue
        if fr.scope and fr.scope.missing:
            failures.append(f"范围缺失（{fname}）：{fr.scope.missing[:5]}…")
        for p in fr.compare_results:
            if p.status == "FAIL":
                failures.append(
                    f"数值 FAIL（{fname}）{p.param_key}："
                    f"file={p.file_value} modbus={p.modbus_value}"
                )
    assert not failures, (
        f"[{case_id}] 三段验证发现 {len(failures)} 个问题：\n"
        + "\n".join(failures[:10])
    )


# ─── C 类推送验证公共流程 ──────────────────────────────────────────────────────

def run_push_case(
    case_id: str,
    logger_n: int,
    protocol: str,
    file_format: str,
    file_length: str,
    timestamp_fmt: str,
    name_fmt: str,
    prefix: str,
    interval: str,
    pool,
    driver,
    full_verify: bool = False,
):
    """
    C 类推送验证完整流程：配置 Logger → 等待文件 → 禁用 → 验证。
    验证项：文件存在、扩展名、文件名格式、时间戳格式、行间隔、文件时长。
    full_verify=True 时额外执行范围 + Modbus 数值比对。
    """
    from datalog_page import DataLoggerConfig, DataLoggerPage
    from datalog_server_verifier import wait_for_files

    channel_n = PROTO_TO_CHANNEL[protocol]
    dl_page = DataLoggerPage(driver)
    dl_page.configure_logger(logger_n, DataLoggerConfig(
        channel_index=channel_n,
        enabled=True,
        log_file_format=file_format,
        log_file_length=file_length,
        timestamp_format=timestamp_fmt,
        log_file_name_format=name_fmt,
        log_file_name_prefix=prefix,
        log_interval=interval,
        device_names=[],
    ))

    target_dir = os.path.normpath(pool[protocol].data_dir)
    timeout = LENGTH_TIMEOUT.get(file_length, 120)
    file_paths = wait_for_files([target_dir], timeout=timeout)

    if file_paths:
        dl_page.disable_logger(logger_n)

    assert file_paths, (
        f"[{case_id}] 超时 {timeout}s 内未收到文件 "
        f"（Logger{logger_n} {protocol} {file_format} {file_length}）"
    )

    expected_ext = f".{file_format.lower()}"
    for fp in file_paths:
        assert Path(fp).suffix.lower() == expected_ext, (
            f"[{case_id}] 扩展名不匹配：{Path(fp).name}，期望 {expected_ext}"
        )
        err = verify_filename_format(fp, prefix, name_fmt)
        assert not err, f"[{case_id}] 文件名格式错误（{Path(fp).name}）：{err}"

        ts = read_timestamps(fp, file_format)
        ts_err = verify_timestamp_format(ts, timestamp_fmt)
        assert not ts_err, f"[{case_id}] Timestamp Format 不符（{Path(fp).name}）：{ts_err}"

        iv_err = verify_log_interval(ts, timestamp_fmt, interval)
        assert not iv_err, f"[{case_id}] Log Interval 错误（{Path(fp).name}）：{iv_err}"

        ln_err = verify_file_length(ts, timestamp_fmt, file_length, interval)
        assert not ln_err, f"[{case_id}] Log File Length 错误（{Path(fp).name}）：{ln_err}"

    if full_verify:
        run_full_3stage_verify(case_id, file_paths, pool, protocol)


# ─── B 类 PostChannel=None 配置 ───────────────────────────────────────────────

def configure_logger_none_channel(dl_page, logger_n: int, file_format: str,
                                   file_length: str, timestamp_fmt: str,
                                   name_fmt: str, prefix: str, interval: str):
    """配置 Data Logger PostChannel=None（不推送到远端服务器）。"""
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
    dl_page._select_devices([])
    dl_page._safe_click(dl_page.SAVE_BTN, "Save")
    time.sleep(2)


# ─── D 类错误提示检测 ─────────────────────────────────────────────────────────

def check_save_error(driver) -> str:
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
        for el in driver.find_elements(By.XPATH, xpath):
            txt = el.text.strip()
            if txt:
                return txt
    return ""


# ─── E 类联动选项读取 ─────────────────────────────────────────────────────────

def get_interval_options(dl_page, logger_n: int, file_length: str) -> list[str]:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    dl_page.navigate_to_logger(logger_n)
    dl_page._select_el_by_text(dl_page.LOG_FILE_LENGTH_SELECT, file_length, "Log File Length")
    time.sleep(0.8)
    options: list[str] = []
    try:
        container = dl_page.driver.find_element(*dl_page.LOG_INTERVAL_SELECT)
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
            dl_page.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(0.2)
        except Exception:
            pass
    return options
