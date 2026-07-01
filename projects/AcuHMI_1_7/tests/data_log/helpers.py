# -*- coding: utf-8 -*-
"""
helpers.py — DataLog 测试共享工具（Playwright 版）
"""
from __future__ import annotations

import asyncio
import concurrent.futures
import csv
import json
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)

_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent.parent.parent.parent  # AcuHMI-1-7/
sys.path.insert(0, str(_HERE))
sys.path.insert(0, str(_PROJECT_ROOT))

# ─── 协议 → Post Channel 编号 ─────────────────────────────────────────────────
PROTO_TO_CHANNEL: dict[str, int] = {
    "FTP":   1,
    "SFTP":  2,
    "HTTP":  3,
    "HTTPS": 3,
}

# ─── 内部 key → Web UI 显示文本 ───────────────────────────────────────────────
_UI_TEXT: dict[str, str] = {
    # Rapid Logger 秒级间隔：若 UI 显示不同形式请在此调整
    "1 second":             "1 second",
    "2 seconds":            "2 seconds",
    "5 seconds":            "5 seconds",
    "10 seconds":           "10 seconds",
    "20 seconds":           "20 seconds",
    "30 seconds":           "30 seconds",
    # 分钟及以上
    "5 minute":             "5 minutes",
    "10 minute":            "10 minutes",
    "15 minute":            "15 minutes",
    "30 minute":            "30 minutes",
    "6 hour":               "6 hours",
    "12 hour":              "12 hours",
    "24 hour":              "1 day",
    "7 day":                "7 days",
    "Time Interval Format": "Time interval Format",
}

def _ui(key: str) -> str:
    return _UI_TEXT.get(key, key)


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
    # Rapid Logger 秒级间隔（内部 key → 秒数）
    "1 second":   1,
    "2 seconds":  2,
    "5 seconds":  5,
    "10 seconds": 10,
    "20 seconds": 20,
    "30 seconds": 30,
    # 分钟及以上
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
        re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?"),
        re.compile(r"^\d{2}/\d{2}/\d{4} \d{2}:\d{2}(:\d{2})?"),
    ],
    "UTC Seconds": [
        re.compile(r"^\d{10,13}$"),
    ],
    "ISO8601 Format": [
        re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?"),
    ],
}

# ─── Timestamp Format → datetime 解析器 ───────────────────────────────────────
TS_PARSERS: dict[str, list] = {
    "Local Time String": [
        lambda s: datetime.strptime(s, "%Y-%m-%d %H:%M:%S"),
        lambda s: datetime.strptime(s, "%Y-%m-%d %H:%M"),
        lambda s: datetime.strptime(s, "%m/%d/%Y %H:%M:%S"),
        lambda s: datetime.strptime(s, "%m/%d/%Y %H:%M"),
    ],
    "UTC Seconds": [
        lambda s: datetime.utcfromtimestamp(float(s)),
    ],
    "ISO8601 Format": [
        lambda s: datetime.fromisoformat(s.replace("Z", "+00:00")),
        lambda s: datetime.strptime(s[:19], "%Y-%m-%dT%H:%M:%S"),
        lambda s: datetime.strptime(s[:16], "%Y-%m-%dT%H:%M"),
    ],
}

# ─── LogFileLength → LogInterval 联动期望表 ───────────────────────────────────
EXPECTED_INTERVALS: dict[str, list[str]] = {
    "1 minute":  ["1 minute"],
    "5 minute":  ["1 minute", "5 minutes"],
    "10 minute": ["1 minute", "5 minutes", "10 minutes"],
    "15 minute": ["1 minute", "5 minutes", "10 minutes", "15 minutes"],
    "30 minute": ["1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes"],
    "1 hour":    ["1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour"],
    "6 hour":    ["1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour", "6 hours"],
    "12 hour":   ["1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour", "6 hours", "12 hours"],
    "24 hour":   ["5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour", "6 hours", "12 hours", "1 day"],
    "7 day":     ["15 minutes", "30 minutes", "1 hour", "6 hours", "12 hours", "1 day", "7 days"],
    "1 month":   ["1 hour", "6 hours", "12 hours", "1 day", "7 days", "1 month"],
}


# ─── 文件收集 ─────────────────────────────────────────────────────────────────

def collect_files(dirs: list[str], exts=(".json", ".csv"),
                  logger_n: int = None) -> list[str]:
    """收集目录中的数据文件；logger_n 非空时只返回 Logger{n}- 开头的文件。"""
    prefix = f"Logger{logger_n}-" if logger_n is not None else None
    result = []
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for fn in os.listdir(d):
            if fn.lower().endswith(exts):
                if prefix and not fn.startswith(prefix):
                    continue
                result.append(os.path.join(d, fn))
    return result


# ─── 文件名格式验证 ────────────────────────────────────────────────────────────

def verify_filename_format(filepath: str, prefix: str, name_fmt: str) -> str:
    stem = Path(filepath).stem
    if not stem.startswith(prefix):
        return f"前缀不匹配：期望 '{prefix}'，实际 '{stem}'"
    rest = stem[len(prefix):]
    if name_fmt == "UTC Timestamp":
        if not re.search(r"-\d{9,11}-\d+\w+$", rest):
            return f"UTC Timestamp 格式不符，实际后缀：'{rest}'"
    elif name_fmt == "Time Interval Format":
        if not re.search(r"-\d{4}-\d{2}-\d{2}T\d{2}[:\-]\d{2}[:\-]\d{2}", rest):
            return f"Time Interval Format 不符（未含 ISO8601 时间段），实际后缀：'{rest}'"
    return ""


# ─── 文件内容时间戳读取 ────────────────────────────────────────────────────────

def read_timestamps(filepath: str, file_format: str) -> list[str]:
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
            _TS_KEYS = ("time", "Time", "timestamp", "Timestamp", "datetime", "DateTime")
            if isinstance(data, list):
                for rec in data:
                    if isinstance(rec, dict):
                        for key in _TS_KEYS:
                            if key in rec:
                                timestamps.append(str(rec[key]))
                                break
            elif isinstance(data, dict):
                for key in _TS_KEYS:
                    if key in data:
                        val = data[key]
                        if isinstance(val, list):
                            timestamps.extend(str(v) for v in val)
                        else:
                            timestamps.append(str(val))
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


# ─── JSON 结构完整性检查 ──────────────────────────────────────────────────────

def _verify_json_structure(file_path: str) -> list[str]:
    issues: list[str] = []
    try:
        with open(file_path, encoding="utf-8", errors="replace") as f:
            data = json.load(f)
    except Exception as e:
        return [f"JSON 解析失败：{e}"]

    gw = data.get("gateway", {})
    for field in ("model", "serial"):
        if not gw.get(field):
            issues.append(f"gateway.{field} 为空")

    dev = data.get("device", {})
    for field in ("name", "model", "serial"):
        if not dev.get(field):
            issues.append(f"device.{field} 为空")
    if dev.get("online") is None:
        issues.append("device.online 字段缺失")

    readings = dev.get("readings", [])
    if not readings:
        issues.append("device.readings 为空，无参数数据")

    return issues


# ─── 完整三段验证（范围 + 单位 + Modbus 数值） ────────────────────────────────

def run_full_3stage_verify(case_id: str, file_paths: list[str],
                            pool, protocol: str):
    from datalog_server_verifier import _select_latest_per_device, verify_files

    fps = _select_latest_per_device(list(file_paths))
    tdir = os.path.normpath(pool[protocol].data_dir)
    d2p = {tdir: protocol.upper()}
    d2c = {tdir: PROTO_TO_CHANNEL[protocol]}

    # Playwright 占用主线程事件循环，asyncio.run() 不能嵌套调用；
    # 在独立线程中执行，避免 "cannot be called from a running event loop" 错误
    def _run():
        return asyncio.run(verify_files(fps, d2p, d2c))

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as _executor:
        results = _executor.submit(_run).result(timeout=180)

    failures: list[str] = []

    for fp in fps:
        if fp.lower().endswith(".json"):
            fname = Path(fp).name
            for issue in _verify_json_structure(fp):
                failures.append(f"【JSON结构】{fname}：{issue}")

    for fr in results:
        fname = Path(fr.file_path).name
        if fr.error:
            failures.append(f"【设备 Modbus 连接/读取失败】{fname}：{fr.error}")
            continue
        if fr.scope and fr.scope.missing_from_file:
            failures.append(
                f"【范围缺失】{fname}：{fr.scope.missing_from_file[:5]}…"
            )
        for u in fr.unit_results:
            if u.status == "FAIL":
                failures.append(
                    f"【单位不符】{fname}  {u.param_key}："
                    f"文件={u.file_unit}  模板={u.tmpl_unit}"
                )
        for p in fr.compare_results:
            if p.status == "FAIL":
                failures.append(
                    f"【数值不符】{fname}  {p.param_key}："
                    f"文件={p.file_value}  Modbus={p.modbus_value}"
                )
    assert not failures, (
        f"[{case_id}] 三段验证发现 {len(failures)} 个问题：\n"
        + "\n".join(failures[:20])
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
    import config
    from datalog_page import DataLoggerConfig, DataLoggerPage
    from datalog_server_verifier import wait_for_files

    # Devices Selection 仅勾选「在线 Modbus 设备」：DISCOVERED_DEVICES 由 conftest 经
    # /api/device/list/modbus 动态发现（只含 deviceType 2/3），天然排除虚拟设备；再按
    # online 过滤离线设备。为空时（发现失败）回退为全选，并告警。
    allowed_devices = [
        d.name for d in getattr(config, "DISCOVERED_DEVICES", [])
        if getattr(d, "online", False)
    ]
    if not allowed_devices:
        log.warning(
            "[%s] 未发现在线 Modbus 设备，Devices Selection 回退为全选"
            "（可能含虚拟/离线设备，比对或误判范围缺失）", case_id,
        )

    channel_n = PROTO_TO_CHANNEL[protocol]
    dl_page = DataLoggerPage(driver)
    dl_page.configure_logger(logger_n, DataLoggerConfig(
        channel_index=channel_n,
        enabled=True,
        log_file_format=file_format,
        log_file_length=_ui(file_length),
        timestamp_format=_ui(timestamp_fmt),
        log_file_name_format=_ui(name_fmt),
        log_file_name_prefix=prefix,
        log_interval=_ui(interval),
        device_names=allowed_devices,
    ))

    target_dir = os.path.normpath(pool[protocol].data_dir)
    timeout = LENGTH_TIMEOUT.get(file_length, 120)
    file_paths = wait_for_files([target_dir], timeout=timeout)

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
        effective_ts_fmt = "UTC Seconds" if file_format.lower() == "json" else timestamp_fmt
        ts_err = verify_timestamp_format(ts, effective_ts_fmt)
        assert not ts_err, f"[{case_id}] Timestamp Format 不符（{Path(fp).name}）：{ts_err}"

        iv_err = verify_log_interval(ts, effective_ts_fmt, interval)
        assert not iv_err, f"[{case_id}] Log Interval 错误（{Path(fp).name}）：{iv_err}"

        ln_err = verify_file_length(ts, effective_ts_fmt, file_length, interval)
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
    # Post Channel=None：hover 触发 mouseenter → Vue showClose=true → clear 按钮可见
    try:
        container = dl_page.page.locator(dl_page._POST_CHANNEL_SELECT).first
        container.wait_for(state="visible", timeout=8000)
        wrapper = container.locator(".el-select__wrapper")
        if wrapper.count() > 0:
            wrapper.first.hover()
        else:
            container.hover()
        dl_page.page.wait_for_timeout(1000)  # 等待 Vue 响应式更新
        cleared = False
        for sel in [
            ".el-select__clear",
            "[class*='el-select__clear']",
            ".el-select__suffix .el-icon:last-child",
        ]:
            clear_btn = container.locator(sel)
            if clear_btn.count() > 0:
                try:
                    clear_btn.first.wait_for(state="visible", timeout=500)
                    clear_btn.first.evaluate("el => el.click()")
                    dl_page.page.wait_for_timeout(300)
                    cleared = True
                    break
                except Exception:
                    continue
        if not cleared:
            log.warning("Post Channel clear 按钮未找到（可能已为空），继续后续配置")
    except Exception as e:
        log.warning("清空 Post Channel 失败：%s", e)
    # Log File Format 必须先选，JSON 模式下 Timestamp Format 字段会隐藏
    dl_page._select_el_by_text(dl_page._LOG_FILE_FORMAT_SELECT, file_format, "Log File Format")
    if file_format.lower() != "json":
        dl_page._set_timestamp_format(timestamp_fmt)
    dl_page._set_log_file_name_format(name_fmt)
    if prefix:
        dl_page._fill(dl_page._LOG_FILE_NAME_PREFIX_INPUT, prefix, "Log File Name Prefix")
    dl_page._select_el_by_text(dl_page._LOG_FILE_LENGTH_SELECT, file_length, "Log File Length")
    dl_page._select_el_by_text(dl_page._LOG_INTERVAL_SELECT, interval, "Log Interval")
    dl_page._select_devices([])
    dl_page._safe_click(dl_page._SAVE_BTN, "Save")
    dl_page.page.wait_for_timeout(2000)


# ─── D 类错误提示检测 ─────────────────────────────────────────────────────────

def check_save_error(page) -> str:
    """检测 Save 操作后是否出现错误提示，返回错误文本；无误则返回空字符串。"""
    page.wait_for_timeout(1500)
    for xpath in [
        "//div[contains(@class,'el-message') and contains(@class,'el-message--error')]",
        "//div[contains(@class,'el-form-item__error')]",
        "//div[contains(@class,'toast') and contains(translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'error')]",
        "//span[contains(@class,'error')]",
        "//*[contains(@class,'is-error')]//div[contains(@class,'el-form-item__error')]",
    ]:
        els = page.locator(f"xpath={xpath}").all()
        for el in els:
            txt = (el.text_content() or "").strip()
            if txt:
                return txt
    return ""


# ─── E 类联动选项读取 ─────────────────────────────────────────────────────────

def get_interval_options(dl_page, logger_n: int, file_length: str) -> list[str]:
    """读取指定 Log File Length 下可用的 Log Interval 下拉选项。"""
    dl_page.navigate_to_logger(logger_n)
    dl_page._set_enable(True)
    dl_page._select_el_by_text(dl_page._LOG_FILE_LENGTH_SELECT, _ui(file_length), "Log File Length")
    dl_page.page.wait_for_timeout(800)
    options: list[str] = []
    try:
        container = dl_page.page.locator(dl_page._LOG_INTERVAL_SELECT).first
        inner = container.locator(".el-select__wrapper")
        if inner.count() > 0:
            inner.first.evaluate("el => el.click()")
        else:
            container.evaluate("el => el.click()")
        dl_page.page.wait_for_timeout(800)
        # 只从当前可见的弹层中读取，避免获取到其他已关闭下拉框的项目
        all_items = dl_page.page.locator(
            "xpath=//li[contains(@class,'el-select-dropdown__item')"
            " and not(contains(@class,'is-disabled'))"
            " and not(contains(@class,'disabled'))]"
        ).all()
        options = [
            el.text_content().strip()
            for el in all_items
            if el.text_content().strip() and el.is_visible()
        ]
    except Exception:
        pass
    finally:
        try:
            dl_page.page.keyboard.press("Escape")
            dl_page.page.wait_for_timeout(200)
        except Exception:
            pass
    return options


# ─── Rapid Logger 联动期望表 ──────────────────────────────────────────────────
# 键为内部 key，值为 Rapid Logger 中对应 Log File Length 下的 Log Interval UI 选项列表
# 注意：秒级选项的实际 UI 文字可能为 "1 s"/"2 s" 等，若不符请在此调整
RAPID_EXPECTED_INTERVALS: dict[str, list[str]] = {
    "1 minute":  ["1 second", "2 seconds", "5 seconds", "10 seconds",
                  "20 seconds", "30 seconds", "1 minute"],
    "5 minute":  ["1 second", "2 seconds", "5 seconds", "10 seconds",
                  "20 seconds", "30 seconds", "1 minute", "5 minutes"],
    "10 minute": ["1 second", "2 seconds", "5 seconds", "10 seconds",
                  "20 seconds", "30 seconds", "1 minute", "5 minutes", "10 minutes"],
    "15 minute": ["1 second", "2 seconds", "5 seconds", "10 seconds",
                  "20 seconds", "30 seconds", "1 minute", "5 minutes", "10 minutes", "15 minutes"],
    "30 minute": ["2 seconds", "5 seconds", "10 seconds", "20 seconds",
                  "30 seconds", "1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes"],
    "1 hour":    ["5 seconds", "10 seconds", "20 seconds", "30 seconds",
                  "1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour"],
    "6 hour":    ["20 seconds", "30 seconds",
                  "1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour", "6 hours"],
    "12 hour":   ["1 minute", "5 minutes", "10 minutes", "15 minutes",
                  "30 minutes", "1 hour", "6 hours", "12 hours"],
    "24 hour":   ["5 minutes", "10 minutes", "15 minutes", "30 minutes",
                  "1 hour", "6 hours", "12 hours", "1 day"],
    "7 day":     ["15 minutes", "30 minutes", "1 hour", "6 hours", "12 hours", "1 day", "7 days"],
    "1 month":   ["1 hour", "6 hours", "12 hours", "1 day", "7 days", "1 month"],
}


def get_rapid_interval_options(rl_page, file_length: str) -> list[str]:
    """读取 Rapid Logger 在指定 Log File Length 下可用的 Log Interval 选项。"""
    return rl_page.get_log_interval_options(_ui(file_length))


# ─── Rapid Logger PostChannel=None 配置 ───────────────────────────────────────

def configure_rapid_logger_none_channel(
    rl_page,
    file_format: str,
    file_length: str,
    timestamp_fmt: str,
    name_fmt: str,
    prefix: str,
    interval: str,
):
    """配置 Rapid Logger 且 Post Channel=None（不推送到远端服务器）。"""
    rl_page.navigate_to_rapid_logger()
    rl_page._set_enable(True)
    # 清空 Post Channel
    try:
        container = rl_page.page.locator(rl_page._POST_CHANNEL_SELECT).first
        container.wait_for(state="visible", timeout=8000)
        wrapper = container.locator(".el-select__wrapper")
        if wrapper.count() > 0:
            wrapper.first.hover()
        else:
            container.hover()
        rl_page.page.wait_for_timeout(1000)
        cleared = False
        for sel in [".el-select__clear", "[class*='el-select__clear']",
                     ".el-select__suffix .el-icon:last-child"]:
            btn = container.locator(sel)
            if btn.count() > 0:
                try:
                    btn.first.wait_for(state="visible", timeout=500)
                    btn.first.evaluate("el => el.click()")
                    rl_page.page.wait_for_timeout(300)
                    cleared = True
                    break
                except Exception:
                    continue
        if not cleared:
            log.warning("Rapid Logger Post Channel clear 按钮未找到，可能已为空")
    except Exception as e:
        log.warning("Rapid Logger 清空 Post Channel 失败：%s", e)

    rl_page._select_el_by_text(rl_page._LOG_FILE_FORMAT_SELECT, file_format, "Log File Format")
    if file_format.lower() != "json":
        rl_page._set_timestamp_format(timestamp_fmt)
    rl_page._set_log_file_name_format(name_fmt)
    if prefix:
        rl_page._fill(rl_page._LOG_FILE_NAME_PREFIX_INPUT, prefix, "Log File Name Prefix")
    rl_page._select_el_by_text(rl_page._LOG_FILE_LENGTH_SELECT, file_length, "Log File Length")
    rl_page._select_el_by_text(rl_page._LOG_INTERVAL_SELECT, interval, "Log Interval")
    rl_page._safe_click(rl_page._SAVE_BTN, "Save")
    rl_page.page.wait_for_timeout(2000)


# ─── Rapid Logger 正向推送公共流程 ───────────────────────────────────────────

def run_rapid_push_case(
    case_id: str,
    protocol: str,
    file_format: str,
    file_length: str,
    ts_fmt: str,
    name_fmt: str,
    prefix: str,
    interval: str,
    pool,
    driver,
    full_verify: bool = False,
):
    """Rapid Logger 正向推送验证（无 logger_n，对应 run_push_case 的 Rapid Logger 版本）。"""
    from datalog_page import RapidLoggerPage
    from datalog_server_verifier import wait_for_files

    channel_n = PROTO_TO_CHANNEL[protocol]
    rl_page = RapidLoggerPage(driver)
    rl_page.configure_rapid_logger(
        channel_index=channel_n,
        file_format=file_format,
        file_length=_ui(file_length),
        timestamp_fmt=_ui(ts_fmt),
        name_fmt=_ui(name_fmt),
        prefix=prefix,
        log_interval=_ui(interval),
        enabled=True,
    )

    target_dir = os.path.normpath(pool[protocol].data_dir)
    timeout = LENGTH_TIMEOUT.get(file_length, 120)
    file_paths = wait_for_files([target_dir], timeout=timeout)

    rl_page.disable_rapid_logger()

    assert file_paths, (
        f"[{case_id}] 超时 {timeout}s 内未收到文件 "
        f"（Rapid Logger {protocol} {file_format} {file_length}）"
    )

    expected_ext = f".{file_format.lower()}"
    for fp in file_paths:
        assert Path(fp).suffix.lower() == expected_ext, (
            f"[{case_id}] 扩展名不匹配：{Path(fp).name}，期望 {expected_ext}"
        )
        err = verify_filename_format(fp, prefix, name_fmt)
        assert not err, f"[{case_id}] 文件名格式错误（{Path(fp).name}）：{err}"

        ts = read_timestamps(fp, file_format)
        effective_ts_fmt = "UTC Seconds" if file_format.lower() == "json" else ts_fmt
        ts_err = verify_timestamp_format(ts, effective_ts_fmt)
        assert not ts_err, f"[{case_id}] Timestamp Format 不符（{Path(fp).name}）：{ts_err}"

        iv_err = verify_log_interval(ts, effective_ts_fmt, interval)
        assert not iv_err, f"[{case_id}] Log Interval 错误（{Path(fp).name}）：{iv_err}"

        ln_err = verify_file_length(ts, effective_ts_fmt, file_length, interval)
        assert not ln_err, f"[{case_id}] Log File Length 错误（{Path(fp).name}）：{ln_err}"

    if full_verify:
        run_full_3stage_verify(case_id, file_paths, pool, protocol)


# ─── Post Historical Data 推送公共流程 ────────────────────────────────────────

def run_post_historical_case(
    case_id: str,
    protocol: str,
    file_format: str,
    file_length: str,
    timestamp_fmt: str,
    name_fmt: str,
    prefix: str,
    interval: str,
    pool,
    driver,
    expect_error: bool = False,
):
    """Post Historical Data 推送验证流程。"""
    from datalog_page import PostHistoricalDataPage
    from datalog_server_verifier import wait_for_files

    channel_n = PROTO_TO_CHANNEL[protocol]
    ph_page = PostHistoricalDataPage(driver)

    error_msg = ph_page.post(
        channel_index=channel_n,
        file_format=file_format,
        file_length=file_length,
        timestamp_fmt=timestamp_fmt,
        name_fmt=name_fmt,
        prefix=prefix,
        interval=interval,
    )

    if expect_error:
        assert error_msg, (
            f"[{case_id}] 期望出现错误提示，但页面未显示任何错误（prefix='{prefix}'）"
        )
        return

    assert not error_msg, f"[{case_id}] Post Historical Data 意外错误：{error_msg}"

    ph_page.wait_for_completion(timeout_ms=300000)

    target_dir = os.path.normpath(pool[protocol].data_dir)
    file_paths = wait_for_files([target_dir], timeout=120)

    assert file_paths, (
        f"[{case_id}] Post Historical Data 完成后 120s 内未在 {protocol} 目录收到文件"
    )

    expected_ext = f".{file_format.lower()}"
    for fp in file_paths:
        assert Path(fp).suffix.lower() == expected_ext, (
            f"[{case_id}] 扩展名不匹配：{Path(fp).name}，期望 {expected_ext}"
        )
        err = verify_filename_format(fp, prefix, name_fmt)
        assert not err, f"[{case_id}] 文件名格式错误（{Path(fp).name}）：{err}"

        ts = read_timestamps(fp, file_format)
        effective_ts_fmt = "UTC Seconds" if file_format.lower() == "json" else timestamp_fmt
        ts_err = verify_timestamp_format(ts, effective_ts_fmt)
        assert not ts_err, f"[{case_id}] Timestamp Format 不符（{Path(fp).name}）：{ts_err}"
