# -*- coding: utf-8 -*-
"""
datalog_server_verifier.py — DataLog Post Channel 推送端到端验证

设计原则：
  - Post Channel 1 / 2 / 3 各自可独立选择 FTP / SFTP / HTTP / HTTPS 四种协议之一
  - 本机维护一个「协议池」（四协议各一个服务器实例），按 channel_configs 中
    实际用到的协议按需启动，未用到的协议不启动
  - 通过 channel_to_post_config() 把协议名翻译成填好本机服务器信息的 PostChannelConfig

流程：
  1. 按需启动 FTP / SFTP / HTTP / HTTPS 服务器（仅启动 channel_configs 中用到的协议）
  2. Selenium 登录网关，配置 Post Channel 1/2/3（每个可以是任意协议）
  3. 配置 Data Logger 1/2/3（推送间隔 / 格式 / 关联 Channel / 设备勾选 / 全选参数）
  4. 等待网关推送文件到各服务器对应目录
  5. 对收到的每个文件执行三段式比对（范围 + 单位（JSON）+ Modbus 数值）
  6. 生成汇总 HTML 报告

用法：
  python Protocols/Datalog/datalog_server_verifier.py
  python Protocols/Datalog/datalog_server_verifier.py --no-webdriver
  python Protocols/Datalog/datalog_server_verifier.py --timeout 600
"""
from __future__ import annotations

import argparse
import asyncio
import html
import logging
import os
import platform
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))        # Protocols/
sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # 仓库根目录（comm/ 在此）

import config
from datalog_comparator import (
    DatalogCompareResult,
    DatalogScopeReport,
    DatalogUnitResult,
    run_datalog_comparison,
    summary,
)
from datalog_page import (
    DataLoggerConfig, DataLoggerPage,
    DataLogParamConfig, DataLogParamConfigPage,
    PostChannelConfig, PostChannelPage,
)
from servers import start_ftp_server, start_http_server, start_sftp_server

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# 服务器信息
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ServerInfo:
    """描述一个本地服务器实例。"""
    protocol: str        # "FTP" | "SFTP" | "HTTP" | "HTTPS"
    host: str            # 本机 IP（网关需能访问）
    port: int
    data_dir: str        # 接收文件的本地目录
    username: str = "datalog"
    password: str = "datalog123"
    remote_path: str = "/"
    ssl_certfile: str = ""   # HTTPS 专用
    ssl_keyfile: str = ""    # HTTPS 专用

    def to_post_channel_config(self) -> PostChannelConfig:
        """转换为填好本机服务器信息的 PostChannelConfig。"""
        return PostChannelConfig(
            protocol=self.protocol,   # "FTP"|"SFTP"|"HTTP"|"HTTPS"
            host=self.host,
            port=self.port,
            username=self.username,
            password=self.password,
            # HTTP/HTTPS 默认值：不固定文件名、不需要认证、有 Header
            post_name_fixed=False,
            auth_required=False,
            include_header=True,
            meter_id="meter_001",
        )


# ─────────────────────────────────────────────────────────────────────────────
# 协议池：四协议各一个本地服务器，从 config.py 读取参数
# ─────────────────────────────────────────────────────────────────────────────

def build_protocol_pool() -> dict[str, ServerInfo]:
    """
    构建四协议服务器池（FTP / SFTP / HTTP / HTTPS）。
    所有参数来自 config.py 的 DATALOG_* 配置项。
    """
    base = config.DATALOG_DATA_DIR
    return {
        "FTP": ServerInfo(
            protocol="FTP",
            host=config.DATALOG_SERVER_HOST,
            port=config.DATALOG_FTP_PORT,
            data_dir=os.path.join(base, "ftp"),
            username=config.DATALOG_FTP_USER,
            password=config.DATALOG_FTP_PASS,
        ),
        "SFTP": ServerInfo(
            protocol="SFTP",
            host=config.DATALOG_SERVER_HOST,
            port=config.DATALOG_SFTP_PORT,
            data_dir=os.path.join(base, "sftp"),
            username=config.DATALOG_SFTP_USER,
            password=config.DATALOG_SFTP_PASS,
        ),
        "HTTP": ServerInfo(
            protocol="HTTP",
            host=config.DATALOG_SERVER_HOST,
            port=config.DATALOG_HTTP_PORT,
            data_dir=os.path.join(base, "http"),
        ),
        "HTTPS": ServerInfo(
            protocol="HTTPS",
            host=config.DATALOG_SERVER_HOST,
            port=config.DATALOG_HTTPS_PORT,
            data_dir=os.path.join(base, "https"),
            ssl_certfile=config.DATALOG_SSL_CERT,
            ssl_keyfile=config.DATALOG_SSL_KEY,
        ),
    }


def channel_to_post_config(
    protocol: str,
    pool: dict[str, ServerInfo],
) -> PostChannelConfig:
    """
    根据协议名生成对应的 PostChannelConfig（填入本机服务器参数）。

    示例：
        pool = build_protocol_pool()
        ch_configs = {
            1: channel_to_post_config("FTP",   pool),   # Channel 1 → FTP
            2: channel_to_post_config("HTTP",  pool),   # Channel 2 → HTTP
            3: channel_to_post_config("SFTP",  pool),   # Channel 3 → SFTP
        }
    """
    key = protocol.upper()
    if key not in pool:
        raise ValueError(f"不支持的协议：{protocol}，可选：{list(pool)}")
    return pool[key].to_post_channel_config()


def build_active_servers(
    channel_configs: dict[int, PostChannelConfig],
    pool: dict[str, ServerInfo],
) -> list[ServerInfo]:
    """
    从 channel_configs 中提取实际用到的协议，返回对应的 ServerInfo 列表。
    确保每个协议只启动一次（即使多个 Channel 使用同一协议）。
    """
    used_protocols = {cfg.protocol.upper() for cfg in channel_configs.values()}
    servers = []
    for proto in used_protocols:
        if proto in pool:
            servers.append(pool[proto])
        else:
            log.warning("协议 %s 在协议池中未找到，跳过", proto)
    return servers


# ─────────────────────────────────────────────────────────────────────────────
# 端口占用检测与释放
# ─────────────────────────────────────────────────────────────────────────────

def _kill_port_process(port: int) -> None:
    """
    检查指定端口是否被其他进程占用，若占用则强制终止该进程。
    Windows 使用 netstat + taskkill；Unix/Mac 使用 lsof + kill。
    """
    pids: list[int] = []
    try:
        if platform.system() == "Windows":
            result = subprocess.run(
                ["netstat", "-ano"],
                capture_output=True, text=True, timeout=10,
            )
            for line in result.stdout.splitlines():
                # 匹配 TCP/UDP 行中 ":port " 或 ":port\t"
                if f":{port} " in line or f":{port}\t" in line:
                    parts = line.split()
                    if parts:
                        try:
                            pid = int(parts[-1])
                            if pid > 0 and pid not in pids:
                                pids.append(pid)
                        except ValueError:
                            pass
        else:
            result = subprocess.run(
                ["lsof", "-i", f":{port}", "-t"],
                capture_output=True, text=True, timeout=10,
            )
            for line in result.stdout.splitlines():
                try:
                    pid = int(line.strip())
                    if pid > 0 and pid not in pids:
                        pids.append(pid)
                except ValueError:
                    pass
    except Exception as e:
        log.warning("检查端口 %d 占用时出错：%s", port, e)
        return

    if not pids:
        return

    log.info("端口 %d 被进程 %s 占用，正在强制终止…", port, pids)
    for pid in pids:
        try:
            if platform.system() == "Windows":
                subprocess.run(
                    ["taskkill", "/F", "/PID", str(pid)],
                    capture_output=True, timeout=5,
                )
            else:
                os.kill(pid, 9)
            log.info("  已终止 PID=%d（占用端口 %d）", pid, port)
        except Exception as e:
            log.warning("  终止 PID=%d 失败：%s", pid, e)
    time.sleep(0.5)   # 等待端口释放


# ─────────────────────────────────────────────────────────────────────────────
# 服务器管理
# ─────────────────────────────────────────────────────────────────────────────

class ServerManager:
    """
    按需启动并管理后台服务器线程。
    只启动 build_active_servers() 返回的服务器（即 channel_configs 中实际用到的协议）。
    """

    def __init__(self, active_servers: list[ServerInfo]):
        self._servers = active_servers
        self._stops: list[threading.Event] = []

    def start_all(self):
        for srv in self._servers:
            ev = self._start_one(srv)
            self._stops.append(ev)
        if self._servers:
            time.sleep(1)   # 给服务器一点启动时间

    def _start_one(self, srv: ServerInfo) -> threading.Event:
        # 启动前先释放端口（若有其他进程占用则强制终止）
        _kill_port_process(srv.port)

        if srv.protocol == "FTP":
            _, ev = start_ftp_server(
                srv.host, srv.port, srv.username, srv.password, srv.data_dir)
        elif srv.protocol == "SFTP":
            _, ev = start_sftp_server(
                srv.host, srv.port, srv.username, srv.password, srv.data_dir)
        elif srv.protocol in ("HTTP", "HTTPS"):
            _, ev, _ = start_http_server(
                srv.host, srv.port, srv.data_dir,
                ssl_certfile=srv.ssl_certfile,
                ssl_keyfile=srv.ssl_keyfile,
            )
        else:
            log.warning("未知协议 %s，跳过启动", srv.protocol)
            ev = threading.Event()
        log.info("%s 服务器已启动：%s:%d  目录：%s",
                 srv.protocol, srv.host, srv.port, srv.data_dir)
        return ev

    def stop_all(self):
        for ev in self._stops:
            ev.set()
        log.info("所有服务器已停止")

    def __enter__(self):
        self.start_all()
        return self

    def __exit__(self, *_):
        self.stop_all()


# ─────────────────────────────────────────────────────────────────────────────
# 等待文件到达
# ─────────────────────────────────────────────────────────────────────────────

def wait_for_files(
    data_dirs: list[str],
    min_files: int = 1,
    timeout: float = 300,
    poll_interval: float = 5,
    extensions: tuple = (".json", ".csv"),
) -> list[str]:
    """
    轮询 data_dirs，直到累计出现 ≥ min_files 个匹配文件或超时。
    返回文件路径列表（按修改时间排序）。
    """
    log.info("等待文件推送（超时 %ds）…", int(timeout))
    deadline = time.time() + timeout
    while time.time() < deadline:
        found = _collect_files(data_dirs, extensions)
        if len(found) >= min_files:
            log.info("找到 %d 个文件", len(found))
            return found
        log.info("当前 %d 个文件，继续等待…", len(found))
        time.sleep(poll_interval)
    log.warning("等待超时（%ds），返回当前已有文件", int(timeout))
    return _collect_files(data_dirs, extensions)


def _collect_files(data_dirs: list[str], extensions: tuple) -> list[str]:
    found = []
    for d in data_dirs:
        if not os.path.isdir(d):
            continue
        for fn in os.listdir(d):
            if any(fn.lower().endswith(ext) for ext in extensions):
                found.append(os.path.join(d, fn))
    return sorted(found, key=lambda p: os.path.getmtime(p))


# ─────────────────────────────────────────────────────────────────────────────
# Selenium 初始化
# ─────────────────────────────────────────────────────────────────────────────

def _init_driver(gateway_url: str):
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    opts = Options()
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=opts)
    driver.maximize_window()
    log.info("打开网关：%s", gateway_url)
    driver.get(gateway_url)
    time.sleep(3)
    return driver


def _login(driver, username: str, password: str):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.wait import WebDriverWait
    try:
        user_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                "input[placeholder='Enter User Name']")))
    except Exception:
        # 登录页未出现 — 可能网关不可达或 URL 有误，直接退出
        raise RuntimeError(
            f"登录页未加载（未找到用户名输入框），当前 URL：{driver.current_url}。"
            "请确认网关地址 GATEWAY_WEB_URL 正确且网络可达。"
        )
    try:
        user_input.clear()
        user_input.send_keys(username)
        pass_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Enter Password']")
        pass_input.clear()
        pass_input.send_keys(password)
        btn = driver.find_element(By.XPATH, "//button[span[text()='Sign In']]")
        driver.execute_script("arguments[0].click()", btn)
        time.sleep(3)
        log.info("登录完成")
    except Exception as e:
        log.warning("登录操作异常（可能已登录）：%s", e)


# ─────────────────────────────────────────────────────────────────────────────
# 比对结果数据结构
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class FileVerifyResult:
    file_path: str
    protocol: str       # 来源服务器协议（FTP / SFTP / HTTP / HTTPS）
    channel_no: int     # 关联的 Post Channel 编号（0=未知）
    device_key: str
    scope: DatalogScopeReport
    unit_results: list[DatalogUnitResult]
    compare_results: list[DatalogCompareResult]
    timestamp_str: str
    error: str = ""


# ─────────────────────────────────────────────────────────────────────────────
# 设备识别
# ─────────────────────────────────────────────────────────────────────────────

_FILENAME_DEVICE_MAP = {
    "acurev4100":  "AcuRev4100",
    "acurev2100":  "AcuRev2100",
    "acuvimiiw":   "AcuvimIIW",
    "acuvimiir":   "AcuvimIIR",
    "acuvim3":     "AcuVIM3",
    "acurev1300":  "AcuRev1300",  # AcuRev1300 / PXM350（模板名 AcuRev-1300）
    "pxm350":      "AcuRev1300",
    "acuiom01":    "AcuIOM01",
    "acuiom02":    "AcuIOM02",
}

# 设备名 → Modbus 参数模块覆盖映射（当模块名与 device_name.lower() 不一致时使用）
_DEVICE_MODULE_OVERRIDE: dict[str, str] = {
    "AcuRev1300": "devices.pxm350",  # 模块文件名为 pxm350.py
}


_DEFAULT_DEVICE_KEY = "acurev4100"  # 文件名无法识别时的兜底设备


def _detect_device_key(filename: str) -> str:
    name_lower = filename.lower()
    for key in _FILENAME_DEVICE_MAP:
        if key in name_lower:
            return key
    log.warning("文件名 '%s' 无法识别设备，使用兜底 %s", filename, _DEFAULT_DEVICE_KEY)
    return _DEFAULT_DEVICE_KEY


def _select_latest_per_device(file_paths: list[str]) -> list[str]:
    """每个设备只保留最新的一个文件（按文件修改时间），跳过历史文件。"""
    from collections import defaultdict
    groups: dict[str, list[str]] = defaultdict(list)
    for fp in file_paths:
        key = _detect_device_key(Path(fp).name)
        groups[key].append(fp)
    result = []
    for key, files in groups.items():
        latest = max(files, key=lambda p: Path(p).stat().st_mtime)
        result.append(latest)
        if len(files) > 1:
            log.info("设备 %s：共 %d 个文件，仅分析最新：%s", key, len(files), Path(latest).name)
    return result


# ─────────────────────────────────────────────────────────────────────────────
# 比对
# ─────────────────────────────────────────────────────────────────────────────

async def verify_files(
    file_paths: list[str],
    dir_to_protocol: dict[str, str],
    dir_to_channel: dict[str, int],
) -> list[FileVerifyResult]:
    """对文件列表逐一运行比对，返回结果列表。"""
    results = []
    # 会话级离线缓存：(host, port, unit) → error_msg，跳过已知离线设备
    offline_cache: dict[tuple, str] = {}

    for fpath in file_paths:
        device_key = _detect_device_key(Path(fpath).name)
        device_name = _FILENAME_DEVICE_MAP.get(device_key, config.DEVICE_NAME)
        fdir = str(Path(fpath).parent)
        protocol = dir_to_protocol.get(fdir, "UNKNOWN")
        channel_no = dir_to_channel.get(fdir, 0)

        config.DEVICE_NAME = device_name
        # 模块名：先查 override 映射，再按 device_name 推导
        config.DEVICE_MODULE = _DEVICE_MODULE_OVERRIDE.get(
            device_name, f"devices.{device_name.lower().replace(' ', '')}")
        if device_name in config.MODBUS_DEVICE_MAP:
            config.MODBUS_HOST, config.MODBUS_PORT, config.MODBUS_UNIT = \
                config.MODBUS_DEVICE_MAP[device_name]

        # 离线缓存命中：跳过本文件，直接记录错误
        host_key = (config.MODBUS_HOST, config.MODBUS_PORT, config.MODBUS_UNIT)
        if host_key in offline_cache:
            cached_err = offline_cache[host_key]
            log.info("跳过（设备已知离线）：%s  %s", Path(fpath).name, cached_err)
            results.append(FileVerifyResult(
                file_path=fpath, protocol=protocol, channel_no=channel_no,
                device_key=device_key,
                scope=DatalogScopeReport(0, 0, [], [], []),
                unit_results=[], compare_results=[], timestamp_str="",
                error=f"[离线缓存] {cached_err}",
            ))
            continue

        log.info("验证：%s  设备：%s  来源：%s（Channel %d）",
                 Path(fpath).name, device_name, protocol, channel_no)
        try:
            scope, unit_results, compare_results, timestamp_str = \
                await run_datalog_comparison(fpath)
            results.append(FileVerifyResult(
                file_path=fpath, protocol=protocol, channel_no=channel_no,
                device_key=device_key, scope=scope, unit_results=unit_results,
                compare_results=compare_results, timestamp_str=timestamp_str,
            ))
        except Exception as exc:
            err_str = str(exc)
            log.error("文件 %s 比对失败：%s", fpath, exc)
            # 连接类错误（TCP 连接失败或读取后断连）加入离线缓存
            _is_conn_err = any(k in err_str for k in (
                "连接失败", "未连接", "No response", "Failed to connect",
                "ConnectionRefusedError", "Connection refused",
            ))
            if _is_conn_err:
                offline_cache[host_key] = err_str
                log.warning("设备 %s:%d unit=%d 已加入离线缓存，后续同设备文件将跳过",
                            *host_key)
            results.append(FileVerifyResult(
                file_path=fpath, protocol=protocol, channel_no=channel_no,
                device_key=device_key,
                scope=DatalogScopeReport(0, 0, [], [], []),
                unit_results=[], compare_results=[], timestamp_str="",
                error=err_str,
            ))
    return results


# ─────────────────────────────────────────────────────────────────────────────
# 汇总 HTML 报告
# ─────────────────────────────────────────────────────────────────────────────

def generate_summary_report(
    results: list[FileVerifyResult],
    output_path: Optional[str] = None,
) -> str:
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = Path(config.REPORT_DIR)
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(report_dir / f"datalog_server_{ts}.html")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total_files  = len(results)
    total_params = sum(len(r.compare_results) for r in results)
    total_pass   = sum(sum(1 for c in r.compare_results if c.status == "PASS")
                       for r in results)
    total_fail   = sum(sum(1 for c in r.compare_results if c.status == "FAIL")
                       for r in results)
    total_err    = total_params - total_pass - total_fail
    pass_rate    = f"{total_pass / total_params * 100:.1f}%" if total_params else "N/A"

    sections_html = ""
    for r in results:
        s = summary(r.compare_results) if r.compare_results else {}
        file_ok = s.get("fail", 0) == 0 and s.get("error", 0) == 0 and not r.error
        border_color = "#28a745" if file_ok else "#dc3545"
        hdr_bg = "#d4edda" if file_ok else "#f8d7da"

        unit_fail_count = sum(1 for u in r.unit_results if u.status == "FAIL")
        unit_badge = ""
        if r.unit_results:
            unit_badge = (
                f'<span class="badge err-badge">单位不匹配 {unit_fail_count}</span>'
                if unit_fail_count else
                '<span class="badge ok-badge">单位一致</span>'
            )

        channel_tag = f"Channel {r.channel_no}" if r.channel_no else ""
        rows_html = []
        status_colors = {
            "PASS": "#d4edda", "FAIL": "#f8d7da",
            "FILE_ERR": "#fff3cd", "MODBUS_ERR": "#fff3cd", "BOTH_ERR": "#e2e3e5",
        }
        status_labels = {
            "PASS": "通过", "FAIL": "失败",
            "FILE_ERR": "文件异常", "MODBUS_ERR": "Modbus异常", "BOTH_ERR": "双路异常",
        }
        for j, c in enumerate(r.compare_results):
            cbg   = status_colors.get(c.status, "#fff")
            cstat = status_labels.get(c.status, c.status)
            err_hint = ""
            if c.file_error:
                err_hint += f"文件: {html.escape(c.file_error)}"
            if c.modbus_error:
                err_hint += f"{'<br>' if err_hint else ''}Modbus: {html.escape(c.modbus_error)}"
            fv = f"{c.file_value:.6g}"   if c.file_value   is not None else "—"
            mv = f"{c.modbus_value:.6g}" if c.modbus_value is not None else "—"
            dp = f"{c.diff_pct:.3f}%"   if c.diff_pct     is not None else "—"
            rows_html.append(f"""
            <tr style="background:{cbg}">
              <td class="num">{j+1}</td>
              <td class="key">{html.escape(c.param_key)}</td>
              <td class="val">{fv}</td><td class="val">{mv}</td>
              <td class="val">{dp}</td>
              <td class="stat">{cstat}</td>
              <td class="err">{err_hint}</td>
            </tr>""")

        sections_html += f"""
<details class="section" style="border-left:4px solid {border_color}">
<summary style="background:{hdr_bg}">
  <span class="proto-tag">{html.escape(r.protocol)}</span>
  {f'<span class="ch-tag">{channel_tag}</span>' if channel_tag else ""}
  {html.escape(Path(r.file_path).name)}
  <span class="sum-info"> · {r.device_key} · 共 {len(r.compare_results)} 参数
    · PASS={s.get("pass",0)} FAIL={s.get("fail",0)} ERR={s.get("error",0)}</span>
  {unit_badge}
  {'<span class="badge ok-badge">全部通过</span>' if file_ok else '<span class="badge err-badge">有失败项</span>'}
</summary>
<div class="section-body">
  {"<p style='color:#dc3545'>错误：" + html.escape(r.error) + "</p>" if r.error else ""}
  <div class="cards">
    <div class="card total"><div class="num">{len(r.compare_results)}</div><div class="lbl">总参数</div></div>
    <div class="card pass"><div class="num">{s.get("pass",0)}</div><div class="lbl">PASS</div></div>
    <div class="card fail"><div class="num">{s.get("fail",0)}</div><div class="lbl">FAIL</div></div>
    <div class="card err"><div class="num">{s.get("error",0)}</div><div class="lbl">ERROR</div></div>
  </div>
  <table>
  <colgroup><col style="width:44px"><col style="width:260px"><col style="width:100px">
    <col style="width:100px"><col style="width:80px"><col style="width:80px"><col></colgroup>
  <thead><tr><th>#</th><th>参数名</th><th>文件值</th><th>Modbus值</th>
    <th>相对差</th><th>结果</th><th>错误</th></tr></thead>
  <tbody>{"".join(rows_html)}</tbody>
  </table>
</div>
</details>"""

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<title>DataLog 服务器推送验证报告</title>
<style>
  body{{font-family:"Microsoft YaHei",Arial,sans-serif;font-size:13px;
       margin:20px;background:#f5f5f5;color:#333}}
  h1{{font-size:20px;margin-bottom:4px}}
  .meta{{color:#666;font-size:12px;margin-bottom:16px}}
  .cards{{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap}}
  .card{{background:#fff;border-radius:6px;padding:14px 20px;
         box-shadow:0 1px 4px rgba(0,0,0,.1);min-width:110px;text-align:center}}
  .card .num{{font-size:28px;font-weight:bold}}
  .card .lbl{{font-size:11px;color:#888;margin-top:2px}}
  .card.pass .num{{color:#28a745}}.card.fail .num{{color:#dc3545}}
  .card.err .num{{color:#ffc107}}.card.total .num{{color:#007bff}}
  table{{border-collapse:collapse;width:100%;background:#fff;
         box-shadow:0 1px 4px rgba(0,0,0,.1);border-radius:6px;
         overflow:hidden;margin-bottom:16px}}
  thead tr{{background:#343a40;color:#fff}}
  th{{padding:8px;text-align:center;font-size:12px}}
  td{{padding:5px 8px;border-bottom:1px solid #eee;vertical-align:middle}}
  td.num{{text-align:right;color:#999;font-size:11px}}
  td.key{{font-family:monospace;font-size:12px;word-break:break-all}}
  td.val{{text-align:right;font-family:monospace;font-size:12px}}
  td.stat{{text-align:center;font-weight:bold;font-size:12px}}
  td.err{{font-size:11px;color:#666;word-break:break-all}}
  details.section{{background:#fff;border-radius:8px;margin-bottom:16px;
                   box-shadow:0 1px 4px rgba(0,0,0,.1);overflow:hidden}}
  details.section>summary{{list-style:none;cursor:pointer;padding:12px 16px;
    font-size:14px;font-weight:bold;display:flex;align-items:center;
    gap:6px;user-select:none}}
  details.section>summary::-webkit-details-marker{{display:none}}
  details.section>summary::before{{content:"▶";font-size:10px;margin-right:4px}}
  details[open].section>summary::before{{content:"▼"}}
  .sum-info{{font-size:12px;font-weight:normal;color:#555;margin-left:4px}}
  .badge,.proto-tag,.ch-tag{{display:inline-block;padding:2px 8px;border-radius:10px;
         font-size:11px;font-weight:bold;margin-left:2px}}
  .ok-badge{{background:#d4edda;color:#155724}}
  .err-badge{{background:#f8d7da;color:#721c24}}
  .proto-tag{{background:#cce5ff;color:#004085}}
  .ch-tag{{background:#e2e3e5;color:#383d41}}
  .section-body{{padding:16px}}
</style>
</head>
<body>
<h1>DataLog 服务器推送验证报告</h1>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  收到文件：{total_files} 个 &nbsp;|&nbsp;
  总参数：{total_params} &nbsp;|&nbsp;
  通过率：{pass_rate}
</div>
<div class="cards">
  <div class="card total"><div class="num">{total_files}</div><div class="lbl">收到文件数</div></div>
  <div class="card total"><div class="num">{total_params}</div><div class="lbl">总参数</div></div>
  <div class="card pass"><div class="num">{total_pass}</div><div class="lbl">PASS ({pass_rate})</div></div>
  <div class="card fail"><div class="num">{total_fail}</div><div class="lbl">FAIL</div></div>
  <div class="card err"><div class="num">{total_err}</div><div class="lbl">ERROR</div></div>
</div>
{sections_html}
</body></html>"""

    Path(output_path).write_text(html_content, encoding="utf-8")
    log.info("汇总报告已保存：%s", output_path)
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────────────────────────────────────────

def run(
    channel_configs: dict[int, PostChannelConfig],
    logger_configs: dict[int, DataLoggerConfig],
    param_config: Optional[DataLogParamConfig] = None,
    pool: dict[str, ServerInfo] = None,
    gateway_url: str = "",
    gateway_user: str = "admin",
    gateway_pass: str = "Admin@110001",
    wait_timeout: float = 300,
    use_webdriver: bool = True,
    use_servers: bool = True,
) -> str:
    """
    主入口。

    channel_configs: {channel_n: PostChannelConfig}
        每个 PostChannelConfig.protocol 决定该 Channel 使用哪种协议。
        例：{1: PostChannelConfig("FTP",...), 2: PostChannelConfig("HTTP",...), ...}
        用 channel_to_post_config() 工具函数快速构建。

    param_config: DataLogParamConfig（可选）
        若提供，在配置 Data Logger 之前先到 Data Log Parameter Config 页面
        为每台设备选择参数（遍历所有参数类型并全选）。

    pool: 协议池，默认由 build_protocol_pool() 从 config.py 读取。
    """
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )

    if pool is None:
        pool = build_protocol_pool()

    # 只启动实际用到的协议服务器
    active_servers = build_active_servers(channel_configs, pool) if use_servers else []

    # 构建目录→协议 和 目录→Channel 映射（用于后续标注报告来源）
    dir_to_protocol: dict[str, str] = {}
    dir_to_channel: dict[str, int] = {}
    for ch_n, ch_cfg in channel_configs.items():
        proto = ch_cfg.protocol.upper()
        if proto in pool:
            d = os.path.normpath(pool[proto].data_dir)
            dir_to_protocol[d] = proto
            # 若多个 Channel 使用同一协议，记录最小 Channel 编号
            if d not in dir_to_channel or ch_n < dir_to_channel[d]:
                dir_to_channel[d] = ch_n

    data_dirs = list(dir_to_protocol.keys())

    mgr = ServerManager(active_servers)
    driver = None
    dl_page = None

    # 每次运行前清空数据目录，避免分析历史缓存文件（历史数据与当前 Modbus 读值存在时序差）
    for d in data_dirs:
        if os.path.isdir(d):
            removed = 0
            for fn in os.listdir(d):
                if fn.lower().endswith((".json", ".csv")):
                    try:
                        os.remove(os.path.join(d, fn))
                        removed += 1
                    except OSError:
                        pass
            if removed:
                log.info("已清除旧文件：%s（共 %d 个）", d, removed)

    try:
        if use_servers:
            mgr.start_all()

        if use_webdriver and gateway_url:
            driver = _init_driver(gateway_url)
            _login(driver, gateway_user, gateway_pass)

            pc_page = PostChannelPage(driver)
            pc_results = pc_page.configure_all(channel_configs, test=True)
            for n, res in pc_results.items():
                log.info("Post Channel %d 测试结果：%s", n, res)

            if param_config and param_config.device_names:
                log.info("配置 Data Log Parameter Config（共 %d 台设备）…",
                         len(param_config.device_names))
                param_page = DataLogParamConfigPage(driver)
                param_page.navigate()
                param_page.configure_all(param_config)

            dl_page = DataLoggerPage(driver)
            dl_page.configure_all(logger_configs)

        file_paths = wait_for_files(data_dirs, timeout=wait_timeout)

        # 收到推送后立即禁用 Data Logger，避免继续累积历史数据
        if file_paths and dl_page is not None:
            log.info("已收到推送数据，正在禁用 Data Loggers…")
            dl_page.disable_all(logger_configs)

        if not file_paths:
            log.warning("未收到任何文件")
        else:
            file_paths = _select_latest_per_device(file_paths)
            log.info("仅分析各设备最新文件：共 %d 个", len(file_paths))

        # 规范化路径后匹配目录
        normalized_dir_to_protocol = {
            os.path.normpath(k): v for k, v in dir_to_protocol.items()}
        normalized_dir_to_channel = {
            os.path.normpath(k): v for k, v in dir_to_channel.items()}

        verify_results = asyncio.run(
            verify_files(file_paths, normalized_dir_to_protocol, normalized_dir_to_channel))

        report_path = generate_summary_report(verify_results)
        print(f"\n  汇总报告：{report_path}\n")
        return report_path

    finally:
        if driver:
            driver.quit()
        if use_servers:
            mgr.stop_all()


# ─────────────────────────────────────────────────────────────────────────────
# 命令行入口
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="DataLog 服务器推送端到端验证")
    parser.add_argument("--no-webdriver", action="store_true",
                        help="跳过网关配置，仅验证本地已有文件")
    parser.add_argument("--no-servers",  action="store_true",
                        help="不启动本地服务器（假设已有外部服务器）")
    parser.add_argument("--device", default="",
                        help="默认设备 key（无法从文件名推断时使用）")
    parser.add_argument("--timeout", type=float, default=300,
                        help="等待文件超时（秒），默认 300")
    args = parser.parse_args()

    if args.device:
        from datalog_comparator import _DEVICE_MAP
        if args.device.lower() in _DEVICE_MAP:
            config.DEVICE_NAME, config.DEVICE_MODULE = _DEVICE_MAP[args.device.lower()]

    # ── 构建协议池 ────────────────────────────────────────────────────────────
    _pool = build_protocol_pool()

    # ── 自由配置 Post Channel：每个 Channel 可独立选任意协议 ─────────────────
    #
    #   channel_to_post_config("FTP",   _pool)  → 用本机 FTP 服务器参数
    #   channel_to_post_config("SFTP",  _pool)  → 用本机 SFTP 服务器参数
    #   channel_to_post_config("HTTP",  _pool)  → 用本机 HTTP 服务器参数
    #   channel_to_post_config("HTTPS", _pool)  → 用本机 HTTPS 服务器参数
    #
    #   修改下方来更改各 Channel 的协议分配：
    _channel_configs: dict[int, PostChannelConfig] = {
        1: channel_to_post_config("FTP", _pool),
    }

    # ── Data Log Parameter Config（可选）──────────────────────────────────────
    # 在 Data Logger 启用之前，为指定设备全选推送参数。
    # device_names: UI Device 下拉中的设备名（须与页面显示文字精确匹配）
    # param_types : 留空 = 自动遍历所有可用类型并全选；
    #               指定 = 只处理列表中的类型，如 ["Basic", "Energy"]
    _param_config = DataLogParamConfig(
        device_names=[],   # 填入要配置的设备名，如 ["AcuRev4100", "AcuRev2100"]
        param_types=[],    # 留空 = 全部类型
    )

    # ── Data Logger 配置（关联 Post Channel 编号 + 推送参数） ─────────────────
    _logger_configs: dict[int, DataLoggerConfig] = {
        1: DataLoggerConfig(channel_index=1, log_interval="1 minute", log_file_format="json",
                            device_names=[]),   # 空列表 = 全选设备
    }

    run(
        channel_configs=_channel_configs,
        logger_configs=_logger_configs,
        param_config=_param_config if _param_config.device_names else None,
        pool=_pool,
        gateway_url=config.GATEWAY_WEB_URL,
        gateway_user=config.GATEWAY_WEB_USER,
        gateway_pass=config.GATEWAY_WEB_PASS,
        wait_timeout=args.timeout,
        use_webdriver=not args.no_webdriver,
        use_servers=not args.no_servers,
    )
