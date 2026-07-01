# -*- coding: utf-8 -*-
"""
conftest.py — DataLog 测试套件共享 fixtures（Playwright 版）

Session 级：
  pool      协议池（FTP/SFTP/HTTP/HTTPS ServerInfo）
  servers   按需启动对应协议服务器（.setup_done 存在时跳过）
  driver    Playwright Page（从 app_page 继承，已登录）；
            .setup_done 存在时跳过 Post Channel 和 Data Log Parameter Config 配置

Function 级（autouse）：
  clear_dirs  每个测试前清空所有协议数据目录
"""
from __future__ import annotations

import json
import logging
import os
import shutil
import sys
import time
from pathlib import Path

import pytest

_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent.parent  # AcuHMI_1_7/
sys.path.insert(0, str(_PROJECT_ROOT))
sys.path.insert(0, str(_HERE))  # 优先级最高，确保本地 config.py 覆盖其他同名模块

import config

# 仓库根目录入 path，便于 import projects.AcuHMI_1_7.helpers（对齐 BacnetIP conftest）
_REPO_ROOT = str(Path(__file__).resolve().parents[4])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from projects.AcuHMI_1_7.helpers.physical_devices_reader import (  # noqa: E402
    discover_modbus_tcp_devices,
)


# ─── pytest-html 报告：直跑 data_log 时自动产出 report.html ──────────────────────

@pytest.hookimpl(tryfirst=True)
def pytest_configure(config: pytest.Config) -> None:
    """直接 `pytest .../data_log/` 运行时也产出 pytest-html 报告（report.html）。

    data_log 子目录有独立 pytest.ini（rootdir 锁定到此），其 addopts 仅含 Allure、命令行不带
    --html，故直跑只生成 Allure、不生成 report.html。这里在**未显式指定 --html** 时补一个默认
    html 路径，使 data_log 与其它模块一样在 reports/AcuHMI_1_7/data_log/<时间戳>/report.html
    产出单文件报告（Allure 报告照常生成，二者并存 = “同时”）。

    路径优先复用 framework runner 注入的 REPORT_DIR/RUN_TS（经 runner 运行时 runner 已带
    --html，此处检测到 htmlpath 已设置即不覆盖）；直跑无这些环境变量时回退到按时间戳建目录。
    tryfirst 保证本钩子在 pytest-html 自身的 pytest_configure 读取 htmlpath 之前先设好值。
    """
    if not config.pluginmanager.hasplugin("html"):
        return
    if getattr(config.option, "htmlpath", None):
        return  # 已由 runner 或命令行 --html 指定，不覆盖
    from datetime import datetime
    ts = os.environ.get("RUN_TS") or datetime.now().strftime("%Y%m%d_%H%M%S")
    report_root = os.environ.get("REPORT_DIR")
    base = Path(report_root) if report_root else (
        Path(_REPO_ROOT) / "reports" / "AcuHMI_1_7" / "data_log" / ts
    )
    base.mkdir(parents=True, exist_ok=True)
    config.option.htmlpath = str(base / "report.html")
    config.option.self_contained_html = True


# ─── Allure 报告自动生成（每次 pytest 结束后生成带时间戳的报告目录）────────────────

def _resolve_allure_exe() -> str | None:
    """定位 allure 可执行文件。

    优先用 PATH 上的 allure（含 npm 安装的 allure.CMD / Scoop / 解压版 bin 目录），
    其次回退到历史硬编码安装路径。两者都没有则返回 None，由调用方降级处理。
    旧实现把路径写死成单一机器的 allure.bat，换机器即失效（report 静默不生成），故改为自动发现。
    """
    found = shutil.which("allure")
    if found:
        return found
    for cand in (
        Path(r"C:\work\tools\allure\allure-2.32.0\bin\allure.bat"),
    ):
        if cand.exists():
            return str(cand)
    return None


def pytest_sessionfinish(session, exitstatus):
    import subprocess
    from datetime import datetime
    _log = logging.getLogger(__name__)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    results_dir = _HERE / "reports" / "allure-results"
    report_dir  = _HERE / "reports" / f"allure-report-{ts}"
    allure_exe = _resolve_allure_exe()
    if not results_dir.exists():
        _log.warning("Allure 结果目录不存在，跳过报告生成：%s", results_dir)
        return
    if not allure_exe:
        _log.warning(
            "未找到 allure 可执行文件（PATH 与已知安装路径均无），跳过报告生成；"
            "结果已存于 %s，可手动执行 `allure generate` 或安装 allure CLI",
            results_dir,
        )
        return
    subprocess.run(
        [allure_exe, "generate", str(results_dir),
         "-o", str(report_dir), "--clean"],
        check=False,
    )
    _log.info("Allure 报告已生成：%s", report_dir)
from datalog_server_verifier import (
    ServerInfo,
    ServerManager,
    build_protocol_pool,
)
from datalog_page import (
    DataLogParamConfigPage,
    PostChannelConfig,
    PostChannelPage,
)

log = logging.getLogger(__name__)

_DEVICE_CACHE = _HERE / ".datalog_device_cache.json"
_SETUP_FLAG   = _HERE / ".setup_done"


def _load_device_cache() -> list:
    try:
        if _DEVICE_CACHE.exists():
            return json.loads(_DEVICE_CACHE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return []


def _save_device_cache(devices: list):
    try:
        _DEVICE_CACHE.write_text(
            json.dumps(devices, ensure_ascii=False, indent=2), encoding="utf-8")
        log.info("已更新设备缓存：%s", devices)
    except Exception as e:
        log.warning("保存设备缓存失败：%s", e)


# ─── session fixtures ──────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def pool() -> dict[str, ServerInfo]:
    return build_protocol_pool()


def _collect_needed_protocols(request) -> set[str]:
    needed: set[str] = set()
    for item in request.session.items:
        if not hasattr(item, "callspec"):
            continue
        params = item.callspec.params
        if "case" in params:
            case = params["case"]
            if hasattr(case, "protocol"):
                needed.add(case.protocol)
        if "check_protos" in params:
            needed.update(params["check_protos"])
    return needed


@pytest.fixture(scope="session")
def servers(pool, request):
    """按需启动协议服务器；.setup_done 存在时跳过。"""
    if _SETUP_FLAG.exists():
        log.info("检测到 .setup_done，服务器已由 setup_env.py 启动，跳过")
        yield None
        return

    needed_protocols = _collect_needed_protocols(request)
    if not needed_protocols:
        needed_protocols = set(pool.keys()) - {"HTTPS"}

    active = [
        si for si in pool.values()
        if si.protocol in needed_protocols
        and (si.protocol != "HTTPS" or (si.ssl_certfile and si.ssl_keyfile))
    ]
    if not active:
        log.info("当前测试集无需服务器，跳过启动")
        yield None
        return

    log.info("按需启动服务器：%s", [si.protocol for si in active])
    mgr = ServerManager(active)
    mgr.start_all()
    time.sleep(1)
    yield mgr
    mgr.stop_all()


@pytest.fixture(scope="session")
def app_page(playwright):
    """Session 级已登录 Playwright Page（复用根 conftest 的唯一 sync_playwright 实例）。

    不再自起 ``sync_playwright().start()``：全量套件运行时主线程已有 pytest-playwright
    插件留下的运行中事件循环，第二个 sync_playwright 会撞上它抛 "Playwright Sync API
    inside the asyncio loop"。改为从根 conftest 的 session 级 ``playwright`` 实例另起
    一个独立浏览器（仍保留 data_log 自己的有头/最大化/忽略证书启动参数与生命周期）。
    """
    browser = playwright.chromium.launch(
        headless=False,
        args=["--ignore-certificate-errors", "--start-maximized"],
    )
    ctx = browser.new_context(ignore_https_errors=True, no_viewport=True)
    page = ctx.new_page()
    page.goto(config.GATEWAY_WEB_URL)
    page.wait_for_timeout(3000)
    try:
        user_input = page.locator("input[placeholder='Enter User Name']").first
        user_input.wait_for(state="visible", timeout=10000)
        user_input.fill(config.GATEWAY_WEB_USER)
        page.locator("input[placeholder='Enter Password']").first.fill(config.GATEWAY_WEB_PASS)
        page.locator("xpath=//button[span[text()='Sign In']]").first.evaluate("el => el.click()")
        page.wait_for_timeout(3000)
        log.info("app_page：登录完成")
    except Exception as e:
        log.warning("app_page：登录操作异常（可能已登录）：%s", e)
    yield page
    try:
        ctx.close()
        browser.close()
    except Exception:
        pass


@pytest.fixture(scope="session")
def driver(app_page, pool, servers):
    """
    Playwright Page session fixture。

    一次性预置（.setup_done 不存在时）：
      1. Post Channel 1=FTP / 2=SFTP / 3=HTTP：配置并 Enable。
      2. Data Log Parameter Config：为所有设备全选参数类型。
    """
    page = app_page  # 已由根 conftest 登录

    if _SETUP_FLAG.exists():
        log.info("检测到 .setup_done，Post Channel 和 Data Log Parameter Config 已由 setup_env.py 配置，跳过")
    else:
        # ── 1. Post Channel 一次性配置 ─────────────────────────────────────────
        pc_page = PostChannelPage(page)
        channel_cfgs: dict[int, PostChannelConfig] = {}
        if "FTP"  in pool: channel_cfgs[1] = pool["FTP"].to_post_channel_config()
        if "SFTP" in pool: channel_cfgs[2] = pool["SFTP"].to_post_channel_config()
        if "HTTP" in pool: channel_cfgs[3] = pool["HTTP"].to_post_channel_config()
        if channel_cfgs:
            pc_page.configure_all(channel_cfgs, enabled=True, test=False)
            log.info("Post Channel 1/2/3 配置并启用完成")

        # ── 2. Data Log Parameter Config 一次性配置 ────────────────────────────
        try:
            param_page = DataLogParamConfigPage(page)
            param_page.navigate()
            current_devices = param_page._get_dropdown_options(param_page._DEVICE_SELECT)
            cached_devices  = _load_device_cache()

            if current_devices and set(current_devices) == set(cached_devices):
                log.info("Data Log Parameter Config：设备列表未变（%s），跳过配置", current_devices)
            elif current_devices:
                log.info("Data Log Parameter Config：发现 %d 台设备，开始配置：%s",
                         len(current_devices), current_devices)
                param_page.configure_all_devices()
                _save_device_cache(current_devices)
                log.info("Data Log Parameter Config 配置完成")
            else:
                log.warning("Data Log Parameter Config：Device 下拉无可用设备，跳过")
        except Exception as e:
            log.warning("Data Log Parameter Config 配置跳过：%s", e)

    # 动态发现下挂 Modbus TCP 设备的连接参数，供 verify_files 优先取用。
    # 失败不致命：留空则 verify_files 自动回退 config.yaml 静态表。
    try:
        config.DISCOVERED_DEVICES = discover_modbus_tcp_devices(page)
        log.info("动态发现 %d 台 Modbus TCP 设备", len(config.DISCOVERED_DEVICES))
    except Exception as e:  # noqa: BLE001
        config.DISCOVERED_DEVICES = []
        log.warning("动态设备发现失败，将回退 config.yaml 静态连接表：%s", e)

    yield page


# ─── function fixtures ─────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def clear_dirs(pool):
    """每个测试前清空所有协议数据目录中的 .json / .csv 文件。"""
    for si in pool.values():
        d = si.data_dir
        if not os.path.isdir(d):
            continue
        for fn in os.listdir(d):
            if fn.lower().endswith((".json", ".csv")):
                try:
                    os.remove(os.path.join(d, fn))
                except OSError:
                    pass
