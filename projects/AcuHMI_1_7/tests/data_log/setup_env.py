#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
setup_env.py — DataLog 测试环境预置脚本

1. 启动本地 FTP / SFTP / HTTP 服务器（后台线程）
2. Playwright 登录网关
3. 配置 Post Channel 1=FTP / 2=SFTP / 3=HTTP（Enabled + 保存）
4. 配置 Data Log Parameter Config（为所有设备全选参数）
5. 写入 .setup_done 和 .setup_pid 标志文件
6. 阻塞等待（保持服务器运行），直至 Ctrl+C

运行方式（从仓库根 autotest/ 执行）：
    python projects/AcuHMI_1_7/tests/data_log/setup_env.py
"""
from __future__ import annotations

import json
import logging
import os
import sys
import time
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent.parent.parent  # AcuHMI-1-7/
sys.path.insert(0, str(_PROJECT_ROOT))
sys.path.insert(0, str(_HERE))  # 优先级最高，确保本地 config.py 覆盖其他同名模块

import config
from datalog_server_verifier import ServerManager, build_protocol_pool
from datalog_page import DataLogParamConfigPage, PostChannelConfig, PostChannelPage

_SETUP_FLAG   = _HERE / ".setup_done"
_PID_FILE     = _HERE / ".setup_pid"
_DEVICE_CACHE = _HERE / ".datalog_device_cache.json"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


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
    except Exception as e:
        log.warning("保存设备缓存失败：%s", e)


def main():
    pool = build_protocol_pool()

    # ── 1. 确保数据目录存在并启动服务器 ───────────────────────────────────────
    for si in pool.values():
        os.makedirs(si.data_dir, exist_ok=True)

    active = [si for si in pool.values() if si.protocol in ("FTP", "SFTP", "HTTP")]
    mgr = ServerManager(active)
    mgr.start_all()
    log.info("FTP / SFTP / HTTP 服务器已启动")

    # ── 2. Playwright 登录配置 ─────────────────────────────────────────────────
    from playwright.sync_api import sync_playwright
    _pw = sync_playwright().start()
    try:
        browser = _pw.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors", "--start-maximized"],
        )
        ctx = browser.new_context(ignore_https_errors=True, no_viewport=True)
        page = ctx.new_page()
        page.goto(config.GATEWAY_WEB_URL)
        page.wait_for_timeout(3000)

        # 登录
        try:
            user_input = page.locator("input[placeholder='Enter User Name']").first
            user_input.wait_for(state="visible", timeout=10000)
            user_input.fill(config.GATEWAY_WEB_USER)
            page.locator("input[placeholder='Enter Password']").first.fill(config.GATEWAY_WEB_PASS)
            page.locator("xpath=//button[span[text()='Sign In']]").first.evaluate("el => el.click()")
            page.wait_for_timeout(3000)
            log.info("登录完成")
        except Exception as e:
            log.warning("登录操作异常（可能已登录）：%s", e)

        # ── 3. Post Channel 1=FTP / 2=SFTP / 3=HTTP ─────────────────────────
        pc_page = PostChannelPage(page)
        channel_cfgs: dict[int, PostChannelConfig] = {}
        if "FTP"  in pool: channel_cfgs[1] = pool["FTP"].to_post_channel_config()
        if "SFTP" in pool: channel_cfgs[2] = pool["SFTP"].to_post_channel_config()
        if "HTTP" in pool: channel_cfgs[3] = pool["HTTP"].to_post_channel_config()
        if channel_cfgs:
            pc_page.configure_all(channel_cfgs, enabled=True, test=False)
            log.info("Post Channel 1/2/3 配置并启用完成")

        # ── 4. Data Log Parameter Config ──────────────────────────────────────
        try:
            param_page = DataLogParamConfigPage(page)
            param_page.navigate()
            current_devices = param_page._get_dropdown_options(param_page._DEVICE_SELECT)
            cached_devices  = _load_device_cache()

            if current_devices and set(current_devices) == set(cached_devices):
                log.info("设备列表未变（%s），跳过 Data Log Parameter Config 配置", current_devices)
            elif current_devices:
                log.info("发现 %d 台设备，开始配置：%s", len(current_devices), current_devices)
                # 已在页面上，直接逐设备配置，无需再次导航
                for device_name in current_devices:
                    param_page.configure_device(device_name, [])
                _save_device_cache(current_devices)
                log.info("Data Log Parameter Config 配置完成")
            else:
                log.warning("Device 下拉无可用设备，跳过配置")
        except Exception as e:
            log.warning("Data Log Parameter Config 配置跳过：%s", e)

        # ── 5. 写入标志文件 ────────────────────────────────────────────────────
        _SETUP_FLAG.write_text("ok", encoding="utf-8")
        _PID_FILE.write_text(str(os.getpid()), encoding="utf-8")
        log.info(".setup_done 已写入，测试可以开始（Ctrl+C 停止服务器）")

    finally:
        # 配置完成后立即关闭浏览器，服务器继续在后台线程运行
        try:
            browser.close()
        except Exception:
            pass
        try:
            _pw.stop()
        except Exception:
            pass

    # ── 6. 阻塞等待（服务器保持运行）────────────────────────────────────────
    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        log.info("收到 Ctrl+C，正在停止…")
    finally:
        mgr.stop_all()
        _SETUP_FLAG.unlink(missing_ok=True)
        _PID_FILE.unlink(missing_ok=True)
        log.info("清理完成")


if __name__ == "__main__":
    main()
