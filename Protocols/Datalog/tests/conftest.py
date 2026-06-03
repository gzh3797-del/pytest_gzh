# -*- coding: utf-8 -*-
"""
conftest.py — DataLog 测试套件共享 fixtures

Session 级（一次性配置，整个 session 复用）：
  pool      协议池（FTP/SFTP/HTTP/HTTPS ServerInfo）
  servers   按测试用例实际需要，只启动用到的协议服务器
  driver    Selenium WebDriver，登录一次复用；
            - Post Channel 1=FTP / 2=SFTP / 3=HTTP：配置一次后 Disable，
              各用例按需 enable_channel(n) 启用，无需每次重新配置
            - Data Log Parameter Config：配置一次（全设备全参数）

Function 级（autouse）：
  clear_dirs  每个测试前清空 4 个协议数据目录，避免缓存文件干扰
"""
from __future__ import annotations

import json
import logging
import os
import sys
import time
from pathlib import Path

import pytest

# 路径设置
sys.path.insert(0, str(Path(__file__).parent))               # Protocols/Datalog/tests/
sys.path.insert(0, str(Path(__file__).parent.parent))        # Protocols/Datalog/
sys.path.insert(0, str(Path(__file__).parent.parent.parent)) # Protocols/
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent))  # 仓库根

import config
from datalog_server_verifier import (
    ServerInfo,
    ServerManager,
    build_protocol_pool,
    _init_driver,
    _login,
)
from datalog_page import (
    DataLogParamConfigPage,
    PostChannelConfig,
    PostChannelPage,
)

log = logging.getLogger(__name__)

# 设备列表缓存文件（避免每次 session 都重新配置 Data Log Parameter Config）
_DEVICE_CACHE = Path(__file__).parent.parent / ".datalog_device_cache.json"


def _load_device_cache() -> list:
    """读取上次已配置的设备列表缓存；文件不存在或解析失败则返回空列表。"""
    try:
        if _DEVICE_CACHE.exists():
            return json.loads(_DEVICE_CACHE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return []


def _save_device_cache(devices: list):
    """将已配置的设备列表写入缓存文件。"""
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
    """
    从当前 session 的所有选中测试项中收集实际用到的协议集合，
    用于按需启动服务器（避免每次全部启动）。
    """
    needed: set[str] = set()
    for item in request.session.items:
        if not hasattr(item, 'callspec'):
            continue
        params = item.callspec.params
        if 'case' in params:
            case = params['case']
            if hasattr(case, 'protocol'):
                needed.add(case.protocol)
        if 'check_protos' in params:
            needed.update(params['check_protos'])
    return needed


@pytest.fixture(scope="session")
def servers(pool, request):
    """
    按测试用例实际需要，只启动对应协议的服务器。
    例：运行 case04（FTP）时只启动 FTP 服务器，不启动 SFTP / HTTP。
    """
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
def driver(pool, servers):
    """
    登录一次，整个 session 复用同一个 WebDriver 实例。

    一次性预置：
      1. Post Channel 1=FTP / 2=SFTP / 3=HTTP：配置后 Disable 保存（配置保留）。
         各用例执行前只需 pc_page.enable_channel(n) 启用对应通道。
      2. Data Log Parameter Config：为所有设备选全部参数类型，Save。
    """
    d = _init_driver(config.GATEWAY_WEB_URL)
    _login(d, config.GATEWAY_WEB_USER, config.GATEWAY_WEB_PASS)

    # ── 1. Post Channel 一次性配置（配置完成后 Disable） ──────────────────────
    pc_page = PostChannelPage(d)
    channel_cfgs: dict[int, PostChannelConfig] = {}
    if "FTP" in pool:
        channel_cfgs[1] = pool["FTP"].to_post_channel_config()
    if "SFTP" in pool:
        channel_cfgs[2] = pool["SFTP"].to_post_channel_config()
    if "HTTP" in pool:
        channel_cfgs[3] = pool["HTTP"].to_post_channel_config()
    if channel_cfgs:
        # enabled=False：填写配置 → Save（Enable） → 再 Disable Save
        # test=False：不做 Test Post Channel（服务器按需启动，部分可能未启动）
        pc_page.configure_all(channel_cfgs, enabled=False, test=False)
        log.info("Post Channel 1/2/3 配置完成（已 Disable，等待各用例按需启用）")

    # ── 2. Data Log Parameter Config 一次性配置（设备列表不变则跳过） ────────
    try:
        param_page = DataLogParamConfigPage(d)
        param_page.navigate()
        current_devices = param_page._get_dropdown_options(param_page.DEVICE_SELECT)
        cached_devices = _load_device_cache()

        if current_devices and set(current_devices) == set(cached_devices):
            log.info("Data Log Parameter Config：设备列表未变（%s），跳过配置", current_devices)
        elif current_devices:
            log.info("Data Log Parameter Config：发现 %d 台设备，开始配置：%s",
                     len(current_devices), current_devices)
            for device_name in current_devices:
                # 不重复进入页面，直接在当前页面切换 Device 下拉
                param_page.configure_device(device_name, [])
            _save_device_cache(current_devices)
            log.info("Data Log Parameter Config 配置完成")
        else:
            log.warning("Data Log Parameter Config：Device 下拉无可用设备，跳过")
    except Exception as e:
        log.warning("Data Log Parameter Config 配置跳过：%s", e)

    yield d
    d.quit()


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


# ─── 工具函数（供测试文件直接 import） ────────────────────────────────────────

def collect_files(dirs: list[str], exts=(".json", ".csv")) -> list[str]:
    """收集指定目录中的所有匹配文件路径。"""
    result = []
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for fn in os.listdir(d):
            if fn.lower().endswith(exts):
                result.append(os.path.join(d, fn))
    return result
