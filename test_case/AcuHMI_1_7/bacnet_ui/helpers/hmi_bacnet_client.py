# -*- coding: utf-8 -*-
"""
hmi_bacnet_client.py — HMI 1-7 BACnet/IP 客户端辅助模块

为 UI 自动化测试提供同步接口，内部用 asyncio.run() 包装 BAC0 异步调用。
连接参数统一从 test_case/AcuHMI_1_7/config.py 读取。

典型用法：
    from test_case.AcuHMI_1_7.bacnet_ui.helpers.hmi_bacnet_client import (
        can_connect, get_device_name, get_object_count
    )

    assert can_connect()                        # BACnet 服务可达
    assert get_object_count() > 0               # 有参数正在发布
    assert can_connect(gateway_port=49000)      # 切换端口后仍可达
"""

from __future__ import annotations

import asyncio
import logging
from typing import Optional

import BAC0

from test_case.AcuHMI_1_7.config import (
    HMI_IP,
    BACNET_PORT,
    LOCAL_IP,
    LOCAL_PORT,
    BACNET_CONNECT_WAIT,
    BACNET_READ_TIMEOUT,
    BACNET_RESTART_WAIT,
)

log = logging.getLogger(__name__)
BAC0.log_level("error")

# 对外别名：测试文件通过这两个名称导入，保持向后兼容
HMI_DEFAULT_PORT    = BACNET_PORT
DEVICE_RESTART_WAIT = BACNET_RESTART_WAIT

# BACnet 广播实例号：无需预知网关真实 Device Instance 即可通信
_BROADCAST_INSTANCE = 4194303


# ─────────────────────────────────────────────────────────────────────────────
# 内部异步实现（不对外暴露）
# ─────────────────────────────────────────────────────────────────────────────

async def _async_get_device_name(gateway_ip: str, gateway_port: int) -> Optional[str]:
    """异步：读取网关 Device Object Name，失败或超时返回 None。"""
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)
            name = await asyncio.wait_for(
                bacnet.read(f"{gw} device {_BROADCAST_INSTANCE} objectName"),
                timeout=BACNET_READ_TIMEOUT,
            )
            return str(name) if name is not None else None
    except Exception as exc:
        log.debug("get_device_name 失败 [%s:%d]: %s", gateway_ip, gateway_port, exc)
        return None


async def _async_get_object_count(gateway_ip: str, gateway_port: int) -> Optional[int]:
    """
    异步：读取 objectList[0]（网关对象总数）。
    返回整数计数；网关不可达或未启用 BACnet 时返回 None。
    """
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)
            count_raw = await asyncio.wait_for(
                bacnet.read(f"{gw} device {_BROADCAST_INSTANCE} objectList",
                            arr_index=0),
                timeout=BACNET_READ_TIMEOUT,
            )
            return int(count_raw) if count_raw is not None else None
    except Exception as exc:
        log.debug("get_object_count 失败 [%s:%d]: %s", gateway_ip, gateway_port, exc)
        return None


# ─────────────────────────────────────────────────────────────────────────────
# 公共同步接口
# ─────────────────────────────────────────────────────────────────────────────

def get_device_name(
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
) -> Optional[str]:
    """
    连接 BACnet 网关并读取 Device Object Name。

    Returns:
        名称字符串；网关不可达、BACnet 未启用则返回 None。
    """
    return asyncio.run(_async_get_device_name(gateway_ip, gateway_port))


def can_connect(
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
) -> bool:
    """
    检查 BACnet 网关是否可达（Device Object Name 可读）。

    适用场景：
    - 验证 BACnet Enable/Disable 状态
    - 验证 BACnet Port 变更后新端口可正常通信
    """
    return get_device_name(gateway_ip, gateway_port) is not None


def get_object_count(
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
) -> Optional[int]:
    """
    读取网关 objectList[0]，返回当前 BACnet 对象总数。

    适用场景：
    - EPICS Enable 前后对比对象数量变化，验证参数已发布/已撤销
    - 值为 None 表示网关不可达

    Note: objectList[0] 包含所有对象（Device Object + AI/BI），
          启用一个参数的 EPICS 会使计数增加 1。
    """
    return asyncio.run(_async_get_object_count(gateway_ip, gateway_port))
