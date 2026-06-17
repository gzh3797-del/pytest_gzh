# -*- coding: utf-8 -*-
"""
hmi_bacnet_client.py — HMI 1-7 BACnet/IP 客户端辅助模块

为 UI 自动化测试提供同步接口，内部以独立事件循环包装 BAC0 异步调用
（自动兼容 Playwright sync API 的运行中事件循环环境）。
连接参数统一从 projects/AcuHMI_1_7/settings.py 读取。

典型用法：
    from projects.AcuHMI_1_7.helpers.hmi_bacnet_client import (
        can_connect, get_device_name, get_object_count,
        get_object_identifiers, read_object_details,
    )

    assert can_connect()                        # BACnet 服务可达
    assert get_object_count() > 0               # 有参数正在发布
    assert can_connect(gateway_port=49000)      # 切换端口后仍可达

    before = get_object_identifiers()           # [(类型, 实例号), ...]
    ...  # UI 上开启某参数 EPICS Enable
    after = get_object_identifiers()
    new = [o for o in after if o not in set(before)]
    read_object_details(new)                    # [(类型, 实例号, objectName, presentValue), ...]
"""

from __future__ import annotations

import asyncio
import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Coroutine, Optional, Sequence, TypeVar

import BAC0

from projects.AcuHMI_1_7.settings import (
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


def _quiet_bac0_loggers() -> None:
    """将 BAC0/bacpypes3/asyncio 全部已注册 logger 压到 WARNING。

    BAC0.log_level() 只调低 BAC0 自身 handler 的输出级别；BAC0 为每个类
    单独创建 logger 并显式设置 DEBUG 级别（子 logger 有自己的级别时父
    logger 的设置不生效），pytest 的日志捕获挂在 logger 层，逐索引读取
    objectList 时会把数万行内部 DEBUG 写进报告，淹没用例自身的 INFO 日志。
    BAC0.lite() 实例化时还可能重设级别，因此每次建连后须重新调用本函数。
    """
    for name in list(logging.root.manager.loggerDict):
        if name.startswith(("BAC0", "bacpypes3", "asyncio")):
            logging.getLogger(name).setLevel(logging.WARNING)


_quiet_bac0_loggers()

# 对外别名：测试文件通过这两个名称导入，保持向后兼容
HMI_DEFAULT_PORT    = BACNET_PORT
DEVICE_RESTART_WAIT = BACNET_RESTART_WAIT

# BACnet 广播实例号：无需预知网关真实 Device Instance 即可通信
_BROADCAST_INSTANCE = 4194303

# objectList 逐索引回退读取时的并发请求上限（网关不支持分段时生效）
_OBJECT_LIST_CONCURRENCY = 10

# 批量读取对象明细（objectName + presentValue）时每批并发请求数
_DETAIL_BATCH_SIZE = 10

_T = TypeVar("_T")


def _run_coro(coro: Coroutine[Any, Any, _T]) -> _T:
    """同步执行协程，兼容 Playwright sync API 场景。

    Playwright sync API 在主线程持有一个运行中的 asyncio 事件循环
    （greenlet 驱动），此时直接 asyncio.run() 会抛
    "asyncio.run() cannot be called from a running event loop"。
    检测到运行中循环时，改在独立线程中以全新事件循环执行。
    """
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        # 当前线程没有运行中的事件循环，直接运行
        return asyncio.run(coro)
    # 已有运行中循环（如 Playwright sync API）→ 在独立线程中运行
    with ThreadPoolExecutor(max_workers=1) as executor:
        return executor.submit(asyncio.run, coro).result()


# ─────────────────────────────────────────────────────────────────────────────
# 内部异步实现（不对外暴露）
# ─────────────────────────────────────────────────────────────────────────────

async def _async_get_device_name(gateway_ip: str, gateway_port: int) -> Optional[str]:
    """异步：读取网关 Device Object Name，失败或超时返回 None。"""
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    _quiet_bac0_loggers()  # BAC0 实例化会注册/重设内部 logger，须再次压制
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
    _quiet_bac0_loggers()  # BAC0 实例化会注册/重设内部 logger，须再次压制
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


async def _async_get_object_identifiers(
    gateway_ip: str,
    gateway_port: int,
) -> Optional[list[tuple[str, int]]]:
    """
    异步：读取网关完整 objectList，返回 [(对象类型, 实例号), ...]。

    优先整表读取（1 次请求）；网关不支持分段且对象过多导致整表读取失败时，
    回退为按 arr_index 逐元素读取。网关不可达返回 None。
    """
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    _quiet_bac0_loggers()  # BAC0 实例化会注册/重设内部 logger，须再次压制
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)
            # 优先整表读取
            try:
                raw = await asyncio.wait_for(
                    bacnet.read(f"{gw} device {_BROADCAST_INSTANCE} objectList"),
                    timeout=BACNET_READ_TIMEOUT,
                )
                if raw is not None:
                    return [(str(obj_type), int(inst)) for obj_type, inst in raw]
            except Exception as exc:
                log.debug("objectList 整表读取失败，回退逐索引读取 [%s]: %s", gw, exc)
            # 回退：先读数量，再按索引并发读取（信号量限制并发请求数，避免压垮网关）
            count_raw = await asyncio.wait_for(
                bacnet.read(f"{gw} device {_BROADCAST_INSTANCE} objectList",
                            arr_index=0),
                timeout=BACNET_READ_TIMEOUT,
            )
            if count_raw is None:
                return None
            sem = asyncio.Semaphore(_OBJECT_LIST_CONCURRENCY)

            async def _read_index(idx: int) -> Any:
                async with sem:
                    return await asyncio.wait_for(
                        bacnet.read(f"{gw} device {_BROADCAST_INSTANCE} objectList",
                                    arr_index=idx),
                        timeout=BACNET_READ_TIMEOUT,
                    )

            raw_items = await asyncio.gather(
                *(_read_index(idx) for idx in range(1, int(count_raw) + 1))
            )
            return [(str(obj_type), int(inst)) for obj_type, inst in raw_items]
    except Exception as exc:
        log.debug("get_object_identifiers 失败 [%s:%d]: %s",
                  gateway_ip, gateway_port, exc)
        return None


async def _async_read_object_details(
    gateway_ip: str,
    gateway_port: int,
    objects: Sequence[tuple[str, int]],
) -> list[tuple[str, int, Optional[str], Any]]:
    """
    异步：读取指定对象的 objectName 与 presentValue（同一连接内逐个读取）。

    单条属性读取失败不中断（该项记为 None）；device 对象无 presentValue，跳过。
    返回 [(对象类型, 实例号, objectName, presentValue), ...]。
    """
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    _quiet_bac0_loggers()  # BAC0 实例化会注册/重设内部 logger，须再次压制
    results: list[tuple[str, int, Optional[str], Any]] = []
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)
            for obj_type, inst in objects:
                name: Optional[str] = None
                value: Any = None
                try:
                    name_raw = await asyncio.wait_for(
                        bacnet.read(f"{gw} {obj_type} {inst} objectName"),
                        timeout=BACNET_READ_TIMEOUT,
                    )
                    name = str(name_raw) if name_raw is not None else None
                except Exception as exc:
                    log.debug("读取 objectName 失败 [%s %s,%d]: %s",
                              gw, obj_type, inst, exc)
                if obj_type != "device":
                    try:
                        value = await asyncio.wait_for(
                            bacnet.read(f"{gw} {obj_type} {inst} presentValue"),
                            timeout=BACNET_READ_TIMEOUT,
                        )
                    except Exception as exc:
                        log.debug("读取 presentValue 失败 [%s %s,%d]: %s",
                                  gw, obj_type, inst, exc)
                results.append((obj_type, inst, name, value))
    except Exception as exc:
        log.debug("read_object_details 失败 [%s:%d]: %s",
                  gateway_ip, gateway_port, exc)
    return results


async def _async_read_object_details_batch(
    gateway_ip: str,
    gateway_port: int,
    objects: Sequence[tuple[str, int]],
    batch_size: int,
) -> list[tuple[str, int, Optional[str], Any]]:
    """
    异步：分批并发读取指定对象的 objectName 与 presentValue（同一连接内复用）。

    对象数量多（全量参数验证可达数百个）时，逐个串行读取会非常慢，本函数
    按 batch_size 分批 asyncio.gather 并发，显著缩短总耗时。

    单条属性读取失败不中断（该项记为 None）；device 对象无 presentValue，跳过。
    返回 [(对象类型, 实例号, objectName, presentValue), ...]，顺序与入参一致。
    """
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    _quiet_bac0_loggers()  # BAC0 实例化会注册/重设内部 logger，须再次压制
    results: list[tuple[str, int, Optional[str], Any]] = []
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)

            async def _read_one(
                obj_type: str, inst: int
            ) -> tuple[str, int, Optional[str], Any]:
                name: Optional[str] = None
                value: Any = None
                try:
                    name_raw = await asyncio.wait_for(
                        bacnet.read(f"{gw} {obj_type} {inst} objectName"),
                        timeout=BACNET_READ_TIMEOUT,
                    )
                    name = str(name_raw) if name_raw is not None else None
                except Exception as exc:
                    log.debug("读取 objectName 失败 [%s %s,%d]: %s",
                              gw, obj_type, inst, exc)
                if obj_type != "device":
                    try:
                        value = await asyncio.wait_for(
                            bacnet.read(f"{gw} {obj_type} {inst} presentValue"),
                            timeout=BACNET_READ_TIMEOUT,
                        )
                    except Exception as exc:
                        log.debug("读取 presentValue 失败 [%s %s,%d]: %s",
                                  gw, obj_type, inst, exc)
                return obj_type, inst, name, value

            total = len(objects)
            for start in range(0, total, batch_size):
                batch = objects[start: start + batch_size]
                batch_results = await asyncio.gather(
                    *(_read_one(obj_type, inst) for obj_type, inst in batch)
                )
                results.extend(batch_results)
                done = min(start + batch_size, total)
                log.info("对象明细批量读取进度 %d/%d", done, total)
    except Exception as exc:
        log.debug("read_object_details_batch 失败 [%s:%d]: %s",
                  gateway_ip, gateway_port, exc)
    return results


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
    return _run_coro(_async_get_device_name(gateway_ip, gateway_port))


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
    return _run_coro(_async_get_object_count(gateway_ip, gateway_port))


def get_object_identifiers(
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
) -> Optional[list[tuple[str, int]]]:
    """
    读取网关完整 objectList，返回 [(对象类型, 实例号), ...]。

    适用场景：
    - EPICS Enable 前后对比对象列表差集，定位本次新发布/撤销的具体对象
    - len(返回值) 等价于 get_object_count()，可直接用于数量断言
    - 返回 None 表示网关不可达
    """
    return _run_coro(_async_get_object_identifiers(gateway_ip, gateway_port))


def read_object_details(
    objects: Sequence[tuple[str, int]],
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
) -> list[tuple[str, int, Optional[str], Any]]:
    """
    读取指定对象的 objectName 与 presentValue。

    Args:
        objects: 对象标识列表 [(类型, 实例号), ...]，
                 通常来自 get_object_identifiers() 前后差集

    Returns:
        [(对象类型, 实例号, objectName, presentValue), ...]；
        单条属性读取失败时该项为 None，网关不可达时返回空列表。
    """
    return _run_coro(_async_read_object_details(gateway_ip, gateway_port, objects))


def read_object_details_batch(
    objects: Sequence[tuple[str, int]],
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
    batch_size: int = _DETAIL_BATCH_SIZE,
) -> list[tuple[str, int, Optional[str], Any]]:
    """
    分批并发读取指定对象的 objectName 与 presentValue（read_object_details 的并发版）。

    对象数量多时优先用本函数：按 batch_size 并发，显著快于逐个串行读取。

    Args:
        objects:    对象标识列表 [(类型, 实例号), ...]。
        batch_size: 每批并发请求数（默认 10）。

    Returns:
        [(对象类型, 实例号, objectName, presentValue), ...]，顺序与入参一致；
        单条属性读取失败时该项为 None，网关不可达时返回空列表。
    """
    return _run_coro(
        _async_read_object_details_batch(gateway_ip, gateway_port, objects, batch_size)
    )
