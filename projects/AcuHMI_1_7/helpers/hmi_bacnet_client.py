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
        read_device_info, read_object_metadata_batch,
        run_protocol_compliance, check_stability,
    )

    assert can_connect()                        # BACnet 服务可达
    assert get_object_count() > 0               # 有参数正在发布
    assert can_connect(gateway_port=49000)      # 切换端口后仍可达

    before = get_object_identifiers()           # [(类型, 实例号), ...]
    ...  # UI 上开启某参数 EPICS Enable
    after = get_object_identifiers()
    new = [o for o in after if o not in set(before)]
    read_object_details(new)                    # [(类型, 实例号, objectName, presentValue), ...]

六段式比对辅助函数（第 3/5/6/7 段，无需 UI）：
    dev_info = read_device_info()               # Device Object 12 项必需属性
    meta = read_object_metadata_batch(objects)  # [(类型,实例,key,desc,unit), ...]
    compliance = run_protocol_compliance(probe) # [ProtocolCheckItem(...), ...]
    stability  = check_stability(probe)         # StabilityCheckResult(...)
"""

from __future__ import annotations

import asyncio
import logging
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
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
    BACNET_SERVICE_READY_TIMEOUT,
)
import time as _time

log = logging.getLogger(__name__)
BAC0.log_level("error")

# ─────────────────────────────────────────────────────────────────────────────
# BACnet 单位枚举 → 模板字符串映射（独立副本，避免 import bacnet_reader 时污染 sys.path）
# 与 tools/Protocols/BACnetIP/bacnet_reader.py 中的 _BACNET_UNIT_MAP 保持同步。
# ─────────────────────────────────────────────────────────────────────────────

_BACNET_UNIT_MAP: dict[str, str] = {
    "hertz":                          "Hz",
    "volts":                          "V",
    "amperes":                        "A",
    "milliamperes":                   "mA",
    "kilowatts":                      "kW",
    "kilovolt-amperes":               "kVA",
    "kilovolt-amperes-reactive":      "kvar",
    "percent":                        "%",
    "kilowatt-hours":                 "kWh",
    "kilovolt-ampere-hours":          "kVAh",
    "kilovolt-ampere-hours-reactive": "kvarh",
    "degrees-angular":                "°",
    "degrees-phase":                  "°",   # BACnet 相位角单位，等价于 °
    "no-units":                       "",
}

# 单位等价对（两两互等，°/deg、大小写等常见等价写法，与 comparator.py 中 MetaCheckResult 一致）
_UNIT_EQUIV_PAIRS: list[tuple[str, str]] = [
    ("°", "deg"),
    ("°", "degrees"),
    ("kvar", "kVAr"),
    ("kvarh", "kVArh"),
    ("kVAh", "kVAH"),
]


def _unit_to_str(raw: str) -> str:
    """将 BAC0 返回的 BACnet 单位枚举名转换为模板中使用的单位字符串。"""
    return _BACNET_UNIT_MAP.get(raw.lower().strip(), raw)


def _units_equivalent(a: str, b: str) -> bool:
    """判断两个单位字符串是否等价（完全相等 + 等价对 + 大小写不敏感）。"""
    if a == b:
        return True
    if a.lower() == b.lower():
        return True
    for x, y in _UNIT_EQUIV_PAIRS:
        if (a == x and b == y) or (a == y and b == x):
            return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# 六段式比对辅助数据结构（第 3/5/6/7 段）
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class DeviceInfoResult:
    """BACnet Device Object 属性（ANSI/ASHRAE 135 §12.11 必需属性子集）。"""
    ok: bool = False
    error: str = ""
    object_identifier: str = ""
    object_name: str = ""
    system_status: str = ""
    vendor_name: str = ""
    vendor_id: str = ""
    model_name: str = ""
    firmware_revision: str = ""
    app_sw_version: str = ""
    protocol_version: str = ""
    protocol_revision: str = ""
    max_apdu_length: str = ""
    segmentation: str = ""


@dataclass
class MetadataItem:
    """单个对象的元数据读取结果（第 3 段用）。

    unit_ok:          True = 单位匹配（或模板单位为空已跳过、或 units 读取失败而不判错）。
    unit_skipped:     True = 模板单位为空，跳过本项单位比对（不计入 FAIL）。
    unit_read_failed: True = units 属性读取失败（超时/异常），未拿到值，
                      不能据此判定与模板不符，故不计入 FAIL（与"网关明确无单位"区分）。
    """
    obj_type: str
    instance: int
    param_key: str
    tmpl_unit: str
    bacnet_unit: str
    unit_ok: bool
    unit_skipped: bool = False
    tmpl_desc: str = ""
    bacnet_desc: str = ""
    unit_read_failed: bool = False


@dataclass
class ProtocolCheckItem:
    """单项协议合规性测试结果（第 6 段用）。"""
    test_name: str
    passed: bool
    detail: str = ""


@dataclass
class StabilityCheckResult:
    """连接稳定性测试结果（第 7 段用）。"""
    attempts: int
    successes: int
    errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        """全部成功才算通过。"""
        return self.successes == self.attempts


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

# 对外别名：测试文件通过这些名称导入，保持向后兼容
HMI_DEFAULT_PORT     = BACNET_PORT
DEVICE_RESTART_WAIT  = BACNET_RESTART_WAIT
SERVICE_READY_TIMEOUT = BACNET_SERVICE_READY_TIMEOUT

# BACnet 广播实例号：无需预知网关真实 Device Instance 即可通信
_BROADCAST_INSTANCE = 4194303

# objectList 逐索引回退读取时的并发请求上限（网关不支持分段时生效）
_OBJECT_LIST_CONCURRENCY = 10

# 批量读取对象明细（objectName + presentValue）时每批并发请求数
_DETAIL_BATCH_SIZE = 10

# 元数据（units/description）单属性读取失败后的重试次数与间隔（秒）。
# 并发批量读时个别对象的属性请求可能因网关高并发响应慢而超时，重试可救回，
# 避免把瞬时读取失败误判为"单位为空/不符"。
_META_READ_RETRIES = 2
_META_READ_RETRY_DELAY = 0.5

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

    先读 objectList[0] 取得权威对象总数，再尝试整表读取（1 次请求）：
    仅当整表结果长度与总数一致时才采信（分段设备快速路径）。网关不支持分段时
    （segmentationSupported=no-segmentation），整表读取会被网关截断为首个 APDU
    能容纳的部分对象——BAC0 仅打印 segmentation-not-supported 后返回这截断列表，
    若直接采信会漏掉排在后面的 AI/BI 对象。此时（整表失败或被截断）一律回退为
    按 arr_index 逐元素读取，补齐完整列表。网关不可达返回 None。
    """
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    _quiet_bac0_loggers()  # BAC0 实例化会注册/重设内部 logger，须再次压制
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)
            # 1) 先读权威对象总数（单元素读取，永不触发分段）
            count_raw = await asyncio.wait_for(
                bacnet.read(f"{gw} device {_BROADCAST_INSTANCE} objectList",
                            arr_index=0),
                timeout=BACNET_READ_TIMEOUT,
            )
            if count_raw is None:
                return None
            count = int(count_raw)

            # 2) 尝试整表读取（分段设备快速路径）：仅当结果完整（长度==总数）才采信。
            #    不支持分段的网关只返回首个 APDU 的部分对象，长度 < count，须回退。
            try:
                raw = await asyncio.wait_for(
                    bacnet.read(f"{gw} device {_BROADCAST_INSTANCE} objectList"),
                    timeout=BACNET_READ_TIMEOUT,
                )
                if raw is not None and len(raw) >= count:
                    return [(str(obj_type), int(inst)) for obj_type, inst in raw]
                got = 0 if raw is None else len(raw)
                log.debug("objectList 整表读取不完整（%d/%d，疑似网关不支持分段被截断），"
                          "回退逐索引读取 [%s]", got, count, gw)
            except Exception as exc:
                log.debug("objectList 整表读取失败，回退逐索引读取 [%s]: %s", gw, exc)

            # 3) 回退：按索引并发读取（信号量限制并发请求数，避免压垮网关）。
            #    单个索引读取失败不中断整体（return_exceptions），仅丢弃该项。
            sem = asyncio.Semaphore(_OBJECT_LIST_CONCURRENCY)

            async def _read_index(idx: int) -> Any:
                async with sem:
                    return await asyncio.wait_for(
                        bacnet.read(f"{gw} device {_BROADCAST_INSTANCE} objectList",
                                    arr_index=idx),
                        timeout=BACNET_READ_TIMEOUT,
                    )

            raw_items = await asyncio.gather(
                *(_read_index(idx) for idx in range(1, count + 1)),
                return_exceptions=True,
            )
            idents: list[tuple[str, int]] = []
            failed = 0
            for item in raw_items:
                if isinstance(item, BaseException) or item is None:
                    failed += 1
                    continue
                obj_type, inst = item
                idents.append((str(obj_type), int(inst)))
            if failed:
                log.warning("objectList 逐索引读取：%d/%d 个索引读取失败（已跳过）[%s]",
                            failed, count, gw)
            return idents
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


def wait_until_connectable(
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
    timeout: float = SERVICE_READY_TIMEOUT,
) -> bool:
    """
    保存配置触发网关 BACnet 服务重启后，轮询等待服务重新可达。

    隔离单设备 + 开启全量 Polling 后，网关需重建大量对象，重启常超过固定
    BACNET_RESTART_WAIT；过早读取会得到 TimeoutError（错误信息为空），被误判为
    "客户端无法连接"而 skip。本函数循环调用 can_connect() 直到成功或超过 ``timeout``。

    每次尝试本身已内含 BACNET_CONNECT_WAIT 建链等待 + 读超时，故无需额外 sleep。

    Returns:
        True  — 在预算内服务恢复可达；
        False — 直到 ``timeout`` 仍不可达（此时大概率为网关侧真实异常，调用方应失败而非跳过）。
    """
    deadline = _time.monotonic() + timeout
    attempt = 0
    while True:
        attempt += 1
        if can_connect(gateway_ip, gateway_port):
            log.info("BACnet 服务已可达 [%s:%d]（第 %d 次尝试）",
                     gateway_ip, gateway_port, attempt)
            return True
        if _time.monotonic() >= deadline:
            log.warning("BACnet 服务在 %.0fs 内仍不可达 [%s:%d]，共尝试 %d 次",
                        timeout, gateway_ip, gateway_port, attempt)
            return False


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


# ─────────────────────────────────────────────────────────────────────────────
# 六段式比对：第 5 段 — Device Object 12 项标准必需属性
# ─────────────────────────────────────────────────────────────────────────────

async def _async_read_device_info(
    gateway_ip: str,
    gateway_port: int,
) -> DeviceInfoResult:
    """异步：读取 Device Object 标准必需属性（ANSI/ASHRAE 135 §12.11）。"""
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    _quiet_bac0_loggers()
    result = DeviceInfoResult()
    props: dict[str, str] = {
        "objectIdentifier":           "object_identifier",
        "objectName":                 "object_name",
        "systemStatus":               "system_status",
        "vendorName":                 "vendor_name",
        "vendorIdentifier":           "vendor_id",
        "modelName":                  "model_name",
        "firmwareRevision":           "firmware_revision",
        "applicationSoftwareVersion": "app_sw_version",
        "protocolVersion":            "protocol_version",
        "protocolRevision":           "protocol_revision",
        "maxApduLengthAccepted":      "max_apdu_length",
        "segmentationSupported":      "segmentation",
    }
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)
            fetched = 0
            for prop, attr in props.items():
                try:
                    val = await asyncio.wait_for(
                        bacnet.read(f"{gw} device {_BROADCAST_INSTANCE} {prop}"),
                        timeout=BACNET_READ_TIMEOUT,
                    )
                    if val is not None:
                        setattr(result, attr, str(val))
                        fetched += 1
                except Exception as exc:
                    log.debug("Device Object 属性 %s 读取失败: %s", prop, exc)
            result.ok = fetched > 0
            if not result.ok:
                result.error = "无法读取任何 Device Object 属性"
            log.info("Device Object 属性：读取 %d/%d 项", fetched, len(props))
    except Exception as exc:
        result.error = str(exc)
        log.debug("read_device_info 失败 [%s:%d]: %s", gateway_ip, gateway_port, exc)
    return result


def read_device_info(
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
) -> DeviceInfoResult:
    """
    读取网关 Device Object 12 项标准必需属性（ANSI/ASHRAE 135 §12.11）。

    Returns:
        DeviceInfoResult；ok=False 时 error 字段说明原因。
    """
    return _run_coro(_async_read_device_info(gateway_ip, gateway_port))


# ─────────────────────────────────────────────────────────────────────────────
# 六段式比对：第 3 段 — 元数据检查（description / units 属性 vs 模板）
# ─────────────────────────────────────────────────────────────────────────────

async def _async_read_object_metadata_batch(
    gateway_ip: str,
    gateway_port: int,
    objects: Sequence[tuple[str, int, str, str, str]],
) -> list[MetadataItem]:
    """
    异步：批量读取 AI/BI 对象的 description 和 units 属性，与模板元数据对比。

    Args:
        objects: [(obj_type, instance, param_key, tmpl_unit, tmpl_desc), ...]

    Returns:
        list[MetadataItem]，顺序与 objects 一致。
    """
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    _quiet_bac0_loggers()
    results: list[MetadataItem] = []
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)

            async def _read_one(
                obj_type: str,
                inst: int,
                param_key: str,
                tmpl_unit: str,
                tmpl_desc: str,
            ) -> MetadataItem:
                async def _read_attr(attr: str) -> tuple[Any, bool]:
                    """读取单个属性，失败重试 _META_READ_RETRIES 次。

                    Returns:
                        (value, read_failed)：read_failed=True 表示重试用尽仍未读到。
                    """
                    for attempt in range(_META_READ_RETRIES + 1):
                        try:
                            value = await asyncio.wait_for(
                                bacnet.read(f"{gw} {obj_type} {inst} {attr}"),
                                timeout=BACNET_READ_TIMEOUT,
                            )
                            return value, False
                        except Exception as exc:
                            if attempt < _META_READ_RETRIES:
                                await asyncio.sleep(_META_READ_RETRY_DELAY)
                            else:
                                log.debug("%s 读取失败 [%s %d]（已重试 %d 次）: %s",
                                          attr, obj_type, inst, _META_READ_RETRIES, exc)
                    return None, True

                desc_raw, _ = await _read_attr("description")
                units_raw, units_failed = await _read_attr("units")
                bacnet_desc = str(desc_raw).strip() if desc_raw is not None else ""
                bacnet_unit = (
                    _unit_to_str(str(units_raw)) if units_raw is not None else ""
                )
                # 诊断：units 读取成功（非超时）但映射后为空、而模板要求有单位时，
                # 记录网关返回的原始值——用于区分"读取超时(None)"、"网关返回 no-units"、
                # "返回了枚举名但未被 _BACNET_UNIT_MAP 覆盖"三种情况，便于定位根因。
                if (not units_failed and units_raw is not None
                        and bacnet_unit == "" and tmpl_unit != ""):
                    log.warning(
                        "[元数据] units 读到但映射为空 [%s %d] param=%s "
                        "模板单位=%r 网关units原始值=%r",
                        obj_type, inst, param_key, tmpl_unit, units_raw,
                    )
                # 模板单位为空时跳过比对（不计入 FAIL）；units 读取失败时同样不判错
                # （未拿到值，不能据此断言与模板不符），仅对"成功读到单位"的项做严格比对。
                unit_skipped = tmpl_unit == ""
                unit_ok = (
                    unit_skipped or units_failed
                    or _units_equivalent(tmpl_unit, bacnet_unit)
                )
                return MetadataItem(
                    obj_type=obj_type,
                    instance=inst,
                    param_key=param_key,
                    tmpl_unit=tmpl_unit,
                    bacnet_unit=bacnet_unit,
                    unit_ok=unit_ok,
                    unit_skipped=unit_skipped,
                    tmpl_desc=tmpl_desc,
                    bacnet_desc=bacnet_desc,
                    unit_read_failed=units_failed,
                )

            total = len(objects)
            for start in range(0, total, _DETAIL_BATCH_SIZE):
                batch = list(objects[start: start + _DETAIL_BATCH_SIZE])
                batch_res = await asyncio.gather(
                    *(_read_one(ot, inst, pk, tu, td) for ot, inst, pk, tu, td in batch)
                )
                results.extend(batch_res)
                done = min(start + _DETAIL_BATCH_SIZE, total)
                log.info("元数据读取进度 %d/%d", done, total)
    except Exception as exc:
        log.debug("read_object_metadata_batch 失败 [%s:%d]: %s",
                  gateway_ip, gateway_port, exc)
    return results


def read_object_metadata_batch(
    objects: Sequence[tuple[str, int, str, str, str]],
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
) -> list[MetadataItem]:
    """
    批量读取已发布对象的 description / units 属性并与模板对比（第 3 段）。

    Args:
        objects: [(obj_type, instance, param_key, tmpl_unit, tmpl_desc), ...]
                 通常由调用方从 get_object_identifiers() + 模板映射构建。

    Returns:
        list[MetadataItem]；unit_ok=False 的项表示单位不一致。
    """
    return _run_coro(
        _async_read_object_metadata_batch(gateway_ip, gateway_port, objects)
    )


# ─────────────────────────────────────────────────────────────────────────────
# 六段式比对：第 6 段 — 协议合规性（§16 错误响应 + AI 必需属性）
# ─────────────────────────────────────────────────────────────────────────────

async def _async_run_protocol_compliance(
    gateway_ip: str,
    gateway_port: int,
    probe_obj: Optional[tuple[str, int]],
) -> list[ProtocolCheckItem]:
    """
    异步：协议合规性测试。
      1. 非法 AI 对象请求（Instance=9999999）应返回错误（None）
      2. 若提供 probe_obj：statusFlags / outOfService / units 三个 AI 必需属性可读
    """
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    _quiet_bac0_loggers()
    tests: list[ProtocolCheckItem] = []
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)

            # 1. 非法对象实例
            try:
                val = await asyncio.wait_for(
                    bacnet.read(f"{gw} analog-input 9999999 presentValue"),
                    timeout=BACNET_READ_TIMEOUT,
                )
            except Exception:
                val = None
            tests.append(ProtocolCheckItem(
                test_name="读取不存在 AI 对象（Instance=9999999）应返回错误",
                passed=(val is None),
                detail="返回 None，符合 §16 标准" if val is None
                       else f"未拒绝，意外返回：{val}",
            ))

            # 2–4. AI 必需属性（§12.2.2）
            if probe_obj is not None:
                obj_type, inst = probe_obj
                ref = f"{obj_type} {inst}"
                for prop in ("statusFlags", "outOfService", "units"):
                    try:
                        pval = await asyncio.wait_for(
                            bacnet.read(f"{gw} {ref} {prop}"),
                            timeout=BACNET_READ_TIMEOUT,
                        )
                    except Exception:
                        pval = None
                    tests.append(ProtocolCheckItem(
                        test_name=f"AI 必需属性 {prop}（{obj_type},{inst}）可读",
                        passed=(pval is not None),
                        detail=str(pval) if pval is not None else "读取失败",
                    ))

        log.info("协议合规性测试：%d/%d 通过",
                 sum(1 for t in tests if t.passed), len(tests))
    except Exception as exc:
        log.debug("run_protocol_compliance 失败 [%s:%d]: %s",
                  gateway_ip, gateway_port, exc)
    return tests


def run_protocol_compliance(
    probe_obj: Optional[tuple[str, int]] = None,
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
) -> list[ProtocolCheckItem]:
    """
    执行 BACnet 协议合规性测试（第 6 段，参考 ANSI/ASHRAE 135 §16 + §12.2.2）。

    Args:
        probe_obj: (obj_type, instance) 用于读 AI 必需属性，如 ("analogInput", 1)；
                   None 时只执行非法请求测试。

    Returns:
        list[ProtocolCheckItem]；passed=False 的项表示合规性问题。
    """
    return _run_coro(
        _async_run_protocol_compliance(gateway_ip, gateway_port, probe_obj)
    )


# ─────────────────────────────────────────────────────────────────────────────
# 六段式比对：第 7 段 — 连接稳定性（同一对象连续读 N 次）
# ─────────────────────────────────────────────────────────────────────────────

async def _async_check_stability(
    gateway_ip: str,
    gateway_port: int,
    probe_obj: tuple[str, int],
    attempts: int,
    delay: float,
) -> StabilityCheckResult:
    """异步：对同一 AI/BI 对象连续读取 attempts 次 presentValue，统计成功率。"""
    gw = f"{gateway_ip}:{gateway_port}"
    bacnet = BAC0.lite(ip=LOCAL_IP, port=LOCAL_PORT)
    _quiet_bac0_loggers()
    obj_type, inst = probe_obj
    successes = 0
    errors: list[str] = []
    try:
        async with bacnet:
            await asyncio.sleep(BACNET_CONNECT_WAIT)
            for i in range(attempts):
                if i > 0:
                    await asyncio.sleep(delay)
                try:
                    raw = await asyncio.wait_for(
                        bacnet.read(f"{gw} {obj_type} {inst} presentValue"),
                        timeout=BACNET_READ_TIMEOUT,
                    )
                    if raw is not None:
                        successes += 1
                    else:
                        errors.append(f"第{i + 1}次：返回 None")
                except Exception as exc:
                    errors.append(f"第{i + 1}次：{exc}")
        log.info("稳定性测试：%d/%d 成功  对象=%s,%d",
                 successes, attempts, obj_type, inst)
    except Exception as exc:
        errors.append(f"连接失败：{exc}")
        log.debug("check_stability 失败 [%s:%d]: %s", gateway_ip, gateway_port, exc)
    return StabilityCheckResult(attempts=attempts, successes=successes, errors=errors)


def check_stability(
    probe_obj: tuple[str, int],
    attempts: int = 5,
    delay: float = 0.5,
    gateway_ip: str = HMI_IP,
    gateway_port: int = HMI_DEFAULT_PORT,
) -> StabilityCheckResult:
    """
    连接稳定性测试：对同一 AI/BI 对象连续读取 attempts 次（第 7 段）。

    Args:
        probe_obj: (obj_type, instance)，如 ("analogInput", 1)。
        attempts:  读取次数（默认 5）。
        delay:     两次读取间隔秒数（默认 0.5s）。

    Returns:
        StabilityCheckResult；ok=True 表示全部成功。
    """
    return _run_coro(
        _async_check_stability(gateway_ip, gateway_port, probe_obj, attempts, delay)
    )
