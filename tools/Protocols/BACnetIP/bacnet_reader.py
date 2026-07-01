# -*- coding: utf-8 -*-
"""
bacnet_reader.py — BACnet/IP 读取模块（BAC0 2025.x async 架构）

已验证：
  - 网关 Device Instance: 4194302，Device Name: WEB2_BACNET_GW
  - AI 实例号范围：1 ~ 1059（EPICS 文件中为 0 ~ 1058，差值 +1）
  - 对象名格式：AcuRev4100-{PARAM}，与 EPICS 一致
  - 读取方式：BAC0 async with + await

核心类：
  EPICSParser  — 解析 .tpi 文件，提供参数名/单位参考（实例号需 +1 对齐设备）
  BACnetReader — 异步 BACnet 客户端，动态发现对象列表并批量读取 Present Value

用法：
  import asyncio
  from bacnet_reader import EPICSParser, BACnetReader

  async def main():
      epics = EPICSParser.parse()           # EPICS 参数参考表
      async with BACnetReader() as reader:
          objects = await reader.discover_objects()  # 从设备读取真实对象列表
          results = await reader.read_all(objects)
          for r in results:
              print(r)

  asyncio.run(main())
"""

from __future__ import annotations

import asyncio
import logging
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

import BAC0

import config
from template_reader import TemplateParam, find_template_file, get_bacnet_params, get_bacnet_params_by_range

log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# 数据结构
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
class ProtocolErrorTestResult:
    """单项协议合规性测试结果。"""
    test_name: str
    passed: bool
    detail: str = ""


@dataclass
class StabilityResult:
    """BACnet 连接稳定性测试结果（多次重复读取）。"""
    attempts: int
    successes: int
    errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.successes == self.attempts


@dataclass
class AIObject:
    """一个 BACnet Analog Input 或 Binary Input 对象。"""
    instance:    int   # 设备上的真实实例号（1-based）
    name:        str   # objectName，如 "AcuRev4100-VLN_a_V"
    param_key:   str   # 参数标识，如 "VLN_a_V"
    unit:        str   # 工程单位（来自模板 unit 列，如 "V"、"kW"）
    description: str   # 参数描述（来自模板 descrption 列）
    obj_type:    str = "analogInput"  # BACnet 对象类型："analogInput" 或 "binaryInput"

    def __str__(self) -> str:
        prefix = "BI" if self.obj_type == "binaryInput" else "AI"
        return f"{prefix}[{self.instance}] {self.name}"


@dataclass
class ReadResult:
    """单次读取结果。"""
    obj:   AIObject
    value: Optional[float] = None
    error: str = ""

    @property
    def ok(self) -> bool:
        return not self.error and self.value is not None

    def __str__(self) -> str:
        if self.ok:
            unit = f" {self.obj.unit}" if self.obj.unit else ""
            return f"{self.obj.name} = {self.value}{unit}"
        return f"{self.obj.name} ERROR: {self.error}"


# ─────────────────────────────────────────────────────────────────────────────
# BACnet 单位枚举 → 字符串映射
# ─────────────────────────────────────────────────────────────────────────────

_BACNET_UNIT_MAP: dict[str, str] = {
    "hertz": "Hz",
    "volts": "V",
    "amperes": "A",
    "milliamperes": "mA",
    "kilowatts": "kW",
    "kilovolt-amperes": "kVA",
    "kilovolt-amperes-reactive": "kvar",
    "percent": "%",
    "kilowatt-hours": "kWh",
    "kilovolt-ampere-hours": "kVAh",
    "kilovolt-ampere-hours-reactive": "kvarh",
    "degrees-angular": "°",
    "degrees-phase": "°",   # BACnet 相位角单位，等价于 °
    "no-units": "",
}

# 单位等价对（两两互等，°/deg、大小写等常见等价写法）。
_UNIT_EQUIV_PAIRS: list[tuple[str, str]] = [
    ("°", "deg"),
    ("°", "degrees"),
    ("kvar", "kVAr"),
    ("kvarh", "kVArh"),
    ("kVAh", "kVAH"),
]


def _bacnet_unit_to_str(raw: str) -> str:
    """将 BAC0 返回的 BACnet 单位枚举名转换为模板中使用的单位字符串。"""
    return _BACNET_UNIT_MAP.get(raw.lower().strip(), raw)


def units_equivalent(a: str, b: str) -> bool:
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
# 模板参数查询（替代 EPICS，提供 param_key → TemplateParam 映射）
# ─────────────────────────────────────────────────────────────────────────────

def _load_template_map(device_name: str) -> dict[str, TemplateParam]:
    """加载设备模板，返回 param_key → TemplateParam 字典。失败时返回空字典。"""
    try:
        path = find_template_file(config.TEMPLATE_DIR, device_name)
        range_marker = getattr(config, 'BACNET_RANGE_MARKER', '')
        if range_marker:
            params = get_bacnet_params_by_range(path, range_marker)
        else:
            params = get_bacnet_params(path)
        log.info("模板加载：%d 个 BACnet 参数（%s）", len(params), path)
        return {p.param_key: p for p in params}
    except Exception as exc:
        log.warning("模板加载失败，将使用空字典：%s", exc)
        return {}


# ─────────────────────────────────────────────────────────────────────────────
# EPICSParser（已弃用，仅保留向后兼容）
# ─────────────────────────────────────────────────────────────────────────────

_UNIT_RE = re.compile(
    r"[-_](?P<unit>kWh|kVAh|kVArh|kVAr|kVA|kvarh|kvar|kW|kV|Hz|mA|VAh|Wh|VAr|VA|W|A|V)$"
)


def _extract_unit(param_key: str) -> str:
    m = _UNIT_RE.search(param_key)
    return m.group("unit") if m else ""


class EPICSParser:
    """已弃用：请使用模板文件（Template/*.xlsx）替代 EPICS 文件。"""

    _PATTERN = re.compile(
        r'object-identifier:\s*\(analog-input,\s*(\d+)\)\s*\n'
        r'\s*object-name:\s*"([^"]+)"',
        re.MULTILINE,
    )

    @classmethod
    def parse(cls, filepath: str) -> list[AIObject]:
        with open(filepath, encoding="utf-8", errors="replace") as f:
            content = f.read()
        objects: list[AIObject] = []
        for m in cls._PATTERN.finditer(content):
            instance = int(m.group(1))
            name = m.group(2).strip()
            _, _, param_key = name.partition("-")
            objects.append(AIObject(instance, name, param_key,
                                    unit=_extract_unit(param_key), description=""))
        objects.sort(key=lambda o: o.instance)
        return objects

    @classmethod
    def build_name_map(cls, filepath: str) -> dict[str, str]:
        return {o.param_key: o.unit for o in cls.parse(filepath)}


# ─────────────────────────────────────────────────────────────────────────────
# BACnet 读取器（全 async，使用 BAC0 2025.x）
# ─────────────────────────────────────────────────────────────────────────────

class BACnetReader:
    """
    BACnet/IP 异步客户端。

    以 async context manager 方式使用：
        async with BACnetReader() as reader:
            objects = await reader.discover_objects()
            results = await reader.read_all(objects)

    内部地址格式：f"{GATEWAY_IP}:{GATEWAY_PORT}"
    设备实例：config.DEVICE_INSTANCE（已知为 4194302）
    """

    # BACnet 广播地址（设备实例通配符），用于跨实例号的读取
    _BROADCAST_INSTANCE = 4194303

    def __init__(self) -> None:
        self._bacnet: Optional[BAC0.lite] = None
        self._gw = f"{config.GATEWAY_IP}:{config.GATEWAY_PORT}"
        self._dev = config.DEVICE_INSTANCE or self._BROADCAST_INSTANCE
        # 从模板加载参数字典（param_key → TemplateParam）
        self._tmpl_map: dict[str, TemplateParam] = _load_template_map(config.DEVICE_NAME)

    # ── async context manager ─────────────────────────────────────────────────

    async def __aenter__(self) -> "BACnetReader":
        log.info("启动 BAC0  本地=%s:%d", config.LOCAL_IP, config.LOCAL_PORT)
        BAC0.log_level("error")
        self._bacnet = BAC0.lite(ip=config.LOCAL_IP, port=config.LOCAL_PORT)
        await self._bacnet.__aenter__()
        await asyncio.sleep(config.CONNECT_WAIT)
        log.info("BAC0 就绪  网关=%s  Device=%d", self._gw, self._dev)
        return self

    async def __aexit__(self, *args) -> None:
        if self._bacnet is not None:
            await self._bacnet.__aexit__(*args)
            self._bacnet = None
            log.info("BAC0 已断开")

    # ── 设备发现 ───────────────────────────────────────────────────────────────

    async def get_device_name(self) -> str:
        """读取网关 Device Object Name，用于验证连接目标正确。"""
        name = await self._read_raw(
            f"{self._gw} device {self._BROADCAST_INSTANCE} objectName"
        )
        return str(name) if name is not None else ""

    async def discover_objects(self) -> list[AIObject]:
        """
        从网关读取完整对象列表，返回所有 AI 和 BI 对象。

        流程：
          1. 读 objectList[0] 取总数
          2. 逐索引读 objectList[i] 取 (type, instance) 对
          3. 过滤出 analog-input 和 binary-input 对象
          4. 对每个实例读 objectName，构建 AIObject

        注意：objectList 不支持分段传输，只能按索引读取。
        """
        assert self._bacnet, "未连接"

        # 读对象总数
        count_raw = await self._read_raw(
            f"{self._gw} device {self._dev} objectList", arr_index=0
        )
        if count_raw is None:
            raise RuntimeError("无法读取 objectList，请确认网关 BACnet 服务已启用")
        total = int(count_raw)
        log.info("objectList 总数：%d", total)

        # 读全部对象引用，收集 AI 和 BI 实例号
        ai_instances: list[int] = []
        bi_instances: list[int] = []
        for i in range(1, total + 1):
            ref = await self._read_raw(
                f"{self._gw} device {self._dev} objectList", arr_index=i
            )
            if ref is None:
                continue
            parts = str(ref).split(",")
            if len(parts) == 2:
                obj_type_raw = parts[0].strip()
                inst = int(parts[1].strip())
                if obj_type_raw == "analog-input":
                    ai_instances.append(inst)
                elif obj_type_raw == "binary-input":
                    bi_instances.append(inst)

        if not ai_instances and not bi_instances:
            raise RuntimeError(
                "BACnet objectList 中未发现任何 analog-input 或 binary-input 对象，"
                "请确认网关设备实例号、IP 和端口是否正确，或设备是否在线。"
            )
        log.info("发现 AI 对象：%d 个，BI 对象：%d 个", len(ai_instances), len(bi_instances))

        objects: list[AIObject] = []

        # 读每个 AI 的 objectName，从模板获取 description 和 unit
        for inst in ai_instances:
            name_raw = await self._read_raw(
                f"{self._gw} analogInput {inst} objectName"
            )
            name = str(name_raw) if name_raw is not None else f"AI_{inst}"
            _, _, param_key = name.partition("-")
            tmpl = self._tmpl_map.get(param_key)
            unit        = tmpl.unit        if tmpl else _extract_unit(param_key)
            description = tmpl.description if tmpl else ""
            objects.append(AIObject(inst, name, param_key, unit, description, "analogInput"))

        # 读每个 BI 的 objectName（IOM-03/04 DI 通道）
        for inst in bi_instances:
            name_raw = await self._read_raw(
                f"{self._gw} binaryInput {inst} objectName"
            )
            name = str(name_raw) if name_raw is not None else f"BI_{inst}"
            _, _, param_key = name.partition("-")
            tmpl = self._tmpl_map.get(param_key)
            description = tmpl.description if tmpl else ""
            objects.append(AIObject(inst, name, param_key, "", description, "binaryInput"))

        log.info("BACnet 对象构建完成：%d 个（AI=%d, BI=%d）",
                 len(objects), len(ai_instances), len(bi_instances))
        return objects

    # ── 读取 Present Value ────────────────────────────────────────────────────

    async def read_present_value(self, obj: AIObject) -> ReadResult:
        """读取单个 AI/BI Present Value，失败自动重试。"""
        obj_type_str = obj.obj_type  # "analogInput" 或 "binaryInput"
        last_err = ""
        for attempt in range(config.MAX_RETRIES + 1):
            try:
                raw = await asyncio.wait_for(
                    self._bacnet.read(
                        f"{self._gw} {obj_type_str} {obj.instance} presentValue"
                    ),
                    timeout=config.READ_TIMEOUT,
                )
                if raw is None:
                    raise ValueError("返回 None")
                if obj_type_str == "binaryInput":
                    # BAC0 对 BI 返回 True/False 或字符串 'active'/'inactive'
                    if isinstance(raw, bool):
                        value = 1.0 if raw else 0.0
                    else:
                        value = 1.0 if str(raw).lower().strip() == "active" else 0.0
                else:
                    value = float(raw)
                return ReadResult(obj=obj, value=value)
            except Exception as exc:
                last_err = str(exc)
                if attempt < config.MAX_RETRIES:
                    log.debug("%s[%d] 第%d次重试：%s",
                              obj_type_str, obj.instance, attempt + 1, last_err)
                    await asyncio.sleep(config.RETRY_WAIT)

        return ReadResult(obj=obj, value=None, error=last_err)

    async def read_all(
        self,
        objects: list[AIObject],
        progress_cb=None,
    ) -> list[ReadResult]:
        """
        顺序读取所有 AI Present Value。

        Args:
            objects:     AIObject 列表（来自 discover_objects()）
            progress_cb: 可选异步或同步回调 f(done, total)

        Returns:
            ReadResult 列表，顺序与 objects 一致
        """
        results: list[ReadResult] = []
        total = len(objects)

        for i, obj in enumerate(objects):
            result = await self.read_present_value(obj)
            results.append(result)

            done = i + 1
            if progress_cb:
                ret = progress_cb(done, total)
                if asyncio.iscoroutine(ret):
                    await ret
            elif done % config.BATCH_SIZE == 0 or done == total:
                ok = sum(1 for r in results if r.ok)
                log.info("进度 %d/%d  成功=%d  失败=%d",
                         done, total, ok, done - ok)

        return results

    async def read_batch(
        self,
        objects: list[AIObject],
        progress_cb=None,
    ) -> list[ReadResult]:
        """
        分批并发读取 Present Value（每批 BATCH_SIZE 个并发）。
        比 read_all 更快，适合对象数量多时使用。
        """
        results: list[ReadResult] = []
        total = len(objects)

        for batch_start in range(0, total, config.BATCH_SIZE):
            batch = objects[batch_start: batch_start + config.BATCH_SIZE]
            batch_results = await asyncio.gather(
                *[self.read_present_value(obj) for obj in batch]
            )
            results.extend(batch_results)

            done = min(batch_start + config.BATCH_SIZE, total)
            if progress_cb:
                ret = progress_cb(done, total)
                if asyncio.iscoroutine(ret):
                    await ret
            else:
                ok = sum(1 for r in results if r.ok)
                log.info("批量进度 %d/%d  成功=%d  失败=%d",
                         done, total, ok, done - ok)

        return results

    async def read_metadata_batch(
        self,
        objects: list[AIObject],
    ) -> list[tuple[AIObject, str, str, bool]]:
        """
        批量读取每个 AI 对象的 BACnet description 和 units 属性。

        units 读取失败时重试 config.MAX_RETRIES 次（与 presentValue 读取一致），并区分
        "读取失败(超时/异常)"与"网关返回 no-units(空单位)"——前者未拿到值，不能据此
        判定单位与模板不符。读取成功但映射为空（如 no-units）而模板要求有单位时，记录
        原始值到日志，便于定位是网关确无单位还是枚举未覆盖。

        Returns:
            list of (AIObject, bacnet_description, bacnet_unit_str, units_read_failed)
            units_read_failed=True 表示 units 重试用尽仍未读到（未拿到值）。
        """
        results: list[tuple[AIObject, str, str, bool]] = []
        total = len(objects)

        async def _read_attr_retry(request: str) -> tuple[object, bool]:
            """读取单个属性，失败重试 config.MAX_RETRIES 次。返回 (value, read_failed)。"""
            last_exc: Optional[Exception] = None
            for attempt in range(config.MAX_RETRIES + 1):
                try:
                    value = await asyncio.wait_for(
                        self._bacnet.read(request), timeout=config.READ_TIMEOUT
                    )
                    return value, False
                except Exception as exc:
                    last_exc = exc
                    if attempt < config.MAX_RETRIES:
                        await asyncio.sleep(config.RETRY_WAIT)
            log.debug("属性读取失败 [%s]（已重试 %d 次）: %s",
                      request, config.MAX_RETRIES, last_exc)
            return None, True

        for batch_start in range(0, total, config.BATCH_SIZE):
            batch = objects[batch_start: batch_start + config.BATCH_SIZE]

            async def _read_one(obj: AIObject):
                obj_type_str = obj.obj_type
                desc_raw, _ = await _read_attr_retry(
                    f"{self._gw} {obj_type_str} {obj.instance} description"
                )
                units_raw, units_failed = await _read_attr_retry(
                    f"{self._gw} {obj_type_str} {obj.instance} units"
                )
                desc = str(desc_raw).strip() if desc_raw is not None else ""
                unit = _bacnet_unit_to_str(str(units_raw)) if units_raw is not None else ""
                # 诊断：units 读取成功但映射为空（如网关返回 no-units），而模板要求有单位，
                # 记录原始值，便于区分"读取超时""网关确无单位""枚举未覆盖"三种情况。
                if (not units_failed and units_raw is not None
                        and unit == "" and obj.unit):
                    log.warning(
                        "[元数据] units 读到但映射为空 [%s %d] param=%s "
                        "模板单位=%r 网关units原始值=%r",
                        obj_type_str, obj.instance, obj.param_key, obj.unit, units_raw,
                    )
                return obj, desc, unit, units_failed

            batch_results = await asyncio.gather(*[_read_one(o) for o in batch])
            results.extend(batch_results)

            done = min(batch_start + config.BATCH_SIZE, total)
            log.info("元数据读取进度 %d/%d", done, total)

        return results

    # ── 工具 ──────────────────────────────────────────────────────────────────

    async def _read_raw(self, request: str, arr_index: Optional[int] = None):
        """内部读取，返回原始值或 None（不抛异常）。"""
        try:
            if arr_index is not None:
                return await asyncio.wait_for(
                    self._bacnet.read(request, arr_index=arr_index),
                    timeout=config.READ_TIMEOUT,
                )
            return await asyncio.wait_for(
                self._bacnet.read(request),
                timeout=config.READ_TIMEOUT,
            )
        except Exception as exc:
            log.debug("_read_raw 失败 [%s] arr=%s: %s", request, arr_index, exc)
            return None

    @staticmethod
    def summary(results: list[ReadResult]) -> dict:
        """返回读取结果统计。"""
        ok  = [r for r in results if r.ok]
        err = [r for r in results if not r.ok]
        return {
            "total":        len(results),
            "success":      len(ok),
            "failed":       len(err),
            "success_rate": f"{len(ok)/len(results)*100:.1f}%" if results else "N/A",
            "errors":       {r.obj.instance: r.error for r in err},
        }

    # ── 协议规范测试 ──────────────────────────────────────────────────────────

    async def read_device_info(self) -> DeviceInfoResult:
        """读取 Device Object 标准必需属性（ANSI/ASHRAE 135 §12.11）。"""
        result = DeviceInfoResult()
        props = {
            'objectIdentifier':          'object_identifier',
            'objectName':                'object_name',
            'systemStatus':              'system_status',
            'vendorName':                'vendor_name',
            'vendorIdentifier':          'vendor_id',
            'modelName':                 'model_name',
            'firmwareRevision':          'firmware_revision',
            'applicationSoftwareVersion':'app_sw_version',
            'protocolVersion':           'protocol_version',
            'protocolRevision':          'protocol_revision',
            'maxApduLengthAccepted':     'max_apdu_length',
            'segmentationSupported':     'segmentation',
        }
        fetched = 0
        for prop, attr in props.items():
            val = await self._read_raw(f"{self._gw} device {self._dev} {prop}")
            if val is not None:
                setattr(result, attr, str(val))
                fetched += 1
        result.ok = fetched > 0
        if not result.ok:
            result.error = "无法读取任何 Device Object 属性"
        log.info("Device Object 属性：读取 %d/%d 项", fetched, len(props))
        return result

    async def test_error_responses(self, objects: list[AIObject]) -> list[ProtocolErrorTestResult]:
        """BACnet 协议合规性测试（参考 ANSI/ASHRAE 135 §16 错误处理 + AI 必需属性）。"""
        tests: list[ProtocolErrorTestResult] = []

        # 1. 非法对象实例 → 设备应拒绝（返回 None）
        val = await self._read_raw(f"{self._gw} analogInput 9999999 presentValue")
        tests.append(ProtocolErrorTestResult(
            test_name="读取不存在 AI 对象（Instance=9999999）应返回错误",
            passed=(val is None),
            detail="返回 None，符合标准" if val is None else f"未拒绝，意外返回：{val}",
        ))

        if objects:
            obj = objects[0]
            ref = f"{obj.obj_type} {obj.instance}"

            # 2–4. AI/BI 必需属性（135 §12.2.2）
            for prop in ("statusFlags", "outOfService", "units"):
                val = await self._read_raw(f"{self._gw} {ref} {prop}")
                tests.append(ProtocolErrorTestResult(
                    test_name=f"AI 必需属性 {prop}（{obj.param_key}）可读",
                    passed=(val is not None),
                    detail=str(val) if val is not None else "读取失败",
                ))

        log.info("协议合规性测试：%d/%d 通过",
                 sum(1 for t in tests if t.passed), len(tests))
        return tests

    async def check_stability(
        self,
        objects: list[AIObject],
        attempts: int = 5,
        delay: float = 0.5,
    ) -> StabilityResult:
        """连接稳定性测试：对同一 AI 对象连续读取 N 次，检查成功率与一致性。"""
        if not objects:
            return StabilityResult(attempts=attempts, successes=0,
                                   errors=["无可用对象"])
        test_obj = objects[0]
        successes = 0
        errors: list[str] = []
        for i in range(attempts):
            if i > 0:
                await asyncio.sleep(delay)
            r = await self.read_present_value(test_obj)
            if r.ok:
                successes += 1
            else:
                errors.append(f"第{i+1}次：{r.error}")
        log.info("稳定性测试：%d/%d 成功  对象=%s", successes, attempts, test_obj.param_key)
        return StabilityResult(attempts=attempts, successes=successes, errors=errors)
