# -*- coding: utf-8 -*-
"""
enip_reader.py — EtherNet/IP Assembly 读取模块

通过 pycomm3 对 AcuRev-4100 发送 CIP 显式消息，
读取 Assembly Instance 100 的原始字节，按 EDS 数据类型解析为工程值。

EDS 关键信息：
  Assembly Object Class = 0x04
  Assembly Instance     = 100 (0x64)
  Attribute             = 3 (data)
  总字节数              = 5648
  参数布局              = EDS [Params] 顺序，各自不同数据类型
"""
from __future__ import annotations

import re
import struct
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)

# ─── CIP 数据类型编码 → (struct格式, 字节数) ────────────────────────────────
# CIP 协议规定所有数值类型均为小端（Intel byte order）
_DTYPE_FMT: dict[str, tuple[str, int]] = {
    "0xca": ("<f",  4),   # REAL    float32
    "0xcb": ("<d",  8),   # LREAL   float64
    "0xc8": ("<I",  4),   # UDINT   uint32
    "0xc7": ("<H",  2),   # UINT    uint16
    "0xc6": ("<B",  1),   # USINT   uint8
    "0xc4": ("<i",  4),   # DINT    int32
    "0xc3": ("<h",  2),   # INT     int16
    "0xc2": ("<b",  1),   # SINT    int8
}


@dataclass
class EnipParam:
    """单个参数的 EDS 元数据。"""
    index:     int        # Param 编号（1-based）
    name:      str        # paramType，如 FREQ_Hz
    unit:      str        # 工程单位
    dtype_hex: str        # 0xca / 0xcb / ...（小写）
    fmt:       str        # struct 格式字符
    size:      int        # 字节数
    offset:    int        # 在 Assembly 中的字节偏移


@dataclass
class EnipResult:
    """单个参数的读取结果。"""
    param_key: str
    value:     Optional[float] = None
    error:     str = ""

    @property
    def ok(self) -> bool:
        return self.error == ""


@dataclass
class IdentityResult:
    """CIP Identity Object（Class=0x01, Instance=0x01）属性读取结果。"""
    vendor_id:      Optional[int] = None
    device_type:    Optional[int] = None
    product_code:   Optional[int] = None
    revision_major: Optional[int] = None
    revision_minor: Optional[int] = None
    status:         Optional[int] = None
    serial_number:  Optional[int] = None
    product_name:   Optional[str] = None
    error:          str = ""

    @property
    def ok(self) -> bool:
        return self.error == ""


@dataclass
class ErrorTestResult:
    """单条 CIP 错误响应测试结果。"""
    name:        str
    description: str
    passed:      bool = False
    detail:      str = ""


@dataclass
class AssemblyIntegrityResult:
    """Assembly 结构合规性检查结果。"""
    eds_total_bytes:  int             # EDS 声明的 Assembly 总字节数
    actual_bytes:     int             # 实际读取到的字节数（0 = 未读取）
    param_count:      int
    out_of_bounds:    list[str]       # 超出 Assembly 范围的参数名

    @property
    def bytes_match(self) -> bool:
        return self.actual_bytes > 0 and self.actual_bytes == self.eds_total_bytes

    @property
    def ok(self) -> bool:
        return self.bytes_match and not self.out_of_bounds


@dataclass
class StabilityResult:
    """连接稳定性测试结果：连续多次读取 Assembly。"""
    attempts:  int
    successes: int
    errors:    list[str]

    @property
    def ok(self) -> bool:
        return self.successes == self.attempts


# ─── EDS 文件查找 ────────────────────────────────────────────────────────────

def find_eds_file(eds_dir: str, device_name: str) -> str:
    """
    在 eds_dir 下递归查找设备 EDS 文件（大小写/连字符/下划线不敏感）。
    例：device_name='AcuRev4100' → 匹配 'AcuRev-4100.eds'
    """
    needle = device_name.lower().replace("-", "").replace("_", "")
    for p in Path(eds_dir).rglob("*.eds"):
        if needle in p.stem.lower().replace("-", "").replace("_", ""):
            return str(p)
    raise FileNotFoundError(f"在 {eds_dir} 中未找到设备 '{device_name}' 的 EDS 文件")


# ─── EDS 解析 ────────────────────────────────────────────────────────────────

def parse_eds(eds_path: str) -> list[EnipParam]:
    """
    解析 EDS 文件，返回按 Assembly 顺序排列的参数列表（含字节偏移）。
    """
    text = Path(eds_path).read_text(encoding="utf-8", errors="replace")

    # ── 1. 解析 [Params] 段 ──────────────────────────────────────────────────
    params_raw: dict[int, dict] = {}
    params_section = text[text.find("[Params]"): text.find("[Assembly]")]
    for num_str, body in re.findall(r"Param(\d+)\s*=(.*?);", params_section, re.DOTALL):
        quoted = re.findall(r'"([^"]*)"', body)
        # 找 Data Type 行（格式固定，第2个 0x 值是 DataType）
        dtype_match = re.search(r"(0x[0-9A-Fa-f]+),\s*\$\s*Data Type", body, re.IGNORECASE)
        if not dtype_match:
            hex_vals = re.findall(r"0x[0-9A-Fa-f]+", body)
            dtype_hex = hex_vals[1].lower() if len(hex_vals) >= 2 else "0xca"
        else:
            dtype_hex = dtype_match.group(1).lower()

        params_raw[int(num_str)] = {
            "name":      quoted[0] if len(quoted) > 0 else f"Param{num_str}",
            "unit":      quoted[1] if len(quoted) > 1 else "",
            "dtype_hex": dtype_hex,
        }

    # ── 2. 解析 Assembly 100 的参数顺序 ─────────────────────────────────────
    asm_section = text[text.find("[Assembly]"): text.find("[Connection Manager]")]
    assem_block = re.search(r"Assem100\s*=(.*?);", asm_section, re.DOTALL)
    if not assem_block:
        raise ValueError("EDS 中未找到 Assem100 定义")

    # 提取 bits,ParamN 条目（排除第一条 size 声明行）
    raw_entries = re.findall(r"(\d+),\s*(Param\d+)", assem_block.group(1))

    # ── 3. 构建有偏移的参数列表 ──────────────────────────────────────────────
    result: list[EnipParam] = []
    offset = 0
    for bits_str, param_ref in raw_entries:
        num = int(param_ref[5:])          # "Param42" → 42
        info = params_raw.get(num, {})
        dtype_hex = info.get("dtype_hex", "0xca")
        fmt, size = _DTYPE_FMT.get(dtype_hex, (">f", 4))

        result.append(EnipParam(
            index     = num,
            name      = info.get("name", param_ref),
            unit      = info.get("unit", ""),
            dtype_hex = dtype_hex,
            fmt       = fmt,
            size      = size,
            offset    = offset,
        ))
        offset += size

    log.info("EDS 解析完成：%d 个参数，Assembly 总字节 %d", len(result), offset)
    return result


# ─── Assembly 字节解析 ────────────────────────────────────────────────────────

def parse_assembly_bytes(raw: bytes, params: list[EnipParam]) -> dict[str, float]:
    """将 Assembly 原始字节按参数布局解析为 {param_key: value} 字典。"""
    values: dict[str, float] = {}
    for p in params:
        end = p.offset + p.size
        if end > len(raw):
            log.warning("参数 %s 超出 Assembly 范围（offset=%d size=%d total=%d）",
                        p.name, p.offset, p.size, len(raw))
            break
        (v,) = struct.unpack_from(p.fmt, raw, p.offset)
        values[p.name] = float(v)
    return values


# ─── Assembly 结构合规性 & 稳定性检查 ─────────────────────────────────────────────

def check_assembly_integrity(params: list[EnipParam], actual_bytes: int = 0
                              ) -> AssemblyIntegrityResult:
    """
    根据 EDS 解析结果验证 Assembly 布局：
      1. EDS 声明总字节数 == 实际读取字节数
      2. 所有参数的 offset+size 均在 Assembly 范围内
    """
    eds_total = sum(p.size for p in params)
    oob: list[str] = []
    bound = actual_bytes if actual_bytes > 0 else eds_total
    for p in params:
        if p.offset + p.size > bound:
            oob.append(f"{p.name}(offset={p.offset}, size={p.size})")
    return AssemblyIntegrityResult(
        eds_total_bytes = eds_total,
        actual_bytes    = actual_bytes,
        param_count     = len(params),
        out_of_bounds   = oob,
    )


def check_connection_stability(driver, attempts: int = 3, delay: float = 1.0
                                ) -> StabilityResult:
    """连续读取 Assembly Instance 100，统计成功/失败次数。"""
    import time
    successes = 0
    errors: list[str] = []
    for i in range(attempts):
        if i > 0:
            time.sleep(delay)
        try:
            resp = driver.generic_message(
                service=b"\x0E",
                class_code=0x04,
                instance=100,
                attribute=3,
                connected=False,
                route_path=True,
            )
            if resp and resp.value is not None:
                successes += 1
            else:
                err = str(getattr(resp, "error", "")) or "无响应"
                errors.append(f"第{i+1}次: {err}")
        except Exception as exc:
            errors.append(f"第{i+1}次: {exc}")
    return StabilityResult(attempts=attempts, successes=successes, errors=errors)


# ─── CIP 对象读取工具 ────────────────────────────────────────────────────────────

def _cip_get_attr(driver, class_code: int, instance: int, attribute: int
                  ) -> tuple[Optional[bytes], str]:
    """
    发送 Get_Attribute_Single (0x0E)，返回 (data_bytes, "") 或 (None, error_str)。
    """
    try:
        resp = driver.generic_message(
            service=b"\x0E",
            class_code=class_code,
            instance=instance,
            attribute=attribute,
            connected=False,
            route_path=True,
        )
        if resp and resp.value is not None:
            raw = resp.value if isinstance(resp.value, (bytes, bytearray)) else bytes(resp.value)
            return raw, ""
        err = str(getattr(resp, "error", "")) or "无响应"
        return None, err
    except Exception as exc:
        return None, str(exc)


def read_identity(driver) -> IdentityResult:
    """读取 CIP Identity Object（Class=0x01, Instance=0x01）全部标准属性。"""
    result = IdentityResult()
    errors: list[str] = []

    def get(attr: int) -> Optional[bytes]:
        data, err = _cip_get_attr(driver, 0x01, 0x01, attr)
        if err:
            errors.append(f"Attr{attr}: {err}")
        return data

    if (d := get(1)) and len(d) >= 2:
        result.vendor_id = struct.unpack_from("<H", d, 0)[0]
    if (d := get(2)) and len(d) >= 2:
        result.device_type = struct.unpack_from("<H", d, 0)[0]
    if (d := get(3)) and len(d) >= 2:
        result.product_code = struct.unpack_from("<H", d, 0)[0]
    if (d := get(4)) and len(d) >= 2:
        result.revision_major, result.revision_minor = d[0], d[1]
    if (d := get(5)) and len(d) >= 2:
        result.status = struct.unpack_from("<H", d, 0)[0]
    if (d := get(6)) and len(d) >= 4:
        result.serial_number = struct.unpack_from("<I", d, 0)[0]
    if (d := get(7)) and len(d) >= 1:
        n = d[0]
        result.product_name = d[1 : 1 + n].decode("ascii", errors="replace")

    if errors:
        result.error = "; ".join(errors)
    return result


def test_error_responses(driver) -> list[ErrorTestResult]:
    """
    向设备发送非法 CIP 请求，验证设备正确返回错误响应而非意外成功。
    四条用例：不存在的实例、不存在的属性、不存在的类、不存在的 Assembly 实例。
    """
    results: list[ErrorTestResult] = []

    def expect_fail(name: str, desc: str, cls: int, inst: int, attr: int) -> None:
        data, err = _cip_get_attr(driver, cls, inst, attr)
        if data is None:
            results.append(ErrorTestResult(name=name, description=desc, passed=True,
                                           detail=err[:150] if err else ""))
        else:
            results.append(ErrorTestResult(name=name, description=desc, passed=False,
                                           detail=f"意外成功，返回 {len(data)} 字节"))

    expect_fail("T-ERR-01",
                "不存在的 Identity 实例（Class=0x01, Instance=0xFF, Attr=0x01）",
                0x01, 0xFF, 0x01)
    expect_fail("T-ERR-02",
                "不存在的属性（Class=0x01, Instance=0x01, Attr=0x7F）",
                0x01, 0x01, 0x7F)
    expect_fail("T-ERR-03",
                "不存在的 CIP 类（Class=0x7F, Instance=0x01, Attr=0x01）",
                0x7F, 0x01, 0x01)
    expect_fail("T-ERR-04",
                "不存在的 Assembly 实例（Class=0x04, Instance=0xFF, Attr=0x03）",
                0x04, 0xFF, 0x03)
    return results


# ─── EtherNet/IP 读取器 ───────────────────────────────────────────────────────

class EnipReader:
    """
    使用 pycomm3 读取 AcuRev-4100 EtherNet/IP Assembly Instance 100。

    用法：
        async with EnipReader() as reader:
            results = await reader.read_all()
    """

    ASSEMBLY_INSTANCE = 100
    ASSEMBLY_CLASS    = 0x04
    ASSEMBLY_ATTR     = 3

    def __init__(self, host: str, eds_path: str, slot: int = 0):
        self.host          = host
        self.eds_path      = eds_path
        self.slot          = slot
        self._params:      list[EnipParam] = []
        self._driver       = None
        self._actual_bytes: int = 0   # 最近一次 Assembly 读取的实际字节数

    async def __aenter__(self) -> "EnipReader":
        from pycomm3 import CIPDriver
        self._params = parse_eds(self.eds_path)
        self._driver = CIPDriver(self.host)
        self._driver.open()
        log.info("EtherNet/IP 已连接：%s  参数数：%d", self.host, len(self._params))
        return self

    async def __aexit__(self, *_):
        if self._driver:
            try:
                self._driver.close()
            except Exception:
                pass

    @property
    def param_names(self) -> list[str]:
        return [p.name for p in self._params]

    def read_all_sync(self) -> dict[str, EnipResult]:
        """同步读取全部参数，返回 {param_key: EnipResult}。"""
        results: dict[str, EnipResult] = {}
        try:
            tag = f"@{self.ASSEMBLY_CLASS}/{self.ASSEMBLY_INSTANCE}/{self.ASSEMBLY_ATTR}"
            resp = self._driver.generic_message(
                service=b"\x0E",              # Get_Attribute_Single
                class_code=self.ASSEMBLY_CLASS,
                instance=self.ASSEMBLY_INSTANCE,
                attribute=self.ASSEMBLY_ATTR,
                connected=False,
                route_path=True,
            )
            if not resp:
                raise RuntimeError(f"读取 Assembly 失败：{resp.error}")

            raw: bytes = resp.value if isinstance(resp.value, (bytes, bytearray)) else bytes(resp.value)
            self._actual_bytes = len(raw)
            log.info("Assembly 读取成功，字节数: %d", len(raw))
            values = parse_assembly_bytes(raw, self._params)
            for name, val in values.items():
                results[name] = EnipResult(param_key=name, value=val)

        except Exception as exc:
            log.error("EtherNet/IP 读取异常: %s", exc)
            for p in self._params:
                results[p.name] = EnipResult(param_key=p.name, error=str(exc))

        return results

    async def read_all(self) -> dict[str, EnipResult]:
        """异步包装（在线程中执行同步读取，保持与现有异步框架兼容）。"""
        import asyncio
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(None, self.read_all_sync)

    def read_identity_sync(self) -> IdentityResult:
        return read_identity(self._driver)

    def test_error_responses_sync(self) -> list[ErrorTestResult]:
        return test_error_responses(self._driver)

    def check_stability_sync(self, attempts: int = 3) -> StabilityResult:
        return check_connection_stability(self._driver, attempts=attempts)
