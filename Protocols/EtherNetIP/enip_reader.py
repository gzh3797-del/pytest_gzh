# -*- coding: utf-8 -*-
"""
enip_reader.py — EtherNet/IP Assembly 读取模块

通过 pycomm3 对 AcuRev-4100 发送 CIP 显式消息，
读取 Assembly Instance 10 的原始字节，按 EDS 数据类型解析为工程值。

EDS 关键信息：
  Assembly Object Class = 0x04
  Assembly Instance     = 10 (0x0A)
  Attribute             = 3 (data)
  总字节数              = 5656
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
    index:             int        # Param 编号（1-based）
    name:              str        # paramType，如 FREQ_Hz
    unit:              str        # 工程单位
    dtype_hex:         str        # 0xca / 0xcb / ...（小写）
    fmt:               str        # struct 格式字符
    size:              int        # 字节数
    offset:            int        # 在所属 Assembly 中的字节偏移
    assembly_instance: int = 10   # 所属 CIP Assembly 实例号（默认 10）


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
    eds_total_bytes:    int         # 所有输入 Assembly 参数字节数之和
    actual_bytes:       int         # 实际读取字节数之和（0 = 未读取）
    param_count:        int
    out_of_bounds:      list[str]   # 超出各 Assembly 范围的参数名
    alignment_issues:   list[str]   # LREAL 等类型未满足对齐要求的参数
    per_assembly_sizes: dict        # {assembly_instance: computed_bytes}（供三路一致性检查）

    @property
    def bytes_match(self) -> bool:
        return self.actual_bytes > 0 and self.actual_bytes == self.eds_total_bytes

    @property
    def ok(self) -> bool:
        # 仅反映运行时读取完整性：字节数匹配 + 无越界
        # alignment_issues 属于静态 EDS 设计合规，由 Section 8 单独展示
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


@dataclass
class ConnectionManagerResult:
    """EDS [Connection Manager] Connection1 静态验证结果。"""
    field_count:      int         # 实际解析到的字段数
    has_ot_direction: bool        # 是否存在 O->T 方向定义（Assem 引用数 >= 2）
    connection_name:  str         # Connection Name 字段内容（应非空）
    path:             str         # Path 字段内容（应非空）
    eds_revision:     str         # EDS [File] Revision（用于与 Identity 比对）
    issues:           list[str]   # 所有发现的问题描述
    raw_fields:       list[str]   # 按 ODVA 15 字段位置的原始 token（不足处补空串）
    to_size_declared: int         # Connection1 T→O Size 字段的声明字节数（-1=未声明/无效）

    # 期望字段数：双向（O->T + T->O）= 15；仅 T->O（Listen-Only）= 12
    EXPECTED_BIDIR = 15
    EXPECTED_TONLY = 12

    # ODVA Connection1 各位置的字段名（0-indexed，共 15 个）
    FIELD_NAMES = [
        "Trigger & Transport",    # 0
        "Connection Parameters",  # 1
        "O→T RPI",                # 2
        "O→T Size",               # 3
        "O→T Format",             # 4
        "T→O RPI",                # 5
        "T→O Size",               # 6
        "T→O Format",             # 7  ← EZ-EDS Error #1
        "Proxy Config Size",      # 8
        "Proxy Config Format",    # 9  ← EZ-EDS Error #2
        "Target Config Size",     # 10 ← EZ-EDS Error #3
        "Target Config Format",   # 11 ← EZ-EDS Error #4
        "Connection Name",        # 12 ← EZ-EDS Error #5
        "Help String",            # 13
        "Path",                   # 14 ← EZ-EDS Error #6
    ]

    # 需要规则检查的关键位置（0-indexed），用于 HTML 高亮
    CHECKED_POSITIONS = {7, 9, 10, 11, 12, 14}

    @property
    def ok(self) -> bool:
        return len(self.issues) == 0


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

def parse_eds(eds_path: str,
              assem_ids: Optional[list[int]] = None) -> list[EnipParam]:
    """
    解析 EDS 文件，返回按 Assembly 顺序排列的参数列表（含字节偏移）。

    Args:
        eds_path:  EDS 文件路径。
        assem_ids: 指定只解析的 Assembly 实例 ID 列表；None 表示解析所有输入 Assembly。
                   多表模式下由 parse_device_assembly_map() 的结果传入，每次只解析
                   当前设备对应的 Assembly，避免不同设备的参数混入同一次比对。
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

    # ── 2. 解析所有输入类 Assembly 实例 ──────────────────────────────────────
    asm_section = text[text.find("[Assembly]"): text.find("[Connection Manager]")]
    result: list[EnipParam] = []

    for m_assem in re.finditer(r"Assem(\d+)\s*=(.*?);", asm_section, re.DOTALL):
        assem_id   = int(m_assem.group(1))
        assem_body = m_assem.group(2)

        # 若指定了 assem_ids 白名单，跳过不属于当前设备的 Assembly
        if assem_ids is not None and assem_id not in assem_ids:
            continue

        # 跳过输出/配置类 Assembly（名称含 Set/Config/Write/Output 关键字）
        name_match = re.search(r'"([^"]*)"', assem_body)
        assem_name = name_match.group(1) if name_match else ""
        if any(kw in assem_name.lower() for kw in ("set", "config", "write", "output")):
            continue

        raw_entries = re.findall(r"(\d+),\s*(Param\d+)", assem_body)
        if not raw_entries:
            continue

        # ── 3. 按 Assembly 逐参数计算偏移并构建列表 ──────────────────────
        # 注意：同一 Param 可在多个 Assembly 中出现（如 Param1062 在 Assem17 和 Assem23 均被引用），
        # 不跨 Assembly 去重，保留各 Assembly 的完整字节布局，字节计数和 read_all_sync 均正确。
        # eds_map / eds_set 在上层由 {p.name: p} 字典构造，自动对 param_key 去重。
        offset = 0
        for _bits_str, param_ref in raw_entries:
            num  = int(param_ref[5:])          # "Param42" → 42
            info = params_raw.get(num, {})
            name = info.get("name", param_ref)
            dtype_hex = info.get("dtype_hex", "0xca")
            fmt, size = _DTYPE_FMT.get(dtype_hex, (">f", 4))
            result.append(EnipParam(
                index             = num,
                name              = name,
                unit              = info.get("unit", ""),
                dtype_hex         = dtype_hex,
                fmt               = fmt,
                size              = size,
                offset            = offset,
                assembly_instance = assem_id,
            ))
            offset += size

    if not result:
        raise ValueError("EDS 中未解析到任何输入 Assembly 参数，请检查 EDS 文件或 assem_ids 过滤条件")

    assem_count = len({p.assembly_instance for p in result})
    log.info("EDS 解析完成：%d 个参数，跨 %d 个输入 Assembly 实例，总字节 %d",
             len(result), assem_count, sum(p.size for p in result))
    return result


def parse_eds_file_revision(text: str) -> str:
    """
    从 EDS 文件文本中解析 [File] 段的 Revision 字段。
    返回字符串如 "1.00"，未找到则返回 ""。
    """
    m = re.search(r'\[File\].*?Revision\s*=\s*([\d.]+)\s*;', text,
                  re.DOTALL | re.IGNORECASE)
    return m.group(1).strip() if m else ""


def parse_device_assembly_map(eds_text: str) -> dict[str, list[int]]:
    """
    解析 EDS [Connection Manager] 所有 ConnectionN 条目，
    返回 {help_string: [to_assem_id, ...]} 映射。

    EDS 中同一物理设备的所有 Connection 共享相同的 help string；
    每条 Connection 的最后一个 AssemN 引用为 T→O Assembly（数据输入方向）。
    结果按 Assembly ID 升序排列，与 parse_eds() 的 assem_ids 参数直接对应。
    """
    cm_start = eds_text.find("[Connection Manager]")
    if cm_start == -1:
        return {}

    next_sec = re.search(r'^\[(?!Connection Manager)', eds_text[cm_start + 1:], re.MULTILINE)
    cm_end   = (cm_start + 1 + next_sec.start()) if next_sec else len(eds_text)
    cm_text  = eds_text[cm_start:cm_end]

    device_map: dict[str, list[int]] = {}

    for m_conn in re.finditer(r'Connection\d+\s*=(.*?);', cm_text, re.DOTALL):
        body = m_conn.group(1)

        # 最后一个 AssemN 引用 = T→O Assembly（输入方向）
        assem_refs = re.findall(r'\bAssem(\d+)\b', body)
        if not assem_refs:
            continue
        to_assem_id = int(assem_refs[-1])

        # help string = Connection 体内第二个引号字符串（顺序：Name / help / Path）
        quoted = re.findall(r'"([^"]*)"', body)
        if len(quoted) < 2:
            continue
        help_str = quoted[1].strip()
        if not help_str:
            continue

        if help_str not in device_map:
            device_map[help_str] = []
        if to_assem_id not in device_map[help_str]:
            device_map[help_str].append(to_assem_id)

    # 各设备的 Assembly 列表按 ID 升序
    for key in device_map:
        device_map[key].sort()

    return device_map


def parse_connection_manager(text: str) -> ConnectionManagerResult:
    """
    静态解析 EDS [Connection Manager] 段的 Connection1 条目，
    按 ODVA 规范 15 字段位置逐一验证，对应 EZ-EDS 的全部 6 条检查项。

    ODVA Connection1 字段布局（1-indexed / 0-indexed）：
      [1/0]  Trigger & Transport
      [2/1]  Connection Parameters
      [3/2]  O→T RPI          [4/3]  O→T Size       [5/4]  O→T Format
      [6/5]  T→O RPI          [7/6]  T→O Size        [8/7]  T→O Format   ← Error #1
      [9/8]  Proxy Cfg Size  [10/9]  Proxy Cfg Fmt                        ← Error #2
      [11/10] Target Cfg Size (0-65535)                                   ← Error #3
      [12/11] Target Cfg Format                                           ← Error #4
      [13/12] Connection Name (quoted, required)                          ← Error #5
      [14/13] Help String     [15/14] Path (quoted, required)             ← Error #6
    """
    issues: list[str] = []

    # ── 解析 [File] Revision ─────────────────────────────────────────────────
    eds_revision = parse_eds_file_revision(text)

    # ── 找 [Connection Manager] 段 ──────────────────────────────────────────
    cm_start = text.find("[Connection Manager]")
    if cm_start == -1:
        issues.append("[Connection Manager] 段缺失")
        return ConnectionManagerResult(
            field_count=0, has_ot_direction=False,
            connection_name="", path="",
            eds_revision=eds_revision, issues=issues, raw_fields=[],
            to_size_declared=-1,
        )

    next_section = re.search(r'^\[(?!Connection Manager)', text[cm_start + 1:], re.MULTILINE)
    cm_end = (cm_start + 1 + next_section.start()) if next_section else len(text)
    cm_text = text[cm_start:cm_end]

    # ── 提取 Connection1 = ...;  ─────────────────────────────────────────────
    c1_match = re.search(r'Connection1\s*=(.*?);', cm_text, re.DOTALL)
    if not c1_match:
        issues.append("Connection1 条目缺失")
        return ConnectionManagerResult(
            field_count=0, has_ot_direction=False,
            connection_name="", path="",
            eds_revision=eds_revision, issues=issues, raw_fields=[],
            to_size_declared=-1,
        )

    c1_body  = c1_match.group(1)
    c1_clean = re.sub(r'\$[^\n]*', '', c1_body)   # 去掉行内注释

    # ── 按逗号分割，保留中间空字段（,, = 合法空值），去掉末尾空串 ──────────
    tokens = [t.strip() for t in c1_clean.split(',')]
    while tokens and tokens[-1] == '':
        tokens.pop()
    field_count = len(tokens)

    # ── 位置辅助函数 ─────────────────────────────────────────────────────────
    def _tok(idx: int) -> str:
        """按 0-based 索引取 token，越界返回空串。"""
        return tokens[idx] if idx < len(tokens) else ""

    def _strip_q(s: str) -> str:
        """去掉首尾引号：'"foo"' → 'foo'，'"' → ''。"""
        s = s.strip()
        if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
            return s[1:-1]
        return s

    # ── ODVA 15 字段位置提取（0-indexed）────────────────────────────────────
    to_format        = _tok(7)    # T→O Format (pos 8)
    proxy_cfg_fmt    = _tok(9)    # Proxy Config Format (pos 10)
    target_cfg_size  = _tok(10)   # Target Config Size (pos 11)
    target_cfg_fmt   = _tok(11)   # Target Config Format (pos 12)
    connection_name  = _strip_q(_tok(12))   # Connection Name (pos 13) — 位置解析
    path             = _strip_q(_tok(14))   # Path (pos 15)            — 位置解析

    # ── 统计 Assembly 引用数（用于双向方向检查）──────────────────────────────
    assem_refs       = re.findall(r'\bAssem\d+\b', c1_clean, re.IGNORECASE)
    has_ot_direction = len(assem_refs) >= 2

    # ── 问题判断（与 EZ-EDS 6 条错误逐一对应）──────────────────────────────

    # EZ-EDS Error #1: T→O Format 字段为空
    # 说明：定义了 T→O 连接但 Format（Assembly引用）缺失；此处以 pos 8 为空触发
    if not to_format or (to_format.startswith('"') and not _strip_q(to_format)):
        issues.append(
            "T→O Format 字段为空或缺失——定义 T→O 连接时必须引用 Assembly"
            "（EZ-EDS: 'T->O Format: This conditional field is empty'）"
        )

    # EZ-EDS Error #2: Proxy Config Format 内容不合规
    # 说明：非模块化设备此位置应为空或合法引用；出现非空带引号字符串说明字段错位
    if proxy_cfg_fmt.startswith('"') and _strip_q(proxy_cfg_fmt):
        issues.append(
            f"Proxy Config Format 位置存在错位字段 {proxy_cfg_fmt!r}——"
            f"非模块化设备此处应为空"
            f"（EZ-EDS: 'Proxy Config Format: For a non-modular device...'）"
        )

    # EZ-EDS Error #3: Target Config Size 数据类型错误
    # 说明：此字段应为整数（0-65535），出现带引号字符串说明字段整体错位
    if target_cfg_size.startswith('"'):
        issues.append(
            f"Target Config Size 字段类型错误：{target_cfg_size!r}——"
            f"应为整数（0-65535）而非带引号字符串"
            f"（EZ-EDS: 'Target Config Size: Field value has wrong data type'）"
        )

    # EZ-EDS Error #4: Target Config Format 引用不合规
    # 说明：此位置出现带引号路径字符串（路径应在 Path 字段 pos 15，不在此处）
    if target_cfg_fmt.startswith('"') and _strip_q(target_cfg_fmt):
        issues.append(
            f"Target Config Format 位置存在错位路径字符串 {target_cfg_fmt!r}——"
            f"路径字符串应在 Path 字段（pos 15），不应出现在此"
            f"（EZ-EDS: 'Target Config Format: Wrong reference'）"
        )

    # 双向方向检查（Error #1 根因：Assembly 引用数不足，缺少双向定义）
    if not has_ot_direction:
        issues.append(
            f"Connection1 仅含 {len(assem_refs)} 个 Assembly 引用，"
            f"缺少双向 O→T + T→O 定义（期望 ≥ 2 个 Assembly 引用）"
        )

    # 字段数检查
    if field_count < ConnectionManagerResult.EXPECTED_TONLY:
        issues.append(
            f"Connection1 字段数不足：实际 {field_count}，"
            f"期望最少 {ConnectionManagerResult.EXPECTED_TONLY}（仅 T→O）"
            f" 或 {ConnectionManagerResult.EXPECTED_BIDIR}（双向）"
        )
    elif field_count < ConnectionManagerResult.EXPECTED_BIDIR and has_ot_direction:
        issues.append(
            f"Connection1 声明了双向连接但字段数不足：实际 {field_count}，"
            f"期望 {ConnectionManagerResult.EXPECTED_BIDIR}"
        )

    # EZ-EDS Error #5: Connection Name 必填项为空（按位置解析）
    if not connection_name:
        issues.append(
            "Connection Name 字段为空（必填项）"
            "（EZ-EDS: 'Connection Name String: This required field is empty'）"
        )

    # EZ-EDS Error #6: Path 必填项为空（按位置解析，修复了之前的误判）
    if not path:
        issues.append(
            "Path 字段为空（必填项）"
            "（EZ-EDS: 'Path: This required field is empty'）"
        )

    # 按 ODVA 15 字段位置保存原始 token（不足处补空串，供 HTML 原始字段表展示）
    padded = (tokens + [''] * 15)[:15]

    # T→O Size（ODVA pos 7 = idx 6）：若为有效整数则记录，供与 Assembly 计算大小交叉验证
    to_size_str = _tok(6)
    to_size_declared = int(to_size_str) if re.match(r'^\d+$', to_size_str) else -1

    log.info(
        "Connection Manager 解析：字段数=%d  O->T=%s  T->O_Size=%d  Name=%r  Path=%r  问题=%d",
        field_count, has_ot_direction, to_size_declared, connection_name, path, len(issues),
    )
    return ConnectionManagerResult(
        field_count      = field_count,
        has_ot_direction = has_ot_direction,
        connection_name  = connection_name,
        path             = path,
        eds_revision     = eds_revision,
        issues           = issues,
        raw_fields       = padded,
        to_size_declared = to_size_declared,
    )


def parse_eds_device_section(text: str) -> dict[str, str]:
    """
    解析 EDS [Device] 段，返回字段字典。
    常用字段：VendCode / VendName / ProdType / ProdTypeStr / ProdCode /
              MajRev / MinRev / ProdName / Catalog
    供与 CIP Identity Object 进行交叉比对。
    """
    result: dict[str, str] = {}
    dev_start = text.find("[Device]")
    if dev_start == -1:
        return result
    next_section = re.search(r'^\[(?!Device)', text[dev_start + 1:], re.MULTILINE)
    dev_end = (dev_start + 1 + next_section.start()) if next_section else len(text)
    dev_text = text[dev_start:dev_end]
    for m in re.finditer(r'(\w+)\s*=\s*("(?:[^"\\]|\\.)*?"|[\w.]+)\s*;', dev_text):
        result[m.group(1).strip()] = m.group(2).strip().strip('"')
    return result


def parse_eds_assembly_declared_size(text: str, assem_id: int = 10) -> int:
    """
    从 EDS [Assembly] 段的 AssemXX 头部提取声明的字节数。

    Assem10 格式：
      Assem10 = "Description", HelpStr, SizeBytes, Flags, [proxy/target,,], [32,ParamN, ...]

    返回 SizeBytes（整数），未找到或格式不合规时返回 -1。
    """
    asm_start = text.find("[Assembly]")
    if asm_start == -1:
        return -1
    next_section = re.search(r'^\[(?!Assembly)', text[asm_start + 1:], re.MULTILINE)
    asm_end = (asm_start + 1 + next_section.start()) if next_section else len(text)
    asm_text = text[asm_start:asm_end]

    pattern = rf'Assem{assem_id}\s*=(.*?);'
    m = re.search(pattern, asm_text, re.DOTALL)
    if not m:
        return -1

    body = re.sub(r'\$[^\n]*', '', m.group(1))   # 去注释
    tokens = [t.strip() for t in body.split(',')]

    # 跳过前两个（Description, HelpStr），第三个是字节数
    non_comment = [t for t in tokens if t]
    # 跳过开头的带引号字段（Description / HelpStr），取第一个纯整数
    idx = 0
    for t in non_comment:
        if t.startswith('"'):
            idx += 1
            continue
        if re.match(r'^\d+$', t):
            return int(t)
        idx += 1
    return -1


def check_eds_orphan_params(text: str, assem_id: int = 10) -> list[str]:
    """
    检查 [Assembly] Assem10 中引用的 ParamN 是否都在 [Params] 中已定义。
    返回孤儿引用列表（在 Assembly 中出现但 [Params] 中缺失）。
    """
    # ── 读取 [Params] 已定义的 Param 编号 ──────────────────────────────────
    params_start = text.find("[Params]")
    params_end   = text.find("[Assembly]") if text.find("[Assembly]") > 0 else len(text)
    if params_start == -1:
        return []
    defined_nums: set[int] = set()
    for m in re.finditer(r'Param(\d+)\s*=', text[params_start:params_end]):
        defined_nums.add(int(m.group(1)))

    # ── 读取 [Assembly] Assem10 中引用的 ParamN ───────────────────────────
    asm_start = text.find("[Assembly]")
    if asm_start == -1:
        return []
    next_section = re.search(r'^\[(?!Assembly)', text[asm_start + 1:], re.MULTILINE)
    asm_end = (asm_start + 1 + next_section.start()) if next_section else len(text)
    asm_text = text[asm_start:asm_end]

    pattern = rf'Assem{assem_id}\s*=(.*?);'
    m = re.search(pattern, asm_text, re.DOTALL)
    if not m:
        return []

    referenced_nums: set[int] = set()
    for ref in re.finditer(r'\bParam(\d+)\b', m.group(1)):
        referenced_nums.add(int(ref.group(1)))

    orphans = sorted(referenced_nums - defined_nums)
    return [f"Param{n}" for n in orphans]


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
    根据 EDS 解析结果验证 Assembly 布局（支持多 Assembly 实例）：
      1. 各 Assembly 组的 EDS 计算字节数之和 vs 实际读取字节数之和
      2. 每个 Assembly 组内，参数的 offset+size 不超出本组字节边界
      3. LREAL(0xcb, 8字节) / LINT(0xc5, 8字节) 参数偏移必须满足 8 字节对齐；
         DINT/UDINT(4字节) 参数偏移须满足 4 字节对齐
    """
    # 各数据类型要求的对齐字节数（0 = 不检查）
    _ALIGN: dict[str, int] = {
        "0xcb": 8,   # LREAL float64
        "0xc5": 8,   # LINT  int64
        "0xc8": 4,   # UDINT uint32
        "0xc4": 4,   # DINT  int32
    }

    # 按 Assembly 实例分组，计算各组字节数
    by_instance: dict[int, list[EnipParam]] = {}
    for p in params:
        by_instance.setdefault(p.assembly_instance, []).append(p)

    per_assembly_sizes: dict[int, int] = {
        aid: sum(p.size for p in grp) for aid, grp in by_instance.items()
    }

    eds_total = sum(per_assembly_sizes.values())
    # 参数总数以唯一 param_key 计；同一参数可被多个 Assembly 引用，不重复计入
    unique_param_count = len({p.name for p in params})
    oob:   list[str] = []
    align: list[str] = []

    for aid, grp in by_instance.items():
        grp_size = per_assembly_sizes[aid]
        for p in grp:
            if p.offset + p.size > grp_size:
                oob.append(f"{p.name}(assem={aid}, offset={p.offset}, size={p.size})")
            req = _ALIGN.get(p.dtype_hex, 0)
            if req and p.offset % req != 0:
                align.append(
                    f"{p.name}(assem={aid}, dtype={p.dtype_hex}, offset={p.offset}, "
                    f"offset%{req}={p.offset % req})"
                )

    return AssemblyIntegrityResult(
        eds_total_bytes    = eds_total,
        actual_bytes       = actual_bytes,
        param_count        = unique_param_count,
        out_of_bounds      = oob,
        alignment_issues   = align,
        per_assembly_sizes = per_assembly_sizes,
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
                instance=10,
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


# ─── Forward_Open 隐式连接建立测试 ──────────────────────────────────────────────

@dataclass
class ForwardOpenResult:
    """
    Forward_Open / Large_Forward_Open 隐式连接建立测试结果。

    验证目的：
      1. 设备支持 Forward_Open 服务（而非仅显式消息）
      2. EDS Connection1 修复后隐式连接能否真正建立（弥补显式消息绕过 Connection Manager 的盲区）
    """
    attempted:     bool = False   # 是否已发送请求
    connected:     bool = False   # 是否收到成功响应
    service_used:  str  = ""      # "LargeForwardOpen(0x5B)"
    error_message: str  = ""      # 失败原因
    t_o_api_us:    int  = 0       # T→O 实际包间隔（μs），成功时有效
    note:          str  = ""      # 补充说明

    @property
    def ok(self) -> bool:
        return self.connected


def _do_forward_close(
    driver,
    conn_serial: int,
    orig_serial: int,
    conn_path: bytes,
    conn_path_size: int,
) -> None:
    """
    发送 Forward_Close (0x4E) 断开已建立的隐式连接。
    失败时仅记录警告，不向上抛异常（连接超时后设备会自动清理）。
    """
    req = struct.pack(
        "<BB HH I BB",
        0x07,           # Priority_And_TimeTick
        5,              # Time_Out_Ticks
        conn_serial,    # Connection Serial Number（须与 Forward_Open 一致）
        0x0001,         # Originator Vendor ID
        orig_serial,    # Originator Serial Number
        conn_path_size, # Connection Path Size（words）
        0x00,           # Reserved
    ) + conn_path
    try:
        driver.generic_message(
            service=b"\x4E",
            class_code=0x06,
            instance=0x01,
            request_data=req,
            connected=False,
            route_path=True,
        )
        log.info("Forward_Close 已发送，连接已清理")
    except Exception as exc:
        log.warning("Forward_Close 异常（连接将在设备侧超时后自动清理）: %s", exc)


def test_forward_open(driver) -> ForwardOpenResult:
    """
    发送 Large_Forward_Open (0x5B) 到 Assembly 100，验证设备能建立隐式 I/O 连接。

    连接参数（Input-Only 模式，不向设备写数据）：
      - O→T: null 连接（size=0, type=null）
      - T→O: Point-to-Point, 低优先级, 固定大小, 5656 字节, RPI=500ms
      - 连接路径: Class=0x04 Config=12 O→T=11 T→O=10（与 EDS Path "20 04 24 0C 24 0B 24 0A" 一致）

    成功后立即发送 Forward_Close 断开，不影响设备运行状态。
    若设备 EDS Connection1 格式不合规，预期此处返回 CIP 错误 0x01
    （Connection_Failure），而修复后应返回成功。
    """
    import random

    result = ForwardOpenResult(attempted=True, service_used="LargeForwardOpen(0x5B)")

    conn_serial = random.randint(0x1000, 0xEFFF)
    ot_conn_id  = random.randint(0x11000000, 0x1FFFFFFF)
    to_conn_id  = random.randint(0x21000000, 0x2FFFFFFF)
    orig_serial = 0x00414343   # "ACC" — 标识本工具为发起方

    # 连接路径：与 EDS [Connection Manager] Connection1 的 Path 字段一致
    # "20 04 24 0C 24 0B 24 0A" → Class 4, Config=Instance 12, O→T=Instance 11, T→O=Instance 10
    conn_path      = bytes([0x20, 0x04, 0x24, 0x0C, 0x24, 0x0B, 0x24, 0x0A])
    conn_path_size = len(conn_path) // 2               # 4 words

    # T→O 网络连接参数（4字节 Large 格式）：
    #   bits 29-30 = 10（Point-to-Point）
    #   bits 0-16  = 5656（fixed size，字节数）
    to_params = (0b10 << 29) | 5656

    # O→T 网络连接参数（4字节 Large 格式）：null 连接
    ot_params = 0x00000000

    req = struct.pack(
        "<BB II HH IB3x IIII BB",
        0x07,           # Priority_And_TimeTick（tick=7, priority=low）
        5,              # Time_Out_Ticks（5 × 128ms = 640ms）
        ot_conn_id,     # O→T Network Connection ID
        to_conn_id,     # T→O Network Connection ID
        conn_serial,    # Connection Serial Number
        0x0001,         # Originator Vendor ID
        orig_serial,    # Originator Serial Number
        0x00,           # Connection Timeout Multiplier
        # 3 bytes padding（reserved）
        500_000,        # O→T RPI（μs）
        ot_params,      # O→T Network Connection Parameters（Large）
        500_000,        # T→O RPI（μs）
        to_params,      # T→O Network Connection Parameters（Large）
        0x01,           # Transport Type/Trigger：class=1, cyclic, client
        conn_path_size, # Connection Path Size（words）
    ) + conn_path

    try:
        resp = driver.generic_message(
            service=b"\x5B",   # Large_Forward_Open
            class_code=0x06,   # Connection Manager
            instance=0x01,
            request_data=req,
            connected=False,
            route_path=True,
        )

        if resp and resp.value is not None:
            raw = (resp.value if isinstance(resp.value, (bytes, bytearray))
                   else bytes(resp.value))
            # Large_Forward_Open_Reply:
            #   O→T ID(4), T→O ID(4), Serial(2), VendorID(2),
            #   Originator Serial(4), O→T API(4), T→O API(4),
            #   App Reply Size(1), Reserved(1)  →  最少 26 字节
            if len(raw) >= 26:
                ot_id, to_id       = struct.unpack_from("<II", raw, 0)
                ot_api_us, to_api_us = struct.unpack_from("<II", raw, 16)
                result.connected   = True
                result.t_o_api_us  = to_api_us
                result.note = (
                    f"隐式连接建立成功 | T→O_API={to_api_us}μs"
                    f" | T→O_ID=0x{to_id:08X}"
                )
                log.info("Forward_Open 成功：T→O_API=%d μs  T→O_ID=0x%08X",
                         to_api_us, to_id)
                _do_forward_close(driver, conn_serial, orig_serial,
                                  conn_path, conn_path_size)
            else:
                result.error_message = f"响应体过短 ({len(raw)} 字节)"
                result.note = result.error_message
        else:
            err_str = str(getattr(resp, "error", "") or "无响应")
            result.error_message = err_str
            result.note = f"设备拒绝 Forward_Open: {err_str}"
            log.warning("Forward_Open 被拒绝: %s", err_str)

    except Exception as exc:
        result.error_message = str(exc)
        result.note = f"发送异常: {exc}"
        log.error("Forward_Open 异常: %s", exc)

    return result


# ─── EtherNet/IP 读取器 ───────────────────────────────────────────────────────

class EnipReader:
    """
    使用 pycomm3 读取 AcuRev-4100 EtherNet/IP Assembly Instance 100。

    用法：
        async with EnipReader() as reader:
            results = await reader.read_all()
    """

    ASSEMBLY_INSTANCE = 10
    ASSEMBLY_CLASS    = 0x04
    ASSEMBLY_ATTR     = 3

    def __init__(self, host: str, eds_path: str, slot: int = 0,
                 assem_ids: Optional[list[int]] = None):
        self.host          = host
        self.eds_path      = eds_path
        self.slot          = slot
        self.assem_ids     = assem_ids  # None = 解析 EDS 全部输入 Assembly
        self._params:      list[EnipParam] = []
        self._driver       = None
        self._actual_bytes: int = 0   # 最近一次 Assembly 读取的实际字节数之和

    async def __aenter__(self) -> "EnipReader":
        from pycomm3 import CIPDriver
        self._params = parse_eds(self.eds_path, self.assem_ids)
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
        """同步读取全部参数（逐 Assembly 实例分别读取），返回 {param_key: EnipResult}。"""
        results: dict[str, EnipResult] = {}
        total_bytes = 0

        # 按 Assembly 实例分组
        by_instance: dict[int, list[EnipParam]] = {}
        for p in self._params:
            by_instance.setdefault(p.assembly_instance, []).append(p)

        for assem_id in sorted(by_instance.keys()):
            params_grp = by_instance[assem_id]
            try:
                resp = self._driver.generic_message(
                    service=b"\x0E",              # Get_Attribute_Single
                    class_code=self.ASSEMBLY_CLASS,
                    instance=assem_id,
                    attribute=self.ASSEMBLY_ATTR,
                    connected=False,
                    route_path=True,
                )
                if not resp:
                    raise RuntimeError(f"读取 Assembly {assem_id} 失败：{resp.error}")

                raw: bytes = (resp.value if isinstance(resp.value, (bytes, bytearray))
                              else bytes(resp.value))
                total_bytes += len(raw)
                log.info("Assembly %d 读取成功，字节数: %d", assem_id, len(raw))
                values = parse_assembly_bytes(raw, params_grp)
                for name, val in values.items():
                    results[name] = EnipResult(param_key=name, value=val)

            except Exception as exc:
                log.error("Assembly %d 读取异常: %s", assem_id, exc)
                for p in params_grp:
                    results[p.name] = EnipResult(param_key=p.name, error=str(exc))

        self._actual_bytes = total_bytes
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
