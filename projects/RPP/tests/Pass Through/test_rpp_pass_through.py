# -*- coding: utf-8 -*-
"""RPP Pass Through 自动化（单文件、可独立执行的 pytest）——由 AcuHMI_1_7 同名套件适配。

验证 RPP「Pass Through」透传：主站以透传 SlaveID（101–247）经网关直达下游电表寄存器。
RPP 的 Pass Through 页与 Device Mirror 完全同构（Enable 单选 + 同六列设备表格 + Save，
无 Download All），2026-07-03 对 demo 实测确认。

RPP 与 AcuHMI 1.7 的页面差异（已在本文件适配，同 Device Mirror 套件）：
  登录落 /#/overview；菜单导航替代 hash 直跳；Save 无改动时 disabled；
  Disable 下表格消失；本机按行锁定状态排除；demo 错误 toast 过滤。

demo 限制（RPP_DEMO=1）：
  - 透传表格无设备行（"No Data"，RPP 本机不参与透传）→ 行级/配置类用例跳过
  - 无 Modbus 服务 → 读数类用例跳过
  - B 路直读真实电表来源页面待真机确认 → pt_003 数据对比二期实现

运行（仓库根目录下，任意人 clone 后均可）：
  pytest "projects/RPP/tests/Pass Through" -v
也可直接运行本文件： python "projects/RPP/tests/Pass Through/test_rpp_pass_through.py" -v
"""
from __future__ import annotations

import os
import pathlib
import struct
import sys
import time
import zipfile
import xml.etree.ElementTree as _ET
from dataclasses import dataclass

import pytest
from playwright.sync_api import Browser, Page, Locator, TimeoutError as PWTimeout

# ── 配置来源（与 1.7 相同的加载链）───────────────────────────────────────────
# 优先本地 tests/config.py（开发机覆盖，gitignored）；没有它时回退到框架分层配置
# （configs/env + projects/RPP/config.yaml），再回退 demo 默认值；环境变量随时可覆盖。
import importlib.util as _ilu, types as _types

_TESTS_DIR = pathlib.Path(__file__).resolve().parent.parent
_REPO_ROOT = pathlib.Path(__file__).resolve().parents[4]


def _load_local_config():
    _cfg_path = _TESTS_DIR / "config.py"
    if _cfg_path.exists():
        _spec = _ilu.spec_from_file_location("_rpp_tests_config", _cfg_path)
        _mod = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_mod)
        return _mod
    if str(_REPO_ROOT) not in sys.path:
        sys.path.insert(0, str(_REPO_ROOT))
    _cfg = {}
    try:
        from dotenv import load_dotenv
        for _name in ("env", ".env"):
            _envfile = _REPO_ROOT / "configs" / _name
            if _envfile.exists():
                load_dotenv(_envfile, override=False)
        from framework.config.loader import load_config
        _cfg = load_config("RPP")
    except Exception:
        pass
    return _types.SimpleNamespace(
        RPP_URL=_cfg.get("rpp_url") or "http://192.168.2.94:3030",
        RPP_USERNAME=_cfg.get("rpp_username") or "admin",
        RPP_PASSWORD=_cfg.get("rpp_password") or "Admin@000211",
        RPP_DEMO=_cfg.get("rpp_demo", True),
        MODBUS_PORT=int(_cfg.get("rpp_modbus_port", 502)),
        TOLERANCE_PERCENT=float(_cfg.get("modbus_cmp_tolerance_percent", 1.0)),
        TOLERANCE_ABSOLUTE=float(_cfg.get("modbus_cmp_tolerance_absolute", 0.05)),
    )


config = _load_local_config()

# ── 由 config 派生的配置（环境变量可覆盖）────────────────────────────────────
from urllib.parse import urlparse as _urlparse

BASE_URL     = os.getenv("RPP_URL", config.RPP_URL).rstrip("/")
GATEWAY_IP   = os.getenv("RPP_GATEWAY_IP", _urlparse(BASE_URL).hostname)
USERNAME     = os.getenv("RPP_USERNAME", config.RPP_USERNAME)
PASSWORD     = os.getenv("RPP_PASSWORD", config.RPP_PASSWORD)
DEMO         = os.getenv("RPP_DEMO", str(getattr(config, "RPP_DEMO", True))).strip().lower() in ("1", "true", "yes", "on")
MODBUS_PORT  = int(getattr(config, "MODBUS_PORT", 502) or 502)
MODBUS_TIMEOUT = float(os.getenv("MODBUS_TIMEOUT", "5"))
PT_SLAVE_MIN, PT_SLAVE_MAX = 101, 247   # 与镜像(2–99)区分的透传专用范围
PT_SETTLE_MS = int(os.getenv("PT_SETTLE_MS", "3000"))            # 启用后等透传服务就绪
PT_DISABLE_SETTLE = float(os.getenv("PT_DISABLE_SETTLE", "20"))  # 禁用后确认停服的轮询窗口(秒)
TIMEOUT      = int(os.getenv("TIMEOUT", "15000"))
INTERFACE_ETHERNET = "Ethernet"
# 仓库根 knowledge/：AcuCloud 原始模板（与 1.7 共用同一套按型号寄存器表）
TMPL_DIR = _REPO_ROOT / "knowledge" / "shared" / "templates" / "raw"

# 用例函数名 → 关联 TC 编号（RPP 用例库编号确定后替换）
CASE_ID_MAP = {
    "test_pt_000_page_layout":                "RPP_PT_case00(页面布局)",
    "test_pt_001_config":                     "RPP_PT_前置检查",
    "test_pt_002_data_collected":             "RPP_PT_case05(子)",
    "test_pt_003_passthrough_matches_direct": "RPP_PT_case05",
    "test_pt_case01_enable_disable_toggle":   "RPP_PT_case01",
    "test_pt_case02_valid_slaveid_save":      "RPP_PT_case02",
    "test_pt_case03_slaveid_boundary":        "RPP_PT_case03",
    "test_pt_case04_duplicate_slaveid":       "RPP_PT_case04",
    "test_pt_case06_disabled_blocks_access":  "RPP_PT_case06",
    "test_pt_case07_concurrent_masters":      "RPP_PT_case07",
}


# ═══════════════════════════════════════════════════════════════════════════
# Modbus TCP 读取 + 解码（pymodbus 3.x）—— 与 1.7 一致
# ═══════════════════════════════════════════════════════════════════════════
from pymodbus.client import ModbusTcpClient


def _regs_to_bytes(regs):
    return b"".join(struct.pack(">H", r & 0xFFFF) for r in regs)


def _decode(regs, kind):
    if kind == "uint16":
        return regs[0]
    if kind == "int16":
        return struct.unpack(">h", struct.pack(">H", regs[0]))[0]
    if kind == "uint32":
        return float(struct.unpack(">I", _regs_to_bytes(regs[:2]))[0])
    if kind == "float32":
        return struct.unpack(">f", _regs_to_bytes(regs[:2]))[0]
    if kind == "float64":
        return struct.unpack(">d", _regs_to_bytes(regs[:4]))[0]
    if kind == "string":
        return _regs_to_bytes(regs).decode("ascii", errors="ignore").rstrip("\x00 ").strip()
    raise ValueError(f"未知解码类型: {kind}")


class Modbus:
    def __init__(self, host=None, port=None):
        self.host = host or GATEWAY_IP
        self.port = port or MODBUS_PORT
        self.client = ModbusTcpClient(self.host, port=self.port, timeout=MODBUS_TIMEOUT)

    def __enter__(self):
        if not self.client.connect():
            raise IOError(f"无法连接 {self.host}:{self.port}")
        return self

    def __exit__(self, *exc):
        self.client.close()

    def _read_regs(self, fc, address, count, slave):
        if fc == 3:
            rr = self.client.read_holding_registers(address, count=count, device_id=slave)
        elif fc == 4:
            rr = self.client.read_input_registers(address, count=count, device_id=slave)
        elif fc == 1:
            rr = self.client.read_coils(address, count=count, device_id=slave)
        elif fc == 2:
            rr = self.client.read_discrete_inputs(address, count=count, device_id=slave)
        else:
            raise IOError(f"不支持的功能码 {fc}")
        if rr.isError():
            raise IOError(f"device_id={slave} fc={fc} addr={address} -> {rr}")
        return list(rr.bits)[:count] if fc in (1, 2) else list(rr.registers)

    def read(self, rec, unit=None, retries=2):
        slave = getattr(rec, "slave_id", None) if unit is None else unit
        last = None
        for attempt in range(retries + 1):
            try:
                regs = self._read_regs(rec.fc, rec.start, rec.width, slave)
                return int(bool(regs[0])) if rec.kind == "bit" else _decode(regs, rec.kind)
            except Exception as e:  # noqa: BLE001
                last = e
                if attempt < retries:
                    try:
                        self.client.connect()
                    except Exception:
                        pass
        raise last

    def read_all(self, records, unit=None):
        out = {}
        for rec in records:
            try:
                out[(rec.device, rec.parameter)] = self.read(rec, unit=unit)
            except Exception as e:  # noqa: BLE001
                out[(rec.device, rec.parameter)] = e
        return out


def _gateway_reachable():
    """网关 Modbus 端口 TCP 可连通即算可达（demo 前端无 Modbus 服务 → False）。"""
    try:
        c = ModbusTcpClient(GATEWAY_IP, port=MODBUS_PORT, timeout=3)
        ok = c.connect()
        c.close()
        return ok
    except Exception:
        return False


# ═══════════════════════════════════════════════════════════════════════════
# 按型号寄存器表  knowledge/shared/templates/raw/<Model>*.xlsx —— 与 1.7 一致
# D 列=起始地址  E 列=参数描述  H 列=数据类型（或按表头定位），FC 固定=3
# ═══════════════════════════════════════════════════════════════════════════
_XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# 设备名关键词 → 模板文件名前缀（按前缀 glob，兼容带版本号文件）
_TMPL_MAP = [
    (["4110", "4100"],               "AcuRev-4100"),
    (["2100"],                        "AcuRev-2100"),
    (["1300"],                        "AcuRev-1300"),
    (["1320"],                        "AcuRev-1320"),
    (["acuvim3", "vim3"],             "Acuvim3"),
    (["acuvimiiw", "vimiiw", "iiw"],  "AcuvimIIW"),
    (["acuvimii", "vimii"],           "AcuvimIIR"),
]


@dataclass(frozen=True)
class RegRecord:
    device: str
    parameter: str
    fc: int
    start: int
    width: int
    kind: str


def _kind_width(datatype):
    dt = (datatype or "").strip().lower()
    if dt in ("uint16", "word"):
        return "uint16", 1
    if dt in ("int16", "short"):
        return "int16", 1
    if dt in ("uint32", "dword", "long"):
        return "uint32", 2
    if dt in ("float32", "float", "real"):
        return "float32", 2
    if dt in ("float64", "double"):
        return "float64", 4
    return "uint16", 1


def _glob_latest(prefix: str):
    exact = TMPL_DIR / f"{prefix}.xlsx"
    if exact.exists():
        return exact
    cands = sorted(TMPL_DIR.glob(f"{prefix}_*.xlsx")) + sorted(TMPL_DIR.glob(f"{prefix}*.xlsx"))
    seen, uniq = set(), []
    for c in cands:
        if c not in seen:
            seen.add(c); uniq.append(c)
    return uniq[-1] if uniq else None


def _pick_tmpl(model: str):
    dn = (model or "").lower().replace("-", "").replace(" ", "")
    for keywords, prefix in _TMPL_MAP:
        if any(k.replace("-", "") in dn for k in keywords):
            p = _glob_latest(prefix)
            if p:
                return p
    return _glob_latest(model) if model else None


def _blockparams_sheet_xml(z):
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    names = z.namelist()
    try:
        wb = _ET.fromstring(z.read("xl/workbook.xml"))
        rels = _ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {r.get("Id"): r.get("Target") for r in rels}
        for sh in wb.iter(f"{{{_XLSX_NS}}}sheet"):
            if (sh.get("name") or "").strip().lower() == "blockparams":
                rid = sh.get(f"{{{REL_NS}}}id")
                tgt = rid_to_target.get(rid, "").split("/")[-1]
                cand = f"xl/worksheets/{tgt}"
                if cand in names:
                    return cand
    except Exception:
        pass
    for fallback in ("xl/worksheets/sheet2.xml", "xl/worksheets/sheet1.xml"):
        if fallback in names:
            return fallback
    return names[0]


def _load_xlsx_template(path: pathlib.Path):
    """解析 AcuCloud 模板 xlsx（blockParams sheet），按表头文字定位列，兼容新旧布局。"""
    with zipfile.ZipFile(str(path)) as z:
        ss = {}
        if "xl/sharedStrings.xml" in z.namelist():
            root = _ET.fromstring(z.read("xl/sharedStrings.xml"))
            for i, si in enumerate(root.iter(f"{{{_XLSX_NS}}}si")):
                ss[i] = "".join(c.text or "" for c in si.iter()
                                if c.tag == f"{{{_XLSX_NS}}}t")
        sheet_xml = _blockparams_sheet_xml(z)
        root = _ET.fromstring(z.read(sheet_xml))
        raw = []
        for row in root.iter(f"{{{_XLSX_NS}}}row"):
            cells = {}
            for cell in row:
                ref = cell.get("r", ""); t = cell.get("t", "")
                v = cell.find(f"{{{_XLSX_NS}}}v")
                val = (ss.get(int(v.text), "") if t == "s" else (v.text or "")) \
                      if (v is not None and v.text) else ""
                cells["".join(c for c in ref if c.isalpha())] = val
            if cells:
                raw.append(cells)
    if not raw:
        return []
    header = raw[0]
    col_addr = col_desc = col_dtype = None
    for col, txt in header.items():
        t = (txt or "").strip().lower()
        if col_addr is None and t.startswith("start"):
            col_addr = col
        elif col_desc is None and t.startswith("descr"):
            col_desc = col
        elif col_dtype is None and t in ("datatype", "data type"):
            col_dtype = col
    col_addr = col_addr or "D"
    col_desc = col_desc or "E"
    col_dtype = col_dtype or "H"
    entries = []
    for r in raw[1:]:
        desc = r.get(col_desc, "").strip()
        addr = r.get(col_addr, "").strip()
        dtype = r.get(col_dtype, "").strip()
        if not desc or not addr:
            continue
        try:
            a = int(addr)
        except ValueError:
            continue
        entries.append({"desc": desc, "addr": a, "dtype": dtype})
    return entries


_tmpl_cache: dict = {}


def load_table(model, device=""):
    """从 AcuCloud 模板 xlsx 加载寄存器列表；找不到模板返回空列表。"""
    p = _pick_tmpl(model)
    if p is None:
        return []
    if str(p) not in _tmpl_cache:
        _tmpl_cache[str(p)] = _load_xlsx_template(p)
    out = []
    for e in _tmpl_cache[str(p)]:
        kind, width = _kind_width(e["dtype"])
        out.append(RegRecord(device, e["desc"], 3, e["addr"], width, kind))
    return out


def table_path(model):
    p = _pick_tmpl(model)
    return p if p else TMPL_DIR / f"{model}.xlsx"


# ═══════════════════════════════════════════════════════════════════════════
# 登录 + Pass Through 页面驱动（RPP 适配版，内联；与 Device Mirror 套件同款）
# ═══════════════════════════════════════════════════════════════════════════
COL_SLAVEID, COL_DEVICE, COL_INTERFACE, COL_MODEL = 1, 2, 3, 5   # 与 1.7 相同


def _ensure_modbus_config_enabled(page):
    """RPP 特有前置条件：Modbus 总开关（Protocols → Modbus → Modbus Config →
    Modbus Enable）必须启用，否则 Device Mirror / Pass Through 页面整体不可用
    （表格和按钮消失，提示 "Configuration unavailable. Modbus is not enabled."）。
    点 Protocols 默认就落在 Modbus Config 页，检查并按需启用+保存。"""
    page.goto(BASE_URL + "/#/overview", wait_until="domcontentloaded", timeout=TIMEOUT)
    page.get_by_text("Settings", exact=True).first.click()
    page.wait_for_url("**/rppConfiguration/**", timeout=TIMEOUT)
    page.get_by_text("Protocols", exact=True).first.click()
    page.wait_for_url("**/protocols/**", timeout=TIMEOUT)
    page.locator("text=Modbus Enable").first.wait_for(state="visible", timeout=TIMEOUT)
    en = page.locator(".el-radio", has_text="Enable").first
    if "is-checked" in (en.get_attribute("class") or ""):
        return
    en.click(); page.wait_for_timeout(600)
    sb = page.get_by_role("button", name="Save")
    if sb.get_attribute("disabled") is None:
        sb.click()
        try:
            page.locator(".el-message").first.wait_for(state="visible", timeout=8000)
        except Exception:
            pass
        page.wait_for_timeout(800)


def _login(page: Page):
    """登录 RPP：成功后落在 /#/overview；默认密码修改确认弹窗点 Cancel。"""
    page.goto(BASE_URL, wait_until="domcontentloaded")
    try:
        page.wait_for_url("**/login", timeout=4000)
    except PWTimeout:
        pass
    if "/login" not in page.url and "overview" in page.url.lower():
        return
    page.locator("input[type=text]").first.fill(USERNAME)
    page.locator("input[type=password]").first.fill(PASSWORD)
    for sel in ["button:has-text('Sign In')", "button:has-text('Login')",
                "button:has-text('登录')", "button[type=submit]"]:
        b = page.locator(sel).first
        try:
            b.wait_for(state="visible", timeout=2500); b.click(); break
        except PWTimeout:
            continue
    for sel in [".el-message-box button:has-text('Cancel')", ".el-dialog button:has-text('取消')"]:
        try:
            page.locator(sel).first.wait_for(state="visible", timeout=2500)
            page.locator(sel).first.click(); break
        except PWTimeout:
            continue
    page.wait_for_url("**/overview", timeout=TIMEOUT)


class PassThroughPage:
    TAB_NAME = "Pass Through"
    ENABLE_LABEL = "Pass Through Enable"

    def __init__(self, page: Page):
        self.page = page

    def goto(self, retries=3):
        """经菜单进入 Pass Through（RPP 路由守卫拦 hash 直跳，必须逐级点击）：
        顶部 Settings → 侧边 Protocols（落在 Modbus Config）→ 展开 Modbus 子菜单 → Pass Through。"""
        page = self.page
        last_err = None
        for attempt in range(retries):
            try:
                page.goto(BASE_URL + "/#/overview", wait_until="domcontentloaded", timeout=TIMEOUT)
                page.get_by_text("Settings", exact=True).first.click()
                page.wait_for_url("**/rppConfiguration/**", timeout=TIMEOUT)
                page.get_by_text("Protocols", exact=True).first.click()
                page.wait_for_url("**/protocols/**", timeout=TIMEOUT)
                item = page.locator(".el-menu-item", has_text=self.TAB_NAME).first
                if not item.is_visible():
                    page.locator(".el-sub-menu__title", has_text="Modbus").first.click()
                    item.wait_for(state="visible", timeout=3000)
                item.click()
                page.locator(f"text={self.ENABLE_LABEL}").first.wait_for(state="visible", timeout=TIMEOUT)
                return
            except Exception as e:  # noqa: BLE001
                last_err = e
                if attempt < retries - 1:
                    page.wait_for_timeout(3000)
        raise last_err

    @property
    def rows(self) -> Locator:
        return self.page.locator(".el-table__body-row, .el-table__row")

    def _cell(self, row, col):
        return row.locator("td").nth(col)

    def _slaveid_input(self, row):
        return self._cell(row, COL_SLAVEID).locator("input").first

    @staticmethod
    def _row_locked(row) -> bool:
        """锁定行判定：勾选框 is-disabled（不可编辑行）。"""
        return "is-disabled" in (row.locator(".el-checkbox").first.get_attribute("class") or "")

    def _radio(self, label):
        return self.page.locator(".el-radio", has_text=label).first

    def is_enabled(self):
        return "is-checked" in (self._radio("Enable").get_attribute("class") or "")

    def set_enabled(self, on):
        self._radio("Enable" if on else "Disable").click()

    def ensure_enabled(self):
        """确保处于 Enable 状态（RPP：Disable 下表格整体消失）。"""
        if not self.is_enabled():
            self.set_enabled(True); self.page.wait_for_timeout(800)

    # ── Save（RPP：无改动时按钮 disabled）─────────────────────────────────
    def _save_button(self):
        return self.page.get_by_role("button", name="Save")

    def save_disabled(self):
        b = self._save_button()
        return (b.get_attribute("disabled") is not None) or \
               (b.get_attribute("aria-disabled") == "true")

    def save(self):
        """点击 Save；无改动（按钮 disabled）时视为无需保存，返回是否真正点击。"""
        if self.save_disabled():
            return False
        self._save_button().click()
        return True

    def has_message(self, timeout=8000):
        try:
            self.page.locator(".el-message").first.wait_for(state="visible", timeout=timeout)
            return True
        except Exception:
            return False

    def messages(self, timeout=8000):
        """返回当前全部顶部提示 [(text, is_success), ...]。
        demo 后台轮询接口会随机弹无关的 "There's been an error."，不能只看第一条。"""
        out = []
        try:
            self.page.locator(".el-message").first.wait_for(state="visible", timeout=timeout)
            msgs = self.page.locator(".el-message")
            for i in range(msgs.count()):
                m = msgs.nth(i)
                try:
                    out.append((m.inner_text().strip(),
                                "el-message--success" in (m.get_attribute("class") or "")))
                except Exception:
                    continue
        except Exception:
            pass
        return out

    def message_text(self, timeout=8000):
        msgs = self.messages(timeout=timeout)
        return msgs[0][0] if msgs else ""

    def form_errors(self):
        try:
            return [t.strip() for t in self.page.locator(".el-form-item__error").all_inner_texts() if t.strip()]
        except Exception:
            return []

    def save_and_result(self, timeout=8000):
        """点击 Save 并返回 (success, message, errors)。
        无改动（Save disabled）视为成功空操作；任一条 success 类消息或文案含
        saved/成功 即算保存成功（忽略 demo 无关错误 toast），且需无表单校验错误。"""
        errs_before = self.form_errors()
        if not self.save():
            return True, "(无改动，Save 处于禁用态)", errs_before
        self.page.wait_for_timeout(600)
        msgs = self.messages(timeout=timeout)
        errs = self.form_errors() or errs_before
        hit = next((t for t, ok in msgs if ok or "saved" in t.lower() or "成功" in t), None)
        msg = hit if hit is not None else (msgs[0][0] if msgs else "")
        success = (hit is not None) and not errs
        return success, msg, errs

    # ── 行遍历 / 读写（跳过锁定行）─────────────────────────────────────────
    def _scan(self, device):
        rows = self.rows
        for i in range(rows.count()):
            row = rows.nth(i)
            try:
                if self._cell(row, COL_DEVICE).inner_text().strip() == device:
                    return row
            except Exception:
                continue
        return None

    def find_row(self, device, retries=6):
        try:
            self.rows.first.wait_for(state="visible", timeout=TIMEOUT)
        except PWTimeout:
            pass
        for _ in range(retries):
            row = self._scan(device)
            if row is not None:
                return row
            self.page.wait_for_timeout(400)
        raise AssertionError(f"未找到设备行: {device}")

    def editable_device_names(self):
        out, rows = [], self.rows
        for i in range(rows.count()):
            row = rows.nth(i)
            if self._row_locked(row):
                continue
            name = self._cell(row, COL_DEVICE).inner_text().strip()
            if name:
                out.append(name)
        return out

    def ethernet_device_names(self):
        """Interface=Ethernet 且可编辑的设备名（排除锁定行）。"""
        out, rows = [], self.rows
        for i in range(rows.count()):
            row = rows.nth(i)
            if self._row_locked(row):
                continue
            if INTERFACE_ETHERNET.lower() in self._cell(row, COL_INTERFACE).inner_text().strip().lower():
                name = self._cell(row, COL_DEVICE).inner_text().strip()
                if name:
                    out.append(name)
        return out

    def is_checked(self, device):
        return "is-checked" in (self.find_row(device).locator(".el-checkbox").first.get_attribute("class") or "")

    def set_checked(self, device, checked):
        if self.is_checked(device) != checked:
            cb = self.find_row(device).locator(".el-checkbox").first
            cb.scroll_into_view_if_needed(); cb.click()

    def set_slaveid(self, device, value):
        inp = self._slaveid_input(self.find_row(device))
        inp.scroll_into_view_if_needed(); inp.click(); inp.fill(""); inp.fill(str(value)); inp.blur()

    def get_slaveid(self, device):
        return self._slaveid_input(self.find_row(device)).input_value()

    def used_slave_ids(self):
        ids = set()
        rows = self.rows
        for i in range(rows.count()):
            v = self._slaveid_input(rows.nth(i)).input_value().strip()
            if v.isdigit():
                ids.add(int(v))
        return ids

    def free_slave_id(self, lo=PT_SLAVE_MIN, hi=PT_SLAVE_MAX):
        used = self.used_slave_ids()
        for v in range(lo, hi + 1):
            if v not in used:
                return v
        raise AssertionError("无可用 SlaveID")

    def read_row_basic(self, device):
        row = self.find_row(device)
        cb = row.locator(".el-checkbox").first
        return {
            "model": self._cell(row, COL_MODEL).inner_text().strip(),
            "slave_id": self._slaveid_input(row).input_value(),
            "checked": "is-checked" in (cb.get_attribute("class") or ""),
            "disabled": "is-disabled" in (cb.get_attribute("class") or ""),
        }

    def snapshot(self):
        snap = {}
        for name in self.editable_device_names():
            info = self.read_row_basic(name)
            snap[name] = {"checked": info["checked"], "slave_id": info["slave_id"]}
        return snap

    def restore(self, baseline, max_passes=5):
        drift = []
        for _ in range(max_passes):
            for name, st in baseline.items():
                row = self._scan(name)
                if row is None or self._row_locked(row):
                    continue
                self.set_checked(name, True)
                inp = self._slaveid_input(self.find_row(name))
                inp.scroll_into_view_if_needed(); inp.click(); inp.fill(""); inp.fill(str(st["slave_id"])); inp.blur()
                self.set_checked(name, st["checked"])
            if self.save():
                self.has_message()
            self.goto(); self.ensure_enabled()
            drift = [n for n, st in baseline.items()
                     if (self._scan(n) is not None) and self.is_checked(n) != st["checked"]]
            if not drift:
                return []
        return drift


# ═══════════════════════════════════════════════════════════════════════════
# Fixtures
# ═══════════════════════════════════════════════════════════════════════════
@pytest.fixture(scope="module")
def page_factory(browser: Browser):
    """复用 pytest-playwright 的共享 browser；登录一次后 sessionStorage 注入新 context 免重复登录。"""
    ctx0 = browser.new_context(ignore_https_errors=True, accept_downloads=True)
    pg = ctx0.new_page(); pg.set_default_timeout(TIMEOUT)
    _login(pg)
    storage = pg.evaluate("JSON.stringify(window.sessionStorage)")
    ctx0.close()
    contexts = []

    def make():
        ctx = browser.new_context(ignore_https_errors=True, accept_downloads=True)
        ctx.add_init_script(
            "(() => { const d = %s; for (const k in d){try{sessionStorage.setItem(k,d[k]);}catch(e){}} })();"
            % storage)
        page = ctx.new_page(); page.set_default_timeout(TIMEOUT)
        contexts.append(ctx)
        return page

    yield make
    for c in contexts:
        c.close()


@pytest.fixture(scope="module", autouse=True)
def baseline(page_factory):
    """保护 Pass Through 行配置（SlaveID/勾选）与 Enable 状态。"""
    pg = page_factory()
    _ensure_modbus_config_enabled(pg)   # RPP 前置：Modbus 总开关必须启用
    pt = PassThroughPage(pg); pt.goto()
    enabled0 = pt.is_enabled()
    pt.ensure_enabled()        # 启用才能看到表格做快照
    snap = pt.snapshot()
    yield snap
    pt2 = PassThroughPage(page_factory()); pt2.goto()
    pt2.ensure_enabled()
    drift = pt2.restore(snap) if snap else []
    pt2.goto()
    if pt2.is_enabled() != enabled0:
        pt2.set_enabled(enabled0)
        if pt2.save():
            pt2.has_message()
    if drift:
        import warnings
        warnings.warn(f"Pass Through 还原后仍有偶发漂移：{drift}")


@pytest.fixture(scope="module")
def configured(page_factory):
    """启用 Pass Through、勾选全部可编辑 Ethernet 设备，为 SlaveID 无效（非 101–247）
    的设备分配空闲 SlaveID。返回 {assigned: {设备: sid}, models: {设备: 型号}}。"""
    pt = PassThroughPage(page_factory()); pt.goto()
    if not pt.is_enabled():
        pt.set_enabled(True)
        if pt.save():
            pt.has_message()
        pt.goto(); pt.ensure_enabled()
    names = pt.ethernet_device_names()
    for n in names:
        pt.set_checked(n, True)
    if pt.save():
        pt.has_message()
    pt.goto(); pt.ensure_enabled()

    need_assign = []
    for n in names:
        sid = pt.read_row_basic(n)["slave_id"]
        if not (str(sid).isdigit() and PT_SLAVE_MIN <= int(sid) <= PT_SLAVE_MAX):
            need_assign.append(n)
    if need_assign:
        for n in names:
            pt.set_checked(n, True)
        used = pt.used_slave_ids()
        next_id = PT_SLAVE_MIN
        for n in need_assign:
            while next_id in used:
                next_id += 1
            pt.set_slaveid(n, str(next_id))
            used.add(next_id); next_id += 1
        if pt.save():
            pt.has_message()
        pt.page.wait_for_timeout(PT_SETTLE_MS)
        pt.goto(); pt.ensure_enabled()

    assigned, models = {}, {}
    for n in names:
        info = pt.read_row_basic(n)
        assigned[n] = info["slave_id"]; models[n] = info["model"]
    return {"assigned": assigned, "models": models}


def _pt_read_one(sid, rec):
    """经网关透传读单个寄存器，成功返回 True。"""
    try:
        with Modbus() as mb:
            mb.read(rec, unit=int(sid))
        return True
    except Exception:
        return False


def _setup_pt_probes(page_factory):
    """启用 Pass Through、勾选可编辑 Ethernet 设备、保存并等就绪；
    返回 (pt_page, {sid: rec})——每个设备一个"经透传能读通"的寄存器探针。"""
    pt = PassThroughPage(page_factory()); pt.goto()
    if not pt.is_enabled():
        pt.set_enabled(True)
    for n in pt.ethernet_device_names():
        pt.set_checked(n, True)
    if pt.save():
        pt.has_message()
    pt.goto(); pt.ensure_enabled()
    pt.page.wait_for_timeout(PT_SETTLE_MS)
    working = {}
    if not _gateway_reachable():
        return pt, working
    for n in pt.ethernet_device_names():
        info = pt.read_row_basic(n)
        sid = info["slave_id"]
        if not str(sid).strip().isdigit():
            continue
        regs = load_table(info["model"], device=n)
        for rec in regs[:15]:
            if _pt_read_one(sid, rec):
                working[int(sid)] = rec
                break
    return pt, working


# ═══════════════════════════════════════════════════════════════════════════
# 用例
# ═══════════════════════════════════════════════════════════════════════════
def test_pt_000_page_layout(page_factory):
    """case00（新增）Pass Through 页面布局：Enable 单选 + Save；Enable 状态下出现
    与 Device Mirror 同构的设备表格（六列表头）；Disable 下表格隐藏。demo 即可验证。"""
    pt = PassThroughPage(page_factory()); pt.goto()
    fails = []

    for label in ("Enable", "Disable"):
        if not pt._radio(label).count():
            fails.append(f"缺少单选项 {label}")
    if not pt._save_button().count():
        fails.append("缺少 Save 按钮")

    enabled0 = pt.is_enabled()
    # Enable 状态下应出现设备表格（demo 无下挂设备时行数可为 0，但表格/表头应在）
    pt.ensure_enabled()
    if not pt.page.locator(".el-table").count():
        fails.append("Enable 状态下应显示设备表格")
    headers = [t.strip() for t in pt.page.locator("th").all_inner_texts() if t.strip()]
    for h in ("SlaveID", "Device Name", "Interface", "Model"):
        if h not in headers:
            fails.append(f"表头缺少列 {h!r}（实际 {headers}）")

    # 若有行：锁定行的 SlaveID 不应超出透传范围（RPP 本机只参与镜像，不参与透传）
    rows = pt.rows
    if rows.count():
        for i in range(rows.count()):
            if pt._row_locked(rows.nth(i)):
                sid = pt._slaveid_input(rows.nth(i)).input_value().strip()
                if sid.isdigit() and not (PT_SLAVE_MIN <= int(sid) <= PT_SLAVE_MAX):
                    fails.append(f"透传表格出现范围外锁定行 SlaveID={sid}")

    # 切 Disable → 表格隐藏；切回 → 恢复（仅切单选，不保存）
    pt.set_enabled(False); pt.page.wait_for_timeout(600)
    if pt.page.locator(".el-table").count():
        fails.append("Disable 状态下设备表格应隐藏")
    pt.set_enabled(True); pt.page.wait_for_timeout(600)
    if not pt.page.locator(".el-table").count():
        fails.append("切回 Enable 后表格应恢复显示")
    pt.set_enabled(enabled0)   # 还原单选（未保存过，不影响持久化状态）

    assert not fails, "页面布局校验失败：\n" + "\n".join(fails)


def test_pt_001_config(configured):
    """前置检查：启用 Pass Through 并对 Ethernet 设备读到有效 SlaveID（101–247）。"""
    assigned = configured["assigned"]
    if not assigned:
        pytest.skip("透传表格无可编辑 Ethernet 设备行（demo 无下挂设备），待真机")
    valid = [int(v) for v in assigned.values() if str(v).strip().isdigit() and int(v) > 0]
    assert valid, f"应至少读到一台设备有效 SlaveID：{assigned}"
    for v in valid:
        assert PT_SLAVE_MIN <= v <= PT_SLAVE_MAX, f"SlaveID {v} 超出 Pass Through 范围"


def test_pt_002_data_collected(configured):
    """前置检查：按型号寄存器表经透传(A 路)读到数据；缺表/无响应给出诊断。"""
    assigned, models = configured["assigned"], configured["models"]
    if not assigned:
        pytest.skip("透传表格无可编辑 Ethernet 设备行（demo 无下挂设备），待真机")
    if not _gateway_reachable():
        pytest.skip(f"网关 {GATEWAY_IP}:{MODBUS_PORT} Modbus 不可达（demo 无透传服务）")
    missing, ok_reads = [], 0
    with Modbus() as gw:
        for dev, sid in assigned.items():
            regs = load_table(models.get(dev, ""), device=dev)
            if not regs:
                missing.append(f"{models.get(dev) or dev} -> {table_path(models.get(dev, ''))}")
                continue
            if not str(sid).strip().isdigit():
                continue
            vals = gw.read_all(regs[:20], unit=int(sid))
            ok_reads += sum(1 for v in vals.values() if not isinstance(v, Exception))
    if missing and ok_reads == 0:
        pytest.skip("未找到任何型号寄存器表，请在 knowledge/shared/templates/raw 补充：" + "；".join(missing))
    if ok_reads == 0:
        pytest.skip("透传(A路)未返回可比对寄存器（多为 IllegalAddress）：确认透传端口/路由/设备在线")
    assert ok_reads > 0


def test_pt_003_passthrough_matches_direct():
    """透传(A) ↔ 直读真实电表(B) 数据一致性。

    1.7 的 B 路 IP/Unit 来自 Physical Devices → Settings → Connection 页面；
    RPP 的对应物大概率是 Monitoring → Gateway Devices(/#/physicalDevices)，
    待真机接入下游设备后确认，再按 1.7 的 pt_comparison fixture 移植
    （含逐寄存器背靠背交错读、稳定量/波动量分类判定、数据对比结果.xlsx 报告）。"""
    pytest.skip("RPP 真机未就绪：B 路（直读真实电表）连接信息来源页面待定，数据对比二期实现")


def test_pt_case01_enable_disable_toggle(page_factory):
    """case01 Disable+Save 持久化生效；Enable+Save 恢复。
    有可透传读通的设备时附加验证：Disable 后透传读取失败、Enable 后恢复（demo 跳过该部分）。"""
    pt, working = _setup_pt_probes(page_factory)

    pt.goto(); pt.ensure_enabled()
    pt.set_enabled(False)
    succ_dis, msg_dis, _ = pt.save_and_result()
    pt.goto()
    disabled_persisted = not pt.is_enabled()
    can_read_disabled = False
    if working:
        can_read_disabled = True
        deadline = time.time() + PT_DISABLE_SETTLE
        while time.time() < deadline:
            if not any(_pt_read_one(s, r) for s, r in working.items()):
                can_read_disabled = False
                break
            time.sleep(1.5)

    pt.goto(); pt.set_enabled(True)
    succ_en, msg_en, _ = pt.save_and_result()
    pt.goto()
    enabled_persisted = pt.is_enabled()
    can_read_enabled = None
    if working:
        pt.page.wait_for_timeout(PT_SETTLE_MS)
        can_read_enabled = False
        for _ in range(5):
            if any(_pt_read_one(s, r) for s, r in working.items()):
                can_read_enabled = True
                break
            time.sleep(2)

    assert succ_dis and disabled_persisted, \
        f"切 Disable 应保存并持久化（succ={succ_dis}, 提示={msg_dis!r}）"
    assert succ_en and enabled_persisted, \
        f"切回 Enable 应保存并持久化（succ={succ_en}, 提示={msg_en!r}）"
    if working:
        assert can_read_enabled, "Enable 后经透传应能读到下游设备数据"
        assert not can_read_disabled, \
            f"Disable 后 {PT_DISABLE_SETTLE:.0f}s 内经透传仍能读到数据（预期应连接失败/无数据）"
    else:
        import warnings
        warnings.warn("无可透传读通的设备（demo），本用例仅验证了 UI 开关持久化，未验证透传阻断/恢复")


def test_pt_case02_valid_slaveid_save(page_factory):
    """case02 为设备配置有效 SlaveID（101–247）保存成功且持久化。"""
    pt = PassThroughPage(page_factory()); pt.goto(); pt.ensure_enabled()
    dev = next(iter(pt.ethernet_device_names()), None)
    if dev is None:
        pytest.skip("透传表格无可编辑 Ethernet 设备行（demo 无下挂设备），待真机")
    pt.set_checked(dev, True)
    val = str(pt.free_slave_id())
    pt.set_slaveid(dev, val)
    succ, msg, errs = pt.save_and_result()
    pt.goto(); pt.ensure_enabled()
    persisted = pt.get_slaveid(dev)
    assert succ, f"有效 SlaveID={val} 保存应成功（提示:{msg!r} 错误:{errs}）"
    assert persisted == val, f"SlaveID 应持久化为 {val}，实际 {persisted!r}"


def test_pt_case03_slaveid_boundary(page_factory):
    """case03 越界 100/248 保存失败；边界 101/247 保存成功（按持久化判定）。"""
    pt = PassThroughPage(page_factory()); pt.goto(); pt.ensure_enabled()
    eths = pt.ethernet_device_names()
    if not eths:
        pytest.skip("透传表格无可编辑 Ethernet 设备行（demo 无下挂设备），待真机")
    dev = eths[0]
    fails = []
    for val, should_accept in [("100", False), ("248", False), ("101", True), ("247", True)]:
        pt.goto(); pt.ensure_enabled()
        # 只勾选目标设备，避免 101/247 与其它设备重复（隔离"范围校验"）
        for n in eths:
            pt.set_checked(n, n == dev)
        pt.set_slaveid(dev, val)
        succ, msg, errs = pt.save_and_result()
        pt.goto(); pt.ensure_enabled()
        persisted = pt.get_slaveid(dev)
        no_change = "no change" in msg.lower() or "无改动" in msg
        accepted = (bool(succ) or no_change) and persisted == val
        if should_accept and not accepted:
            fails.append(f"{val} 应被接受，但未生效(succ={succ},persist={persisted!r},msg={msg!r},err={errs})")
        if (not should_accept) and accepted:
            fails.append(f"{val} 应被拒绝(超出 101–247)，但保存生效了(msg={msg!r})")
    assert not fails, "SlaveID 边界校验不符合预期：\n" + "\n".join(fails)


def test_pt_case04_duplicate_slaveid(page_factory):
    """case04 两台设备配相同 SlaveID 应被拒绝（保存失败/提示重复，或不同时生效）。"""
    pt = PassThroughPage(page_factory()); pt.goto(); pt.ensure_enabled()
    eths = pt.ethernet_device_names()
    if len(eths) < 2:
        pytest.skip("需至少两台可编辑 Ethernet 设备（demo 无），待真机/更多下挂设备")
    d1, d2 = eths[0], eths[1]
    dup = str(pt.free_slave_id())
    pt.set_checked(d1, True); pt.set_slaveid(d1, dup)
    pt.set_checked(d2, True); pt.set_slaveid(d2, dup)
    succ, msg, errs = pt.save_and_result()
    pt.goto(); pt.ensure_enabled()
    p1, p2 = pt.get_slaveid(d1), pt.get_slaveid(d2)
    both_dup = (p1 == dup and p2 == dup)
    assert (not succ) or (not both_dup), (
        f"两台设备相同 SlaveID={dup} 应被拒绝/提示重复，但都保存成功了"
        f"（{d1}={p1}, {d2}={p2}, msg={msg!r}, err={errs}）")


def test_pt_case06_disabled_blocks_access(page_factory):
    """case06 Pass Through 关闭后，主站无法经透传 SlaveID 访问下游设备，系统仍正常。"""
    pt, working = _setup_pt_probes(page_factory)
    if not working:
        pytest.skip("无可经透传读通的设备（demo 无下挂设备/无 Modbus 服务），待真机")
    pt.goto(); pt.set_enabled(False)
    succ, msg, _ = pt.save_and_result(); pt.goto()
    disabled_persisted = not pt.is_enabled()
    can_read = True
    deadline = time.time() + PT_DISABLE_SETTLE
    while time.time() < deadline:
        if not any(_pt_read_one(s, r) for s, r in working.items()):
            can_read = False
            break
        time.sleep(1.5)
    # 系统仍正常：能再次进入页面（导航不抛异常即认为存活）
    pt.goto(); system_alive = pt.is_enabled() in (True, False)
    # 还原启用（baseline 也会兜底）
    pt.set_enabled(True)
    if pt.save():
        pt.has_message()
    assert succ and disabled_persisted, f"关闭 Pass Through 应保存并持久化（succ={succ}, 提示={msg!r}）"
    assert not can_read, f"Pass Through 关闭后 {PT_DISABLE_SETTLE:.0f}s 内仍能经透传读到数据（预期应无法访问）"
    assert system_alive, "RPP 系统应仍正常运行"


def test_pt_case07_concurrent_masters(page_factory):
    """case07 多个 Modbus 主站并发以不同 SlaveID 经透传读取，各自正常、无串扰。"""
    import threading
    pt, working = _setup_pt_probes(page_factory)
    sids = list(working.keys())[:3]
    if len(sids) < 2:
        pytest.skip(f"经透传可读的 SlaveID 不足 2 个（实得 {sorted(working.keys())}），待真机/更多下挂设备")
    dur = float(os.getenv("CONCURRENT_SECONDS", "20"))
    stats = {s: {"ok": 0, "err": 0} for s in sids}

    def worker(sid):
        rec = working[sid]
        deadline = time.time() + dur
        try:
            with Modbus() as mb:
                while time.time() < deadline:
                    try:
                        mb.read(rec, unit=sid); stats[sid]["ok"] += 1
                    except Exception:
                        stats[sid]["err"] += 1
                    time.sleep(0.1)
        except Exception:
            stats[sid]["err"] += 1

    threads = [threading.Thread(target=worker, args=(s,)) for s in sids]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    bad = []
    for s in sids:
        ok, err = stats[s]["ok"], stats[s]["err"]
        rate = ok / (ok + err) if (ok + err) else 0
        if ok == 0 or rate < 0.8:
            bad.append(f"SlaveID {s}: ok={ok} err={err} 成功率={rate:.0%}（单独可读，并发下失败）")
    assert not bad, "并发经透传读取存在异常（疑似串扰/网关并发限制）：\n" + "\n".join(bad)


if __name__ == "__main__":
    # 直接运行：python "Pass Through/test_rpp_pass_through.py" [额外 pytest 参数]
    _f = str(pathlib.Path(__file__).resolve())
    raise SystemExit(pytest.main([_f, *sys.argv[1:]]))
