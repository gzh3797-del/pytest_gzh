# -*- coding: utf-8 -*-
"""RPP Device Mirror 自动化（单文件、可独立执行的 pytest）——由 AcuHMI_1_7 同名套件适配。

验证 RPP「Device Mirror」镜像透传：对外用分配的 SlaveID 提供 Modbus TCP 访问。
与 1.7 相同的单文件结构：配置层 → Modbus 层 → 页面驱动层 → fixtures → 用例。

RPP 与 AcuHMI 1.7 的页面差异（2026-07-03 对 demo 实测确认，已在本文件适配）：
  ① 登录后落 /#/overview（1.7 是 dashboard）；默认密码弹窗点 Cancel。
  ② 路由守卫拦 hash 直跳 → goto() 走菜单：Settings → Protocols → Modbus 子菜单
     → Device Mirror（路由 /#/protocols/modbus/logicalParameterMapping）。
  ③ Save 按钮无改动时 disabled → save() 需先判断，无改动视为无需保存。
  ④ Disable 保存后整个表格 + Download All 按钮消失 → 读表格前 ensure_enabled()。
  ⑤ 本机行(SlaveID=1)勾选框 is-disabled、输入框 disabled、Device Name 可能为空
     → 按行锁定状态排除本机，不再按设备名。
  ⑥ demo 后台随机弹 "There's been an error." toast → 保存结果扫描全部消息。

demo 限制（RPP_DEMO=1，真机到位后置 0）：
  - 表格仅本机 1 行且锁定 → 行级编辑用例(case02/03/04)跳过
  - Download All 返回假文件（index.html）→ CSV 解析类断言跳过
  - 无 Modbus 服务 → 读数类用例跳过
  - B 路直读真实电表的来源页面（1.7 是 Physical Devices）待真机确认 → dm_003 二期

运行（仓库根目录下，任意人 clone 后均可）：
  pytest "projects/RPP/tests/Device Mirror" -v
也可直接运行本文件： python "projects/RPP/tests/Device Mirror/test_rpp_device_mirror.py" -v
"""
from __future__ import annotations

import os
import pathlib
import struct
import sys
import time
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
REL_TOL      = float(getattr(config, "TOLERANCE_PERCENT", 1.0)) / 100.0
ABS_FLOOR    = float(getattr(config, "TOLERANCE_ABSOLUTE", 0.05))
DISABLE_SETTLE = float(os.getenv("DISABLE_SETTLE", "30"))   # 禁用后等网关停止供数的最长秒数(轮询)
ENABLE_SETTLE  = float(os.getenv("ENABLE_SETTLE", "10"))    # 启用+保存后等镜像服务起来的最长秒数(轮询)
SETTLE_POLL    = float(os.getenv("SETTLE_POLL", "2.0"))
SLAVE_MIN, SLAVE_MAX = 2, 99          # 本机固定占 SlaveID=1
TIMEOUT      = int(os.getenv("TIMEOUT", "15000"))
INTERFACE_ETHERNET = "Ethernet"
REPORT = pathlib.Path(__file__).resolve().parent
DOWNLOAD_DIR = REPORT / "downloads"

# 用例函数名 → 关联 TC 编号（RPP 用例库编号确定后替换；供本地 conftest 并入 Test 列）
CASE_ID_MAP = {
    "test_dm_000_page_layout":              "RPP_DM_case00(页面布局)",
    "test_dm_001_config_and_export":        "RPP_DM_case06",
    "test_dm_002_modbus_read":              "RPP_DM_case05(子)",
    "test_dm_003_mirror_matches_direct":    "RPP_DM_case05",
    "test_dm_case01_enable_disable_toggle": "RPP_DM_case01",
    "test_dm_case02_valid_slaveid_save":    "RPP_DM_case02",
    "test_dm_case03_slaveid_boundary":      "RPP_DM_case03",
    "test_dm_case04_duplicate_slaveid":     "RPP_DM_case04",
    "test_dm_case07_offline_stability":     "RPP_DM_case07",
    "test_dm_case08_concurrent_masters":    "RPP_DM_case08",
}


# ═══════════════════════════════════════════════════════════════════════════
# Modbus TCP 读取 + 解码（pymodbus 3.x，device_id= 参数）—— 与 1.7 一致
# ═══════════════════════════════════════════════════════════════════════════
from pymodbus.client import ModbusTcpClient


def _regs_to_bytes(regs):
    return b"".join(struct.pack(">H", r & 0xFFFF) for r in regs)


def _decode(regs, kind):
    if kind == "uint16":
        return regs[0]
    if kind == "float32":
        return struct.unpack(">f", _regs_to_bytes(regs[:2]))[0]
    if kind == "float64":
        return struct.unpack(">d", _regs_to_bytes(regs[:4]))[0]
    if kind == "string":
        return _regs_to_bytes(regs).decode("ascii", errors="ignore").rstrip("\x00 ").strip()
    raise ValueError(f"未知解码类型: {kind}")


class Modbus:
    """Modbus TCP 客户端（context manager）。read 支持指定 unit；读取自动重连重试。"""

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
        slave = rec.slave_id if unit is None else unit
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
# MirrorConfALL.csv 解析（Download All 导出）—— 与 1.7 一致，另加有效性校验
# ═══════════════════════════════════════════════════════════════════════════
import csv as _csv

CSV_REQUIRED_COLS = {"Slave Id", "Device", "Parameter", "Type",
                     "Start Address(Dec.)", "End Address(Dec.)"}


@dataclass(frozen=True)
class ParamRecord:
    slave_id: int
    device: str
    parameter: str
    fc: int
    start: int
    end: int

    @property
    def width(self):
        return self.end - self.start + 1

    @property
    def kind(self):
        if self.fc in (1, 2):
            return "bit"
        w = self.width
        return {1: "uint16", 2: "float32", 4: "float64"}.get(w, "string")


def parse_mirror_csv(path):
    """解析导出的映射表；表头不含必需列时抛 ValueError（demo 返回的是假 HTML 文件）。"""
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        reader = _csv.DictReader(f)
        cols = set(reader.fieldnames or [])
        if not CSV_REQUIRED_COLS.issubset(cols):
            raise ValueError(f"非有效 MirrorConf CSV，表头={sorted(cols)[:8]}")
        out = []
        for row in reader:
            out.append(ParamRecord(
                slave_id=int(str(row["Slave Id"]).strip()),
                device=row["Device"].strip(),
                parameter=row["Parameter"].strip(),
                fc=int(str(row["Type"]).strip()),
                start=int(str(row["Start Address(Dec.)"]).strip()),
                end=int(str(row["End Address(Dec.)"]).strip()),
            ))
    return out


# ═══════════════════════════════════════════════════════════════════════════
# 登录 + Device Mirror 页面驱动（RPP 适配版，内联）
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


class MirrorPage:
    TAB_NAME = "Device Mirror"
    ENABLE_LABEL = "Device Mirror Enable"

    def __init__(self, page: Page):
        self.page = page

    def goto(self, retries=3):
        """经菜单进入 Device Mirror（RPP 路由守卫拦 hash 直跳，必须逐级点击）：
        顶部 Settings → 侧边 Protocols（落在 Modbus Config）→ 展开 Modbus 子菜单 → Device Mirror。"""
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
        """本机行判定：勾选框 is-disabled（RPP 本机固定 SlaveID=1、不可编辑）。"""
        return "is-disabled" in (row.locator(".el-checkbox").first.get_attribute("class") or "")

    def _radio(self, label):
        return self.page.locator(".el-radio", has_text=label).first

    def is_enabled(self):
        return "is-checked" in (self._radio("Enable").get_attribute("class") or "")

    def set_enabled(self, on):
        self._radio("Enable" if on else "Disable").click()

    def ensure_enabled(self):
        """确保处于 Enable 状态（RPP：Disable 下表格等控件整体消失）。"""
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

    def download_all(self):
        with self.page.expect_download(timeout=TIMEOUT) as dl:
            self.page.get_by_role("button", name="Download All").click()
        return dl.value

    # ── 行遍历 / 读写（跳过本机锁定行）─────────────────────────────────────
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
        """全部可编辑（非本机锁定）行的设备名。"""
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
        """Interface=Ethernet 且可编辑的设备名（排除本机锁定行）。"""
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

    def free_slave_id(self, lo=SLAVE_MIN, hi=SLAVE_MAX):
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

    def enable_all_distinct(self, slave_min=SLAVE_MIN):
        """勾选全部可编辑 Ethernet 设备 + 分配唯一 SlaveID + Save。demo 无可编辑行时为空操作。"""
        self.ensure_enabled()
        assigned, nid = {}, slave_min
        for name in self.ethernet_device_names():
            self.set_checked(name, True)
            self.set_slaveid(name, str(nid))
            assigned[name] = str(nid); nid += 1
        if self.save():
            self.has_message()
        return assigned

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
    """复用 pytest-playwright 的共享 browser；登录一次后把 sessionStorage 注入
    新 context 免重复登录（RPP 登录态在 sessionStorage，与 1.7 同模式，实测可用）。"""
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
    pg = page_factory()
    _ensure_modbus_config_enabled(pg)   # RPP 前置：Modbus 总开关必须启用
    dm = MirrorPage(pg); dm.goto()
    enabled0 = dm.is_enabled()
    dm.ensure_enabled()        # 启用才能看到表格做快照（RPP：Disable 下表格消失）
    snap = dm.snapshot()
    yield snap
    dm2 = MirrorPage(page_factory()); dm2.goto()
    dm2.ensure_enabled()
    drift = dm2.restore(snap) if snap else []
    # 还原 Enable/Disable 开关到原始状态
    dm2.goto()
    if dm2.is_enabled() != enabled0:
        dm2.set_enabled(enabled0)
        if dm2.save():
            dm2.has_message()
    if drift:
        import warnings
        warnings.warn(f"Device Mirror 还原后仍有偶发漂移：{drift}")


def _gateway_can_read(records):
    """经网关读这些记录，任一成功即返回 True；连不上/全错返回 False。"""
    if not records:
        return False
    try:
        with Modbus() as mb:
            vals = mb.read_all(records)
        return any(not isinstance(v, Exception) for v in vals.values())
    except Exception:
        return False


def _sample_records(records, per_device=1):
    """每台设备取若干条非字符串(测量)记录，作为读取探针。"""
    out, cnt = [], {}
    for r in records:
        if r.kind == "string":
            continue
        if cnt.get(r.device, 0) < per_device:
            out.append(r); cnt[r.device] = cnt.get(r.device, 0) + 1
    return out


def _wait_gateway_ready(records, timeout=ENABLE_SETTLE, poll=SETTLE_POLL):
    """启用+保存后轮询，直到网关镜像服务开始供数（任一探针读成功即就绪）。"""
    probes = _sample_records(records, per_device=2) or records[:10]
    if not probes:
        return False, 0.0
    deadline = time.time() + timeout
    last_rate = 0.0
    while time.time() < deadline:
        try:
            with Modbus() as mb:
                vals = mb.read_all(probes)
            ok = sum(1 for v in vals.values() if not isinstance(v, Exception))
            last_rate = ok / len(vals) if vals else 0.0
            if last_rate > 0:
                return True, last_rate
        except Exception:
            pass
        time.sleep(poll)
    return False, last_rate


@pytest.fixture(scope="module")
def exported_csv(page_factory):
    dm = MirrorPage(page_factory()); dm.goto()
    if not dm.is_enabled():
        dm.set_enabled(True)
        if dm.save():
            dm.has_message()
        dm.goto()
    assigned = dm.enable_all_distinct()   # demo 无可编辑行 → {}
    dm.goto(); dm.ensure_enabled()
    dl = dm.download_all()
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    target = DOWNLOAD_DIR / dl.suggested_filename
    dl.save_as(str(target))
    csv_valid, records, csv_err = True, [], ""
    try:
        records = parse_mirror_csv(str(target))
    except Exception as e:  # demo 假文件（index.html）走这里
        csv_valid, csv_err = False, f"{type(e).__name__}: {e}"
    ready = False
    if records and _gateway_reachable():
        ready, rate = _wait_gateway_ready(records)
        print(f"[exported_csv] 网关镜像就绪={ready} 探针成功率={rate:.0%}")
    return {"assigned": assigned, "records": records, "ready": ready,
            "csv_valid": csv_valid, "csv_err": csv_err, "file": target}


# ═══════════════════════════════════════════════════════════════════════════
# 用例
# ═══════════════════════════════════════════════════════════════════════════
def test_dm_000_page_layout(page_factory):
    """case00（新增）Device Mirror 页面布局：核心控件齐备、表头正确、本机行锁定、
    Disable 下表格隐藏 / Enable 恢复。demo 即可全量验证。"""
    dm = MirrorPage(page_factory()); dm.goto()
    fails = []

    # 1) Enable/Disable 单选 + Save + Download All
    for label in ("Enable", "Disable"):
        if not dm._radio(label).count():
            fails.append(f"缺少单选项 {label}")
    enabled0 = dm.is_enabled()
    dm.ensure_enabled()
    if not dm._save_button().count():
        fails.append("缺少 Save 按钮")
    if not dm.page.get_by_role("button", name="Download All").count():
        fails.append("Enable 状态下缺少 Download All 按钮")

    # 2) 表头
    headers = [t.strip() for t in dm.page.locator("th").all_inner_texts() if t.strip()]
    for h in ("SlaveID", "Device Name", "Interface", "Model"):
        if h not in headers:
            fails.append(f"表头缺少列 {h!r}（实际 {headers}）")

    # 3) 至少一行，且本机行锁定（勾选框 is-disabled + SlaveID 输入 disabled）
    rows = dm.rows
    if rows.count() < 1:
        fails.append("表格无数据行")
    else:
        locked = [i for i in range(rows.count()) if dm._row_locked(rows.nth(i))]
        if not locked:
            fails.append("未发现本机锁定行（预期本机 SlaveID=1 行不可编辑）")
        else:
            r0 = rows.nth(locked[0])
            if not dm._slaveid_input(r0).is_disabled():
                fails.append("本机行 SlaveID 输入框应为 disabled")
            if dm._slaveid_input(r0).input_value().strip() != "1":
                fails.append(f"本机行 SlaveID 应为 1，实际 {dm._slaveid_input(r0).input_value()!r}")

    # 4) 切 Disable → 表格与 Download All 隐藏；切回 Enable → 恢复（仅切单选，不保存）
    dm.set_enabled(False); dm.page.wait_for_timeout(600)
    if dm.rows.count() != 0:
        fails.append("Disable 状态下设备表格应隐藏")
    if dm.page.get_by_role("button", name="Download All").count() != 0:
        fails.append("Disable 状态下 Download All 应隐藏")
    dm.set_enabled(True); dm.page.wait_for_timeout(600)
    if dm.rows.count() < 1:
        fails.append("切回 Enable 后表格应恢复显示")
    dm.set_enabled(enabled0)   # 还原单选（未保存过，不影响持久化状态）

    assert not fails, "页面布局校验失败：\n" + "\n".join(fails)


def test_dm_001_config_and_export(exported_csv):
    """选中全部 Ethernet 设备 + 唯一 SlaveID + Download All，导出可解析且 SlaveID 唯一。"""
    assigned, records = exported_csv["assigned"], exported_csv["records"]
    assert exported_csv["file"].exists() and exported_csv["file"].stat().st_size > 0, \
        "Download All 未产生下载文件"
    if assigned:
        assert len(set(assigned.values())) == len(assigned), f"SlaveID 应互不相同：{assigned}"
    elif DEMO:
        pass   # demo 仅本机锁定行，无可编辑设备可勾选——下载动作已验证
    else:
        pytest.fail("应至少启用一个 Ethernet 设备")
    if not exported_csv["csv_valid"]:
        if DEMO:
            pytest.skip(f"demo 导出为假文件，CSV 解析跳过（{exported_csv['csv_err']}）；下载动作已验证")
        pytest.fail(f"导出 CSV 无法解析：{exported_csv['csv_err']}")
    assert records, "CSV 应解析出参数记录"


def test_dm_002_modbus_read(exported_csv):
    """按 CSV 经网关 Modbus 读取，成功率达标。"""
    if not exported_csv["records"]:
        pytest.skip("无有效 CSV 记录（demo 假文件），Modbus 读取跳过")
    if not _gateway_reachable():
        pytest.skip(f"网关 {GATEWAY_IP}:{MODBUS_PORT} Modbus 不可达（demo 无镜像服务）")
    with Modbus() as mb:
        results = mb.read_all(exported_csv["records"])
    ok = sum(1 for v in results.values() if not isinstance(v, Exception))
    total = len(results)
    rate = ok / total if total else 0
    assert rate >= 0.6, f"网关 Modbus 读取成功率过低 {rate:.0%}（{ok}/{total}）"


def test_dm_003_mirror_matches_direct():
    """镜像(A) ↔ 直读真实电表(B) 数据一致性。

    1.7 的 B 路 IP/Unit 来自 Physical Devices → Settings → Connection 页面；
    RPP 的对应物大概率是 Monitoring → Gateway Devices(/#/physicalDevices)，
    待真机接入下游设备后确认，再按 1.7 的 comparison fixture 移植。"""
    pytest.skip("RPP 真机未就绪：B 路（直读真实电表）连接信息来源页面待定，数据对比二期实现")


def test_dm_case01_enable_disable_toggle(page_factory, exported_csv):
    """case01 Disable+Save 持久化生效；Enable+Save 恢复。
    网关 Modbus 可达时附加验证：Disable 后停止供数、Enable 后恢复供数（demo 跳过该部分）。"""
    dm = MirrorPage(page_factory()); dm.goto()
    probe = _sample_records(exported_csv["records"]) or exported_csv["records"][:3]
    modbus_ok = bool(probe) and _gateway_reachable()

    # —— 切 Disable 并确认持久化 ——
    dm.ensure_enabled()
    dm.set_enabled(False)
    succ_dis, msg_dis, _ = dm.save_and_result()
    dm.goto()
    disabled_persisted = not dm.is_enabled()
    still_serving = False
    if modbus_ok:
        still_serving = True
        deadline = time.time() + DISABLE_SETTLE
        while time.time() < deadline:
            if not _gateway_can_read(probe):
                still_serving = False
                break
            time.sleep(2)

    # —— 切回 Enable 并确认持久化（+ 可读）——
    dm.goto(); dm.set_enabled(True)
    succ_en, msg_en, _ = dm.save_and_result()
    dm.goto()
    enabled_persisted = dm.is_enabled()
    can_read_enabled = None
    if modbus_ok:
        can_read_enabled = False
        for _ in range(6):
            if _gateway_can_read(probe):
                can_read_enabled = True
                break
            time.sleep(2)

    assert succ_dis and disabled_persisted, \
        f"切到 Disable 应保存并持久化（succ={succ_dis}, 仍enabled={not disabled_persisted}, 提示={msg_dis!r}）"
    assert succ_en and enabled_persisted, \
        f"切回 Enable 应保存并持久化（succ={succ_en}, 提示={msg_en!r}）"
    if modbus_ok:
        assert can_read_enabled, "Enable 后经 Device Mirror 应能读到数据"
        assert not still_serving, \
            f"Disable 后 {DISABLE_SETTLE:.0f}s 内经 Device Mirror 仍持续能读到数据（预期应停止供数）"
    else:
        import warnings
        warnings.warn("网关 Modbus 不可达（demo），本用例仅验证了 UI 开关持久化，未验证供数停止/恢复")


def _first_editable_device(dm):
    return next(iter(dm.ethernet_device_names()), None)


def test_dm_case02_valid_slaveid_save(page_factory):
    """case02 为设备配置有效 SlaveID（2–99）保存成功且持久化生效。"""
    dm = MirrorPage(page_factory()); dm.goto(); dm.ensure_enabled()
    dev = _first_editable_device(dm)
    if dev is None:
        pytest.skip("无可编辑 Ethernet 设备行（demo 仅本机锁定行），待真机/更多下挂设备")
    dm.set_checked(dev, True)
    val = str(dm.free_slave_id())
    dm.set_slaveid(dev, val)
    succ, msg, errs = dm.save_and_result()
    dm.goto(); dm.ensure_enabled()
    persisted = dm.get_slaveid(dev)
    assert succ, f"有效 SlaveID={val} 保存应成功（提示:{msg!r} 错误:{errs}）"
    assert persisted == val, f"SlaveID 应持久化为 {val}，实际 {persisted!r}"


def test_dm_case03_slaveid_boundary(page_factory):
    """case03 越界 1/100 保存失败；边界 2/99 保存成功（按持久化结果判定，兼容客户端钳制）。"""
    dm = MirrorPage(page_factory()); dm.goto(); dm.ensure_enabled()
    dev = _first_editable_device(dm)
    if dev is None:
        pytest.skip("无可编辑 Ethernet 设备行（demo 仅本机锁定行），待真机/更多下挂设备")
    fails = []
    for val, should_accept in [("1", False), ("100", False), ("2", True), ("99", True)]:
        dm.goto(); dm.ensure_enabled()
        dm.set_checked(dev, True)
        dm.set_slaveid(dev, val)
        succ, msg, errs = dm.save_and_result()
        dm.goto(); dm.ensure_enabled()
        persisted = dm.get_slaveid(dev)
        accepted = bool(succ) and persisted == val
        if should_accept and not accepted:
            fails.append(f"{val} 应被接受，但未生效(succ={succ},persist={persisted!r},msg={msg!r},err={errs})")
        if (not should_accept) and accepted:
            fails.append(f"{val} 应被拒绝(超出 2–99)，但保存生效了(msg={msg!r})")
    assert not fails, "SlaveID 边界校验不符合预期：\n" + "\n".join(fails)


def test_dm_case04_duplicate_slaveid(page_factory):
    """case04 两台设备配相同 SlaveID 应被拒绝（保存失败/提示重复，或不同时生效）。"""
    dm = MirrorPage(page_factory()); dm.goto(); dm.ensure_enabled()
    eths = dm.ethernet_device_names()
    if len(eths) < 2:
        pytest.skip("需至少两台可编辑 Ethernet 设备（demo 无），待真机/更多下挂设备")
    d1, d2 = eths[0], eths[1]
    dup = str(dm.free_slave_id())
    dm.set_checked(d1, True); dm.set_slaveid(d1, dup)
    dm.set_checked(d2, True); dm.set_slaveid(d2, dup)
    succ, msg, errs = dm.save_and_result()
    dm.goto(); dm.ensure_enabled()
    p1, p2 = dm.get_slaveid(d1), dm.get_slaveid(d2)
    both_dup = (p1 == dup and p2 == dup)
    assert (not succ) or (not both_dup), (
        f"两台设备相同 SlaveID={dup} 应被拒绝/提示重复，但都保存成功了"
        f"（{d1}={p1}, {d2}={p2}, msg={msg!r}, err={errs}）")


@pytest.mark.skip(reason="case07 需物理断开下游设备网络模拟离线，自动化无法控制设备上下线；请人工执行")
def test_dm_case07_offline_stability():
    """case07 下游设备离线后 Mirror 数据稳定、系统不崩溃、其他设备不受影响（人工）。"""


def test_dm_case08_concurrent_masters(exported_csv):
    """case08 多个 Modbus 主站并发读不同 SlaveID，各自正常、无串扰、系统稳定。"""
    if not exported_csv["records"]:
        pytest.skip("无有效 CSV 记录（demo 假文件），并发读取跳过")
    if not _gateway_reachable():
        pytest.skip(f"网关 {GATEWAY_IP}:{MODBUS_PORT} Modbus 不可达（demo 无镜像服务）")
    import threading
    recs = exported_csv["records"]
    by_sid_recs = {}
    for r in recs:
        if r.kind != "string":
            by_sid_recs.setdefault(r.slave_id, []).append(r)

    working = {}
    for sid, rs in by_sid_recs.items():
        for rec in rs[:12]:
            try:
                with Modbus() as mb:
                    mb.read(rec, unit=sid)
                working[sid] = rec
                break
            except Exception:
                continue
    sids = list(working.keys())[:3]
    if len(sids) < 2:
        pytest.skip(f"单独读取可成功的 SlaveID 不足 2 个（实得 {sorted(working.keys())}），无法做并发对比")

    dur = float(os.getenv("CONCURRENT_SECONDS", "20"))
    stats = {s: {"ok": 0, "err": 0} for s in sids}

    def worker(sid):
        rec = working[sid]
        deadline = time.time() + dur
        try:
            with Modbus() as mb:
                while time.time() < deadline:
                    try:
                        mb.read(rec, unit=sid)
                        stats[sid]["ok"] += 1
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
    assert not bad, "并发读取存在异常（疑似串扰/网关并发限制）：\n" + "\n".join(bad)


if __name__ == "__main__":
    # 直接运行：python "Device Mirror/test_rpp_device_mirror.py" [额外 pytest 参数]
    _f = str(pathlib.Path(__file__).resolve())
    raise SystemExit(pytest.main([_f, *sys.argv[1:]]))
