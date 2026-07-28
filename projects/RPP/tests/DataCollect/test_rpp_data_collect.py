# -*- coding: utf-8 -*-
"""RPP Metering 数据采集与比对自动化（单文件、可独立执行的 pytest）——由 AcuHMI_1_7 DataCollect 适配。

1.7 流程（采集 Physical Devices 各设备 Metering 页面 → 匹配寄存器模板 → Modbus 直读比对）
在 RPP 上的对应物（2026-07-03 对 demo 实测确认）：
  - 采集对象不再是下挂设备，而是 RPP 自身面板：Monitoring → RPP Panel → Metering，
    视图 Realtime/Energy/Demand/Power Quality/Max Demand/Min/Max，
    页面有 VMM / Update Rate / Meter Point(channel) 三个下拉，按 VMM×channel 遍历采集。
  - 寄存器匹配与 Modbus 比对：待 RPP 寄存器表（blockParams 模板）与真机 Modbus 服务，
    到位后按 1.7 metering.py 的 match_registers/compare 移植（二期）。

demo 限制：Metering 页 VMM 下拉无选项、表格 "No Data" → 采集类用例自动跳过。

运行（仓库根目录下，任意人 clone 后均可）：
  pytest "projects/RPP/tests/DataCollect" -v
也可直接运行本文件： python "projects/RPP/tests/DataCollect/test_rpp_data_collect.py" -v
"""
from __future__ import annotations

import json
import os
import pathlib
import sys

import pytest
from playwright.sync_api import Browser, Page, TimeoutError as PWTimeout

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
    )


config = _load_local_config()

BASE_URL = os.getenv("RPP_URL", config.RPP_URL).rstrip("/")
USERNAME = os.getenv("RPP_USERNAME", config.RPP_USERNAME)
PASSWORD = os.getenv("RPP_PASSWORD", config.RPP_PASSWORD)
TIMEOUT  = int(os.getenv("TIMEOUT", "15000"))

REPORT = pathlib.Path(__file__).resolve().parent
JSON_OUT = REPORT / "metering_collect.json"

# RPP Panel → Metering 下的视图（2026-07-03 demo 左侧菜单实测）
VIEWS = ["Realtime", "Energy", "Demand", "Power Quality", "Max Demand", "Min/Max"]

CASE_ID_MAP = {
    "test_dc_000_page_layout":     "RPP_DC_case00(页面布局)",
    "test_dc_001_collect_json":    "RPP_DC_采集",
    "test_dc_002_register_match":  "RPP_DC_寄存器匹配",
    "test_dc_003_compare_modbus":  "RPP_DC_数据比对",
}


# ═══════════════════════════════════════════════════════════════════════════
# 登录（RPP 适配版，内联；与 Device Mirror / Pass Through 套件同款）
# ═══════════════════════════════════════════════════════════════════════════
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


# ═══════════════════════════════════════════════════════════════════════════
# RPP Panel Metering 页面驱动
# ═══════════════════════════════════════════════════════════════════════════
class PanelMetering:
    """Monitoring → RPP Panel → Metering 各视图；VMM / channel 下拉遍历 + 表格抓取。"""

    def __init__(self, page):
        self.page = page

    def goto(self):
        """进入 RPP Panel（默认落在 Metering/Realtime）。"""
        page = self.page
        page.goto(BASE_URL + "/#/overview", wait_until="domcontentloaded", timeout=TIMEOUT)
        page.get_by_text("RPP Panel", exact=True).first.click()
        page.wait_for_url("**/rppPanel/**", timeout=TIMEOUT)
        page.wait_for_timeout(1200)

    def _expand_metering(self):
        """确保左侧 Metering 子菜单展开（视图项可见）。"""
        item = self.page.locator(".el-menu-item", has_text="Realtime").first
        for _ in range(2):
            if item.count() and item.is_visible():
                return
            sub = self.page.locator(".el-sub-menu__title", has_text="Metering").first
            try:
                sub.click()
            except Exception:
                pass
            self.page.wait_for_timeout(500)

    def open_view(self, view):
        """打开 Metering 下的指定视图（注意 Logs 下也有 Realtime/Energy，取首个可见项）。"""
        self._expand_metering()
        loc = self.page.locator(".el-menu-item")
        for i in range(loc.count()):
            it = loc.nth(i)
            try:
                if it.inner_text().strip() == view and it.is_visible():
                    it.click(); self.page.wait_for_timeout(1200)
                    return True
            except Exception:
                continue
        return False

    def available_views(self):
        self._expand_metering()
        present, loc = [], self.page.locator(".el-menu-item")
        seen = set()
        for i in range(loc.count()):
            try:
                t = loc.nth(i).inner_text().strip()
                if t in VIEWS and t not in seen and loc.nth(i).is_visible():
                    present.append(t); seen.add(t)
            except Exception:
                continue
        return present

    # ── 下拉 ──────────────────────────────────────────────────────────────
    def _select_options(self, sel):
        """点开 el-select 收集可见选项文本，Esc 关闭。"""
        try:
            sel.click(); self.page.wait_for_timeout(600)
            opts = self.page.locator(".el-select-dropdown__item")
            items = []
            for i in range(opts.count()):
                o = opts.nth(i)
                try:
                    if o.is_visible():
                        t = o.inner_text().strip()
                        if t:
                            items.append(t)
                except Exception:
                    continue
            self.page.keyboard.press("Escape"); self.page.wait_for_timeout(300)
            return items
        except Exception:
            try:
                self.page.keyboard.press("Escape")
            except Exception:
                pass
            return []

    def _select_pick(self, sel, text):
        sel.click(); self.page.wait_for_timeout(500)
        self.page.locator(".el-select-dropdown__item", has_text=text).first.click()
        self.page.wait_for_timeout(1000)

    def vmm_select(self):
        return self.page.locator(".el-select").first

    def channel_select(self):
        sels = self.page.locator(".el-select")
        return sels.nth(sels.count() - 1) if sels.count() >= 2 else None

    def vmm_options(self):
        return self._select_options(self.vmm_select())

    def channel_options(self):
        ch = self.channel_select()
        return self._select_options(ch) if ch else []

    # ── 表格抓取（沿用 1.7 _read_tables_sep：逐表独立、仅表内去重）──────────
    def read_tables(self):
        out = []
        tables = self.page.locator(".el-table")
        for ti in range(tables.count()):
            t = tables.nth(ti)
            hdr = [h.strip() for h in t.locator("th").all_inner_texts()]
            if not hdr or len(hdr) < 2:
                continue
            cols = hdr[1:]
            rows, seen = [], set()
            body = t.locator(".el-table__row")
            for ri in range(body.count()):
                cells = [c.strip() for c in body.nth(ri).locator("td").all_inner_texts()]
                if not cells or not cells[0] or cells[0] in seen:
                    continue
                seen.add(cells[0])
                values = {cols[ci - 1]: cells[ci]
                          for ci in range(1, min(len(cells), len(cols) + 1))}
                rows.append({"parameter": cells[0], "values": values})
            if rows:
                out.append({"param_header": hdr[0], "columns": cols, "rows": rows})
        return out


def collect(page):
    """遍历 VMM × Metering 视图 ×（有 channel 下拉时逐 channel）抓取表格。
    返回 {vmm: {view: {channel|_default: [tables...]}}}；demo 无 VMM 选项时返回 {}。"""
    pm = PanelMetering(page)
    pm.goto()
    if not pm.open_view("Realtime"):
        return {}
    vmms = pm.vmm_options()
    if not vmms:
        return {}
    data = {}
    for vmm in vmms:
        pm._select_pick(pm.vmm_select(), vmm)
        vd = {}
        for view in pm.available_views():
            if not pm.open_view(view):
                continue
            chans = pm.channel_options()
            per = {}
            if chans:
                for ch in chans:
                    try:
                        pm._select_pick(pm.channel_select(), ch)
                        per[ch] = pm.read_tables()
                    except Exception:
                        continue
            else:
                per["_default"] = pm.read_tables()
            vd[view] = per
        data[vmm] = vd
    return data


# ═══════════════════════════════════════════════════════════════════════════
# Fixtures
# ═══════════════════════════════════════════════════════════════════════════
@pytest.fixture(scope="module")
def page_factory(browser: Browser):
    """复用 pytest-playwright 的共享 browser；登录一次后 sessionStorage 注入新 context 免重复登录。"""
    ctx0 = browser.new_context(ignore_https_errors=True)
    pg = ctx0.new_page(); pg.set_default_timeout(TIMEOUT)
    _login(pg)
    storage = pg.evaluate("JSON.stringify(window.sessionStorage)")
    ctx0.close()
    contexts = []

    def make():
        ctx = browser.new_context(ignore_https_errors=True)
        ctx.add_init_script(
            "(() => { const d = %s; for (const k in d){try{sessionStorage.setItem(k,d[k]);}catch(e){}} })();"
            % storage)
        page = ctx.new_page(); page.set_default_timeout(TIMEOUT)
        contexts.append(ctx)
        return page

    yield make
    for c in contexts:
        c.close()


@pytest.fixture(scope="module")
def collected(page_factory):
    data = collect(page_factory())
    JSON_OUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


# ═══════════════════════════════════════════════════════════════════════════
# 用例
# ═══════════════════════════════════════════════════════════════════════════
def test_dc_000_page_layout(page_factory):
    """case00（新增）RPP Panel Metering 页面布局：Metering 菜单含全部视图、
    Realtime 页有 VMM / Update Rate / Meter Point 下拉。demo 即可验证。"""
    pm = PanelMetering(page_factory()); pm.goto()
    fails = []

    views = pm.available_views()
    for v in VIEWS:
        if v not in views:
            fails.append(f"Metering 菜单缺少视图 {v!r}（实际 {views}）")

    if not pm.open_view("Realtime"):
        fails.append("无法打开 Realtime 视图")
    else:
        body = pm.page.locator("body").inner_text()
        for label in ("VMM", "Update Rate", "Meter Point"):
            if label not in body:
                fails.append(f"Realtime 页缺少 {label!r} 区块")
        if pm.page.locator(".el-select").count() < 2:
            fails.append("Realtime 页应有 VMM 与 channel 下拉")

    assert not fails, "页面布局校验失败：\n" + "\n".join(fails)


def test_dc_001_collect_json(collected):
    """Step1 采集：遍历 VMM × 视图 × channel 抓取表格，生成 metering_collect.json。"""
    assert JSON_OUT.exists(), f"采集结果文件未生成: {JSON_OUT}"
    if not collected:
        pytest.skip("demo Metering 页 VMM 下拉无选项（No Data），采集待真机/完整 demo 数据")
    total = sum(
        len(tbl["rows"])
        for vd in collected.values()
        for per in vd.values()
        for tables in per.values()
        for tbl in tables
    )
    assert total > 0, "采集到的参数条目为 0"
    print(f"\n  采集参数条目合计: {total}")


def test_dc_002_register_match(collected):
    """Step2 寄存器匹配：采集参数名 ↔ RPP 寄存器表。

    待 RPP 的 blockParams 寄存器模板（knowledge/shared/templates/raw/RPP*.xlsx 或
    随固件发布的地址表）到位后，按 1.7 metering.py 的 match_registers 移植。"""
    pytest.skip("RPP 寄存器模板未提供：寄存器匹配二期实现（参照 1.7 metering.py match_registers）")


def test_dc_003_compare_modbus(collected):
    """Step3 网页值 ↔ Modbus 直读比对（稳定量 FAIL=0，波动量仅报告）。

    待真机 Modbus 服务 + 寄存器模板到位后，按 1.7 metering.py 的 compare 移植
    （含稳定量/波动量分类、容差判定、数据对比结果.xlsx 报告）。"""
    pytest.skip("RPP 真机 Modbus 未就绪：数据比对二期实现（参照 1.7 metering.py compare）")


if __name__ == "__main__":
    # 直接运行：python "DataCollect/test_rpp_data_collect.py" [额外 pytest 参数]
    _f = str(pathlib.Path(__file__).resolve())
    raise SystemExit(pytest.main([_f, *sys.argv[1:]]))
