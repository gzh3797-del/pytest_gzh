"""通用自动化用例引擎 —— 把"手工用例"转成的自动化用例统一在这里跑。

按用例类型分派(可扩展):
  - run_write_verify_case : 配置下发 + 跨传输回读校验 + 还原(read→write→verify→restore)
  - run_read_compare_case : 界面数据获取(OCR) + 与 Modbus 真值/期望比对(只读, 不还原)

共享骨架: _connect(连接/选表/最大化) / navigate_to(导航) / elevate(提权) / commit_update(下发确认)。
全部复用 ctl_acuview 既有能力(gui_driver / modbus_client / verify), 不依赖 demos。

坐标基线: 1920×1080 / 125% / 窗口最大化; content_origin 由 config 标定(见 README)。
注意: TCP 连接 5min 空闲会断, 用例应连续跑。
"""
from __future__ import annotations

import re
import time
from pathlib import Path

import pyautogui

from .config import get_config
from .gui_driver import AppDriver, GuiError, is_session_locked
from .modbus_client import MeterClient
from .verify import Check, Report, Verifier, now, values_match

ROOT = Path(__file__).resolve().parent.parent
TPL = ROOT / "templates_acuview"
PASSWORD = "0000"


def _t(name: str) -> str:
    return str(TPL / f"{name}.png")


# 页面导航注册表: page -> (tab 模板, 父节点展开模板, 该页树节点模板, 落地判据模板)
# 只有验证过的页放进来; 其它页走通用猜测(tab_setting + tree_<page 小写>)。
PAGE_NAV = {
    "General":       dict(tab="tab_setting", expand="tree_metering", tree="tree_general", landmark="lbl_backlight"),
    "Communication": dict(tab="tab_setting", expand=None,            tree="tree_communication", landmark="anchor_communication"),
}


# --------------------------------------------------------------------------
# 共享: 连接 / 激活 / 导航 / 提权 / 下发
# --------------------------------------------------------------------------
def _add_connection_open() -> bool:
    from pywinauto import Application
    try:
        Application(backend="uia").connect(title_re="Add Connection", timeout=2)
        return True
    except Exception:
        return False


def _connect(cfg) -> AppDriver:
    """启动/连接 Acuview2, 处理 Add Connection 选表, 最大化。返回就绪的 AppDriver。"""
    drv = AppDriver().launch_or_connect()
    if _add_connection_open():
        row = int(cfg.app.get("connection_row", 0))
        drv.connect_meter(row_index=row)
        drv.app = None
        drv.launch_or_connect()
    drv.maximize()
    return drv


def _activate(drv: AppDriver):
    try:
        drv.win.set_focus()
    except Exception:
        pass
    time.sleep(0.5)
    r = drv.window_rect()
    drv.click_abs(r.left + 160, r.top + 10)   # 点标题栏激活(非控件区)
    time.sleep(0.4)


def _scroll_to_top(drv: AppDriver):
    """把页面滚到顶, 使 widget_abs(scroll_y=0) 成立。

    实测: 上位机翻页后会保留上次滚动量(verify_sync_interval 滚过就会残留),
    导致按 scroll=0 算的坐标落到错误行。滚轮悬停在*标签列*(非控件)上, 避免误改 spinBox 值。
    """
    r = drv.window_rect()
    pyautogui.moveTo(r.left + 340, r.top + 520)   # 内容区标签/背景区, 不在数值框上
    for _ in range(12):
        pyautogui.scroll(1000)
        time.sleep(0.04)
    time.sleep(0.3)


def navigate_to(drv: AppDriver, page: str):
    """导航到指定页面; 以"落地判据模板可见"或"点到树节点"为成功。带重试。

    已注册页(PAGE_NAV)走验证过的路径; 未注册页尝试 tab_setting + tree_<page小写>,
    找不到模板则抛清晰错误(提示补 comm/templates_acuview/tree_<page>.png)。
    """
    nav = PAGE_NAV.get(page)
    if nav is None:
        guess = f"tree_{page.strip().lower().replace(' ', '_')}"
        if not (TPL / f"{guess}.png").exists():
            raise GuiError(
                f"页面 '{page}' 未注册导航且缺模板 comm/templates_acuview/{guess}.png; "
                f"请在 PAGE_NAV 注册或补该树节点截图。")
        nav = dict(tab="tab_setting", expand=None, tree=guess, landmark=None)

    for _ in range(3):
        _activate(drv)
        for _ in range(3):
            if drv.click_template(_t(nav["tab"])):
                break
            time.sleep(0.5)
        time.sleep(1.0)
        if nav.get("expand") and not drv.find_template(_t(nav["tree"]), threshold=0.8):
            drv.click_template(_t(nav["expand"]))
            time.sleep(1.2)
        clicked = drv.click_template(_t(nav["tree"]), threshold=0.8)
        time.sleep(1.6)
        if nav.get("landmark"):
            if drv.find_template(_t(nav["landmark"]), threshold=0.75):
                _scroll_to_top(drv)
                return
        elif clicked:
            _scroll_to_top(drv)
            return
    raise GuiError(f"导航到 '{page}' 失败(未出现落地判据)")


def elevate(drv: AppDriver, password: str = PASSWORD):
    """在 General 页 Password 框输入密码并回车, 权限 View→Admin(全局生效)。

    实测: 必须 0000 + Enter 才能稳定提权, 否则 Update 写入失败。
    """
    x, y = drv.widget_abs("General", "Password_Value_Edit")
    drv.click_abs(x, y); time.sleep(0.3)
    drv.type_text(password, clear=True); time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(1.2)


def commit_update(drv: AppDriver):
    """点 Update 并依次确认对话框(密码/Do you want to update?/结果框)。"""
    hit = drv.find_template(_t("btn_update"))
    if not hit:
        raise GuiError("找不到 Update 按钮")
    if not drv.update_and_confirm((hit[0], hit[1])):
        raise GuiError("Update 确认对话框未完成")


def _goto_and_elevate(drv: AppDriver, page: str):
    """到 General 提权(View→Admin)再导到目标页。

    实测: Update 后权限会回落, 故 *每次写入(含还原)前都要重新提权*, 否则下发不生效。
    """
    navigate_to(drv, "General")
    elevate(drv)
    if page != "General":
        navigate_to(drv, page)


# 文本类控件直接键入(带 settle 间隔, 复刻 demo 已验证的 _type_into;
# gui_driver.set_value 无间隔, 点击后立刻按键会因焦点未就绪而丢值 -> 下发旧值)。
_TEXT_TYPES = ("spinBox", "lineEdit", "ipEdit", "doubleSpinBox", None)


def _type_value(drv: AppDriver, page: str, widget: str, value):
    w = drv._widget(page, widget)
    wt = w.get("type")
    x, y = drv.widget_abs(page, widget)
    if wt in _TEXT_TYPES:
        drv.click_abs(x, y); time.sleep(0.3)
        pyautogui.hotkey("ctrl", "a"); pyautogui.press("delete"); time.sleep(0.2)
        pyautogui.typewrite(str(value), interval=0.05); time.sleep(0.3)
    else:
        # comboBox / switchButton 走 gui_driver 的类型分派(下拉需 OCR)
        drv.set_value(page, widget, value)


# --------------------------------------------------------------------------
# 工具
# --------------------------------------------------------------------------
def _coerce_like(value, register, vf: Verifier):
    """按寄存器 dtype 把读到的原值规整(整型寄存器去掉 .0)。"""
    try:
        reg = vf.client._resolve(register)
        if str(reg.get("dtype", "")).lower().rstrip("_t") in ("uint16", "int16", "uint32", "int32", "u16", "i16", "u32", "i32"):
            return int(round(float(value)))
    except Exception:
        pass
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _parse_num(text: str):
    m = re.search(r"-?\d+(?:\.\d+)?", str(text).replace(",", ""))
    return float(m.group()) if m else None


def _header(case_meta: dict) -> Report:
    cid = case_meta.get("编号", "?")
    report = Report(title=f"{cid} | {case_meta.get('标题', '')}", started=now())
    print("=" * 78)
    print(f"自动化用例 {cid}")
    for k in ("标题", "预置条件", "测试步骤", "预期结果"):
        if case_meta.get(k):
            print(f"  {k}: {str(case_meta[k])[:100]}")
    print("=" * 78)
    return report


def _locked(report: Report) -> bool:
    if is_session_locked():
        report.add(Check(label="前置: 会话未锁屏(界面可截图/点击)", expected="unlocked",
                         actual="locked", passed=False,
                         detail="锁屏/远程断开时无法驱动界面, 请解锁活动桌面后重试"))
        return True
    return False


def _write_and_check(drv: AppDriver, vf: Verifier, report: Report, label: str,
                     page: str, widget: str, value, register, attempts: int = 3) -> bool:
    """提权→设值→下发→跨传输回读, 失败自动重试(应对 TCP 空闲断连/Write to Device failed 的瞬态)。

    每次都重新 _goto_and_elevate(权限 Update 后会回落)。回读匹配即成功; attempts 次仍不符记 FAIL。
    """
    actual = None
    for i in range(attempts):
        _goto_and_elevate(drv, page)
        _type_value(drv, page, widget, value)
        try:
            commit_update(drv)
        except GuiError as exc:
            actual = f"Update错误: {exc}"
            time.sleep(1.2)
            continue
        time.sleep(1.2)
        actual = _coerce_like(vf.read_truth(register), register, vf)
        ok, detail = values_match(value, actual, vf.tol)
        if ok:
            report.add(Check(label=label, expected=value, actual=actual, passed=True,
                             detail=f"{detail} (尝试 {i + 1}/{attempts})"))
            return True
        print(f"    写入未生效(尝试 {i + 1}/{attempts}): 回读={actual}, 重试...")
        time.sleep(1.0)
    report.add(Check(label=label, expected=value, actual=actual, passed=False,
                     detail=f"{attempts} 次写入仍不符(可能 TCP 空闲断连 / Write to Device failed, 见 README)"))
    return False


# --------------------------------------------------------------------------
# 用例类型 1: 配置下发 + 校验 + 还原
# --------------------------------------------------------------------------
def run_write_verify_case(case_meta: dict, register, page: str, widget: str,
                          target_value, physical_note: str = "",
                          config_path: str | None = None, restore: bool = True) -> Report:
    """界面把 (page/widget) 设为 target_value 并下发, 跨传输回读断言, 再还原原值。

    case_meta : {编号,标题,预置条件,测试步骤,预期结果}
    register  : 校验/还原用的寄存器(名或十进制地址)
    """
    cfg = get_config(config_path) if config_path else get_config()
    report = _header(case_meta)
    if _locked(report):
        return report

    original = None
    need_fallback_to = None
    with Verifier(transport=cfg.transport.verify) as vf:
        original = _coerce_like(vf.read_truth(register), register, vf)
        print(f"[1] 读原值 {register} = {original}; 计划界面写入 = {target_value}")

        drv = _connect(cfg)
        print(f"[2] 界面写入 {page}/{widget} = {target_value} 并下发(失败自动重试)")
        ok = _write_and_check(drv, vf, report, f"界面下发后回读 {register} == {target_value}",
                              page, widget, target_value, register)
        if physical_note:
            report.add(Check(label="物理/目视确认项", expected=physical_note, actual="(软件不可读)",
                             passed=ok, detail="MANUAL: 无对应寄存器, 以设置值写入成功为前提, 请按手工用例目视核对"))

        if restore:
            print(f"[3] 还原 {register} = {original}")
            _write_and_check(drv, vf, report, f"还原后回读 == 原值 {original}",
                             page, widget, original, register)
            cur = _coerce_like(vf.read_truth(register), register, vf)
            if cur != original:
                need_fallback_to = original

    if need_fallback_to is not None:
        try:
            with MeterClient(transport="rtu") as c:
                c.write(register, need_fallback_to)
            report.add(Check(label="RTU 兜底还原", expected=need_fallback_to,
                             actual=need_fallback_to, passed=True, detail="best-effort"))
        except Exception as exc:  # noqa: BLE001
            report.add(Check(label="RTU 兜底还原", expected=need_fallback_to,
                             actual="失败", passed=False, detail=str(exc)))

    report.save(f"auto_{case_meta.get('编号', 'case')}")
    print(f"\n用例 {case_meta.get('编号')} 结果: {'[PASS]' if report.passed else '[FAIL]'}")
    return report


# --------------------------------------------------------------------------
# 用例类型 2: 界面数据获取 + 比对(只读)
# --------------------------------------------------------------------------
def run_read_compare_case(case_meta: dict, page: str, widget: str,
                          expect=None, register=None, tolerance: float | None = None,
                          config_path: str | None = None) -> Report:
    """导航到读数页, OCR 取界面显示值, 与 Modbus 真值(register)或期望常量(expect)比对。

    需 Tesseract OCR; 未装则该用例记 FAIL 并在 detail 注明(供 Skill 转 skip / 回填调试结果)。
    """
    cfg = get_config(config_path) if config_path else get_config()
    report = _header(case_meta)
    if _locked(report):
        return report
    tol = tolerance if tolerance is not None else float(cfg.run.value_tolerance)

    drv = _connect(cfg)
    try:
        navigate_to(drv, page)
    except GuiError as exc:
        report.add(Check(label=f"导航到 {page}", expected="可达", actual="失败", passed=False, detail=str(exc)))
        report.save(f"auto_{case_meta.get('编号', 'case')}")
        return report

    try:
        gui_text = drv.get_value(page, widget)
    except GuiError as exc:
        report.add(Check(label=f"OCR 读取 {page}/{widget}", expected="数值", actual="无法读取",
                         passed=False, detail=f"{exc} (read_compare 依赖 Tesseract OCR, 见 README)"))
        report.save(f"auto_{case_meta.get('编号', 'case')}")
        return report

    gui_val = _parse_num(gui_text)
    if register is not None:
        with Verifier(transport=cfg.transport.verify) as vf:
            truth = vf.read_truth(register)
        ok, detail = values_match(truth, gui_val, tol)
        report.add(Check(label=f"界面显示值 vs Modbus真值({register})",
                         expected=truth, actual=gui_val, passed=ok,
                         detail=f"GUI原文={gui_text!r} {detail}"))
    else:
        ok, detail = values_match(expect, gui_val, tol)
        report.add(Check(label=f"界面显示值 vs 期望",
                         expected=expect, actual=gui_val, passed=ok,
                         detail=f"GUI原文={gui_text!r} {detail}"))

    report.save(f"auto_{case_meta.get('编号', 'case')}")
    print(f"\n用例 {case_meta.get('编号')} 结果: {'[PASS]' if report.passed else '[FAIL]'}")
    return report
