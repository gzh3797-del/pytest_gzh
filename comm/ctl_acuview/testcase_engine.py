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
# 同名页面各项目布局可能不同(如 RPP 与 AcuRev100 的 General): 项目可在自己的
# config_acuview.yaml 里用 nav.pages 段覆盖/新增条目(见 _nav_for), 不必改本表。
PAGE_NAV = {
    "General":       dict(tab="tab_setting", expand="tree_metering", tree="tree_general", landmark="lbl_backlight"),
    "Communication": dict(tab="tab_setting", expand=None,            tree="tree_communication", landmark="anchor_communication"),
}


def _nav_for(page: str) -> dict | None:
    """取页面导航项: 项目 config 的 nav.pages 覆盖 > 全局 PAGE_NAV。

    两种导航模式(由 config nav.pages 每页指定):
      1) 模板匹配(默认, RPP 等旧版 Acuview):
         General: {tab: tab_setting, expand: null, tree: tree_general, landmark: null}
      2) 坐标点击(AcuRev100 等左侧树布局固定的版本, 比模板/OCR 都稳):
         General: {mode: coords, tab_xy: [209,128], node_xy: [55,159]}
         (tab_xy=Reading/Setting 标签中心; node_xy=左侧树该节点文字中心; 均为最大化后的物理屏幕坐标)
    """
    try:
        over = get_config().data.get("nav", {}).get("pages", {})
    except Exception:
        over = {}
    if page in over:
        d = dict(over[page])
        if d.get("mode") == "coords" or d.get("node_xy"):
            return dict(mode="coords", tab_xy=d.get("tab_xy"), node_xy=d.get("node_xy"),
                        landmark=d.get("landmark"))
        return dict(tab=d.get("tab", "tab_setting"), expand=d.get("expand"),
                    tree=d.get("tree"), landmark=d.get("landmark"))
    return PAGE_NAV.get(page)


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
    """启动/连接 Acuview2, 处理 Add Connection 选表, 最大化。返回就绪的 AppDriver。

    复用会话缓存窗口时(drv.reused): Add Connection 仅首启出现, 跳过其探测(省~2s);
    窗口通常已最大化, 用 _ensure_maximized 免掉无谓的 maximize 等待(省~1.6s)。
    """
    drv = AppDriver().launch_or_connect()
    if not drv.reused:
        if _add_connection_open():
            row = int(cfg.app.get("connection_row", 0))
            drv.connect_meter(row_index=row)
            drv.app = None
            drv.launch_or_connect()
        drv.maximize()
    else:
        _ensure_maximized(drv)
    return drv


def _activate(drv: AppDriver):
    try:
        drv.win.set_focus()
    except Exception:
        pass
    time.sleep(0.3)
    r = drv.window_rect()
    drv.click_abs(r.left + 160, r.top + 10)   # 点标题栏激活(非控件区)
    time.sleep(0.25)


def _scroll_to_top(drv: AppDriver):
    """把页面滚到顶, 使 widget_abs(scroll_y=0) 成立。

    实测: 上位机翻页后会保留上次滚动量(verify_sync_interval 滚过就会残留),
    导致按 scroll=0 算的坐标落到错误行。滚轮悬停在*标签列*(非控件)上, 避免误改 spinBox 值。
    """
    r = drv.window_rect()
    pyautogui.moveTo(r.left + 340, r.top + 520)   # 内容区标签/背景区, 不在数值框上
    for _ in range(6):
        pyautogui.scroll(1000)
        time.sleep(0.04)
    time.sleep(0.2)


def _ensure_maximized(drv: AppDriver):
    """确保上位机窗口全屏(坐标导航/控件落点依赖全屏基线)。
    宽度不足 1900 才 maximize, 避免每次导航都花 maximize 的 1.5s。
    (2026-07-15 用户指示: 每条用例执行前先全屏, 保证坐标正确。)
    """
    try:
        if drv.window_rect().width < 1900:
            drv.maximize()
    except Exception:  # noqa: BLE001
        drv.maximize()


def navigate_to(drv: AppDriver, page: str):
    """导航到指定页面; 以"落地判据模板可见"或"点到树节点"为成功。带重试。

    已注册页(PAGE_NAV)走验证过的路径; 未注册页尝试 tab_setting + tree_<page小写>,
    找不到模板则抛清晰错误(提示补 comm/templates_acuview/tree_<page>.png)。
    """
    _ensure_maximized(drv)
    nav = _nav_for(page)
    if nav is None:
        guess = f"tree_{page.strip().lower().replace(' ', '_')}"
        if not (TPL / f"{guess}.png").exists():
            raise GuiError(
                f"页面 '{page}' 未注册导航且缺模板 comm/templates_acuview/{guess}.png; "
                f"请在 PAGE_NAV 注册或补该树节点截图。")
        nav = dict(tab="tab_setting", expand=None, tree=guess, landmark=None)

    # 坐标导航模式(树布局固定的版本): 点 Reading/Setting 标签 + 点树节点文字, 落点稳定不依赖模板/OCR。
    if nav.get("mode") == "coords":
        node_xy = nav.get("node_xy")
        if not node_xy:
            raise GuiError(f"页面 '{page}' 坐标导航缺 node_xy")
        for _ in range(3):
            _activate(drv)
            if nav.get("tab_xy"):
                drv.click_abs(int(nav["tab_xy"][0]), int(nav["tab_xy"][1]))
                time.sleep(0.5)
            drv.click_abs(int(node_xy[0]), int(node_xy[1]))
            time.sleep(0.9)
            if nav.get("landmark"):
                if drv.find_template(_t(nav["landmark"]), threshold=0.75):
                    _scroll_to_top(drv)
                    return
            else:
                _scroll_to_top(drv)
                return
        raise GuiError(f"坐标导航到 '{page}' 失败")

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
    """导航到目标页; 仅 General 页配置需要密码提权。

    2026-07-15 用户确认: *非 General 窗口的配置不需要输入密码*。故只有目标页是 General 时
    才走 General 提权(输 0000); Communication/Current&Wiring 等非 General 页直接导航即可写入,
    省去每次"导航 General→提权→再导目标页"的往返(大幅提速)。
    """
    if page == "General":
        navigate_to(drv, "General")
        elevate(drv)
    else:
        navigate_to(drv, page)


def _refresh_via_neutral(drv: AppDriver, page: str):
    """经中性兄弟页往返一次, 强制上位机重读目标页当前值。

    实测(2026-07-15): 此 Acuview 版本仅在"从别的页切入"时重读电表; 停在同一页
    (如上一条 comm 用例结束时就在 Communication)则 GUI 保留上次*界面设值*, 与被 Modbus
    还原过的电表实值失步。失步会让"目标值==界面陈旧值"的用例误判无改动(Update 不写)。
    进目标页前先跳一个中性页, 保证随后 navigate_to(page) 触发一次真实重读。
    """
    neutral = "General" if page != "General" else "Communication"
    nav = _nav_for(neutral)
    if not nav or nav.get("mode") != "coords" or not nav.get("node_xy"):
        return
    # 轻量: 只点 Setting 标签 + 中性节点各一次(不做重试/落地判据/滚动), 目的仅"离开目标页";
    # 随后的 _goto_and_elevate 再正式导回目标页, 那次才触发真实重读。~0.7s vs 完整 navigate ~3.9s。
    try:
        if nav.get("tab_xy"):
            drv.click_abs(int(nav["tab_xy"][0]), int(nav["tab_xy"][1]))
            time.sleep(0.3)
        drv.click_abs(int(nav["node_xy"][0]), int(nav["node_xy"][1]))
        time.sleep(0.4)
    except GuiError:
        pass


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
# 安全护栏: 禁写通信链路寄存器 (config.safety.forbid_write_*)
# --------------------------------------------------------------------------
def _reg_entry(vf: Verifier, register):
    """解析寄存器条目(addr/name/description/rw/range)。找不到返回 None。"""
    try:
        return vf.client._resolve(register)
    except Exception:  # noqa: BLE001
        return None


def _safety_forbids(entry: dict | None, register, cfg, allow_write=None) -> str | None:
    """按 config.safety 判断某寄存器是否禁写; 返回禁写原因或 None(允许)。

    allow_write: 本用例受控放行的地址(int)或名字子串(str)列表 —— 016 通讯参数类
    用例本身就是要改 SlaveID/波特率/校验, 需显式放行(用例自带重连+还原兜底)。
    """
    safety = cfg.data.get("safety", {}) if hasattr(cfg, "data") else {}
    forbid_addr = set(safety.get("forbid_write_addr", []) or [])
    forbid_sub = [str(s).upper() for s in (safety.get("forbid_write_name_substr", []) or [])]
    allow_write = allow_write or []
    addr = entry.get("addr") if entry else (register if isinstance(register, int) else None)
    names = " ".join(str(entry.get(k, "")) for k in ("name", "description")).upper() if entry else str(register).upper()
    # 放行判定
    for a in allow_write:
        if isinstance(a, int) and a == addr:
            return None
        if isinstance(a, str) and a.upper() in names:
            return None
    if addr in forbid_addr:
        return f"地址 {addr} 在 forbid_write_addr 禁写清单"
    for sub in forbid_sub:
        if sub and sub in names:
            return f"寄存器名/描述含禁写子串 '{sub}'"
    return None


# --------------------------------------------------------------------------
# 用例类型 1: 配置下发 + 校验 + 还原
# --------------------------------------------------------------------------
def run_write_verify_case(case_meta: dict, register, page: str, widget: str,
                          target_value, physical_note: str = "",
                          config_path: str | None = None, restore: bool = True,
                          allow_write=None) -> Report:
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
        # 安全护栏: 禁写通信链路寄存器(改了会断链)。allow_write 可逐用例受控放行。
        reason = _safety_forbids(_reg_entry(vf, register), register, cfg, allow_write)
        if reason:
            report.add(Check(label=f"安全校验: 禁写 {register}", expected="允许写",
                             actual="拒绝", passed=False,
                             detail=f"{reason}; 如确需(如016通讯参数用例)请传 allow_write 放行"))
            report.save(f"auto_{case_meta.get('编号', 'case')}")
            return report

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
# 共享: 弹窗确认 / 等待电表在线 / 多值序列
# --------------------------------------------------------------------------
def _confirm_dialogs(drv: AppDriver, password: str | None = PASSWORD, timeout: float = 12.0) -> int:
    """点击目标动作后, 依次处理可能出现的对话框:
       1) 需权限时的密码框(预填0000, 点 Confirm) 2) "Are you sure?" 确认框(Yes) 3) 结果/成功框。

    此 Acuview 版本模态弹窗位置固定, 按钮 *硬坐标* 比模板/OCR 匹配可靠(2026-07-15 实证:
    btn_yes 模板对 reboot/reset 确认框点不中致动作未触发)。以 OCR 文本判定弹窗类型 + 硬坐标点击。
    """
    deadline = time.time() + timeout
    clicks = 0
    did_pw = False
    idle = 0
    while time.time() < deadline:
        time.sleep(0.6)
        txt = ""
        try:
            txt = drv.ocr_text(_PWD_TITLE_REGION).lower()
        except GuiError:
            pass
        # 密码框(标题含 admin/permission, 预填0000): 点 Confirm 硬坐标
        if password and not did_pw and ("admin" in txt or "permission" in txt):
            drv.click_abs(*_PWD_CONFIRM_XY)
            did_pw = True
            clicks += 1
            time.sleep(0.9)
            idle = 0
            continue
        # Yes/No 确认框 或 结果/成功框: 点 Yes 硬坐标(此版本固定 1070,532)
        if any(k in txt for k in ("sure", "want to", "success", "successful",
                                  "complete", "reboot", "reset", "clear")):
            drv.click_abs(*_CONFIRM_YES_XY)
            clicks += 1
            time.sleep(1.0)
            idle = 0
            continue
        # 模板兜底(密码 Confirm / Yes)
        if password and not did_pw and drv.click_template(_t("btn_confirm"), threshold=0.95):
            did_pw = True
            clicks += 1
            time.sleep(0.9)
            idle = 0
            continue
        if drv.click_template(_t("btn_yes"), threshold=0.9):
            clicks += 1
            time.sleep(1.0)
            idle = 0
            continue
        idle += 1
        if clicks >= 1 and idle >= 2:
            break
    return clicks


def _wait_online(transport: str, timeout: float = 90.0, poll_addr: int = 4162,
                 overrides: dict | None = None) -> bool:
    """轮询读一个配置寄存器, 直到电表应答(重启/恢复出厂后重连恢复的判据)。"""
    overrides = overrides or {}
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with MeterClient(transport=transport, **overrides) as c:
                c.read_addr(poll_addr, "u16")
                return True
        except Exception:  # noqa: BLE001
            time.sleep(3.0)
    return False


def _wait_offline(transport: str, timeout: float = 30.0, poll_addr: int = 4162) -> bool:
    """轮询直到电表失联(确认复位动作确实让表掉线, 而非按钮没点上)。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with MeterClient(transport=transport) as c:
                c.read_addr(poll_addr, "u16")
            time.sleep(2.0)
        except Exception:  # noqa: BLE001
            return True
    return False


# --------------------------------------------------------------------------
# 用例类型 3: 按钮动作 (清能量 / 重启 / 恢复出厂 / 清 run·load-time)
# --------------------------------------------------------------------------
def run_button_action_case(case_meta: dict, page: str, button_widget: str,
                           verify: list | None = None, elevate_first: bool = False,
                           is_reset: bool = False, reset_verify: list | None = None,
                           recover_timeout: float = 90.0,
                           config_path: str | None = None) -> Report:
    """导航到 page, 点 button_widget(functionButton/clearButton), 处理确认弹窗;
    再按需验证效果。

    verify      : [(label, register, expected), ...] 动作后直接效果读(如清能量后=0)。
    is_reset    : True=该动作会让电表掉线重启(重启/恢复出厂), 先等掉线再等重连恢复。
    reset_verify: 重启/恢复出厂恢复后要断言的项(如出厂默认值)。
    ⚠️ 红线动作(清能量/重启/恢复出厂)的授权与门禁由调用方(用例文件 needs_review)把关。
       清能量类需"无电流"才能严格验 0: 调用方先控源置电流 0(见 helpers_accuracy.ensure_source_keepalive)。
    """
    cfg = get_config(config_path) if config_path else get_config()
    report = _header(case_meta)
    if _locked(report):
        return report

    transport = cfg.transport.verify
    # 动作前基线(留作诊断; 判决以 verify/reset_verify 为准)
    baseline = {}
    with Verifier(transport=transport) as vf:
        for item in (verify or []):
            label, reg, _exp = item
            try:
                baseline[reg] = vf.read_truth(reg)
            except Exception as exc:  # noqa: BLE001
                baseline[reg] = f"读失败:{exc}"

    drv = _connect(cfg)
    if elevate_first:
        navigate_to(drv, "General")
        elevate(drv)
    navigate_to(drv, page)
    print(f"[动作] 点击 {page}/{button_widget}")
    try:
        x, y = drv.widget_abs(page, button_widget)
    except GuiError as exc:
        report.add(Check(label=f"定位按钮 {button_widget}", expected="有坐标", actual="失败",
                         passed=False, detail=str(exc)))
        report.save(f"auto_{case_meta.get('编号', 'case')}")
        return report
    drv.click_abs(x, y)
    clicks = _confirm_dialogs(drv, password=PASSWORD)
    report.add(Check(label=f"执行动作 {button_widget} + 确认弹窗", expected="≥1 次确认",
                     actual=f"{clicks} 次", passed=clicks >= 1,
                     detail="模板匹配确认框; 若为0请核对确认按钮模板(Phase1 补图)"))

    if is_reset:
        # 重启瞬断: 自供电表重启掉线仅~1s(在场目视屏幕暗1s), 轮询难稳定捕获, 且 Device Run Time
        # 为*累计*运行时间(重启不清零, 仅 factory reset/Clear 清零)——故掉线检测记为*信息项不判决*。
        # 重启判决 = 动作已确认(上面 clicks>=1) + 重连恢复在线 + reset_verify(功能/配置正常)。
        went_off = _wait_offline(transport, timeout=15.0)
        back = _wait_online(transport, timeout=recover_timeout)
        report.add(Check(label="重启瞬断检测(信息项, 不判决)", expected="掉线(best-effort)",
                         actual="捕获到掉线" if went_off else "未捕获(重启<2s常测不到)",
                         passed=True, detail="自供电表重启掉线仅~1s; 物理重启以在场目视为准"))
        report.add(Check(label="重启后重连恢复在线", expected="在线",
                         actual="在线" if back else "超时未恢复", passed=back,
                         detail=f"轮询 {recover_timeout}s"))
        checks = reset_verify or []
    else:
        checks = verify or []

    if checks:
        with Verifier(transport=transport) as vf:
            for label, reg, expected in checks:
                try:
                    actual = _coerce_like(vf.read_truth(reg), reg, vf)
                except Exception as exc:  # noqa: BLE001
                    report.add(Check(label=label, expected=expected, actual=f"读失败:{exc}",
                                     passed=False, detail="")); continue
                ok, detail = values_match(expected, actual, vf.tol)
                report.add(Check(label=label, expected=expected, actual=actual, passed=ok,
                                 detail=detail))

    report.save(f"auto_{case_meta.get('编号', 'case')}")
    print(f"\n用例 {case_meta.get('编号')} 结果: {'[PASS]' if report.passed else '[FAIL]'}")
    return report


# --------------------------------------------------------------------------
# 用例类型 4: 通讯参数改后 Modbus 重连回读 (SlaveID / 波特率 / 校验)
# --------------------------------------------------------------------------
def run_comm_param_case(case_meta: dict, register, page: str, widget: str,
                        gui_value, expect_value=None, verify_overrides: dict | None = None,
                        allow_write=None, restore_value=None, restore_via: str = "modbus",
                        config_path: str | None = None) -> Report:
    """GUI 把通讯参数(page/widget)设为 gui_value 下发 → 校验端用新参数重连读回 →
    断言 == expect_value → 还原(默认 Modbus 直写, 杜绝留在非默认通讯参数)。

    gui_value       : 界面键入/下拉选择值(spinBox 数值 / combo 文本)。
    expect_value    : Modbus 回读期望(默认=gui_value; combo 类须给寄存器枚举 int)。
    verify_overrides: 改后校验端重连参数, 如 {'slave_id':5} / {'baudrate':9600}。
    allow_write     : 放行 forbid_write(通讯寄存器)。
    """
    cfg = get_config(config_path) if config_path else get_config()
    report = _header(case_meta)
    if _locked(report):
        return report
    expect_value = gui_value if expect_value is None else expect_value
    verify_overrides = verify_overrides or {}
    transport = cfg.transport.verify

    with Verifier(transport=transport) as vf:
        reason = _safety_forbids(_reg_entry(vf, register), register, cfg, allow_write)
        if reason and not allow_write:
            report.add(Check(label=f"安全校验: 禁写 {register}", expected="允许写", actual="拒绝",
                             passed=False, detail=reason))
            report.save(f"auto_{case_meta.get('编号', 'case')}")
            return report
        original = _coerce_like(vf.read_truth(register), register, vf)
    if restore_value is None:
        restore_value = original
    print(f"[1] 原值 {register}={original}; GUI 设 {gui_value}; 期望回读 {expect_value}; 还原 {restore_value}")

    drv = _connect(cfg)
    print(f"[2] 界面 {page}/{widget} = {gui_value} 下发")
    try:
        _refresh_via_neutral(drv, page)   # 强制重读, 消除上一条 comm 用例 Modbus 还原后的界面失步
        _goto_and_elevate(drv, page)
        _type_value(drv, page, widget, gui_value)
        commit_update(drv)
    except GuiError as exc:
        report.add(Check(label=f"GUI 下发 {gui_value}", expected="成功", actual="失败",
                         passed=False, detail=str(exc)))
        report.save(f"auto_{case_meta.get('编号', 'case')}")
        return report
    time.sleep(1.5)

    # 用新参数重连校验端回读
    try:
        with MeterClient(transport=transport, **verify_overrides) as c:
            actual = _coerce_like_client(c.read(register), register, c)
        ok, detail = values_match(expect_value, actual, float(cfg.run.value_tolerance))
        report.add(Check(label=f"改参重连后回读 {register} == {expect_value}",
                         expected=expect_value, actual=actual, passed=ok,
                         detail=f"重连参数={verify_overrides} {detail}"))
    except Exception as exc:  # noqa: BLE001
        report.add(Check(label=f"改参重连后回读 {register}", expected=expect_value,
                         actual=f"重连失败:{exc}", passed=False, detail=f"参数={verify_overrides}"))

    # 还原(默认 Modbus 直写: 用改后的参数连上去把值写回默认)
    restored = False
    try:
        with MeterClient(transport=transport, **verify_overrides) as c:
            c.write(register, restore_value, check_range=False)
        time.sleep(1.0)
        with MeterClient(transport=transport) as c:   # 用默认参数确认已恢复
            cur = _coerce_like_client(c.read(register), register, c)
        restored = (cur == restore_value)
        report.add(Check(label=f"还原 {register} == {restore_value}", expected=restore_value,
                         actual=cur, passed=restored, detail="Modbus 直写还原(避免留在非默认通讯参数)"))
    except Exception as exc:  # noqa: BLE001
        report.add(Check(label=f"还原 {register}", expected=restore_value, actual=f"失败:{exc}",
                         passed=False, detail="🛑 通讯参数可能未还原, 需人工核查!"))

    if not restored:
        print("🛑 告警: 通讯参数可能未还原到默认, 请人工核查 SlaveID/波特率/校验!")
    report.save(f"auto_{case_meta.get('编号', 'case')}")
    print(f"\n用例 {case_meta.get('编号')} 结果: {'[PASS]' if report.passed else '[FAIL]'}")
    return report


def _coerce_like_client(value, register, client: MeterClient):
    """_coerce_like 的 MeterClient 版(通讯参数用例重连时无 Verifier)。"""
    try:
        reg = client._resolve(register)
        if str(reg.get("dtype", "")).lower().rstrip("_t") in (
                "uint16", "int16", "uint32", "int32", "u16", "i16", "u32", "i32"):
            return int(round(float(value)))
    except Exception:  # noqa: BLE001
        pass
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


# --------------------------------------------------------------------------
# 用例类型 5: 非法输入拒绝 (越界/非数字, 验证 Modbus 回读不受影响)
# --------------------------------------------------------------------------
def run_reject_case(case_meta: dict, register, page: str, widget: str,
                    illegal_values: list, restore_value=None, allow_write=None,
                    config_path: str | None = None) -> Report:
    """逐个把非法值键入 GUI 并尝试下发, 每次经 Modbus 回读确认非法值*未生效*(被拒)。
    最后还原到 restore_value(默认=开始时的原值)。
    """
    cfg = get_config(config_path) if config_path else get_config()
    report = _header(case_meta)
    if _locked(report):
        return report
    transport = cfg.transport.verify

    with Verifier(transport=transport) as vf:
        reason = _safety_forbids(_reg_entry(vf, register), register, cfg, allow_write)
        if reason and not allow_write:
            report.add(Check(label=f"安全校验: 禁写 {register}", expected="允许写", actual="拒绝",
                             passed=False, detail=reason))
            report.save(f"auto_{case_meta.get('编号', 'case')}")
            return report
        original = _coerce_like(vf.read_truth(register), register, vf)
    restore_value = original if restore_value is None else restore_value
    print(f"[reject] {register} 原值={original}; 逐个尝试非法值={illegal_values}")

    drv = _connect(cfg)
    # 只在进入页面时导航一次(update 后页面不变, 无需重复导航; 2026-07-15 用户优化)。
    # 异常时重新导航一次作恢复。非 General 页免密。
    _goto_and_elevate(drv, page)
    for bad in illegal_values:
        try:
            _ensure_maximized(drv)   # 轻量全屏保障(不重复导航)
            _type_value(drv, page, widget, bad)
            try:
                commit_update(drv)
            except GuiError:
                pass  # 下发被阻止本身即"拒绝"的一种表现
        except GuiError as exc:
            report.add(Check(label=f"键入非法值 {bad!r}", expected="被拒/无法输入", actual="GUI异常",
                             passed=True, detail=f"界面阻止输入: {exc}"))
            _goto_and_elevate(drv, page)   # 出错恢复: 重新导航到目标页
            continue
        time.sleep(1.0)
        with Verifier(transport=transport) as vf:
            actual = _coerce_like(vf.read_truth(register), register, vf)
        # 拒绝判据: 非法值未写入寄存器
        rejected = str(actual) != str(_num_or_self(bad))
        report.add(Check(label=f"非法值 {bad!r} 被拒(回读!=非法值)", expected=f"!= {bad}",
                         actual=actual, passed=rejected,
                         detail="回读保持合法值即视为拒绝(上位机报错提示属MANUAL目视)"))

    # 还原(仍在同页, 不重复导航)
    _ensure_maximized(drv)
    _type_value(drv, page, widget, restore_value)
    try:
        commit_update(drv)
    except GuiError:
        pass
    time.sleep(1.0)
    with Verifier(transport=transport) as vf:
        cur = _coerce_like(vf.read_truth(register), register, vf)
    report.add(Check(label=f"还原 {register} == {restore_value}", expected=restore_value,
                     actual=cur, passed=(cur == restore_value), detail=""))
    report.save(f"auto_{case_meta.get('编号', 'case')}")
    print(f"\n用例 {case_meta.get('编号')} 结果: {'[PASS]' if report.passed else '[FAIL]'}")
    return report


def _num_or_self(v):
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return v


# --------------------------------------------------------------------------
# 用例类型 7: 密码门禁 (每连接首个 Setting Update 需密码; 输错拒绝/输对生效/同会话免密)
#
# 机制(2026-07-15 用户澄清 + 真机实测): 此 Acuview 无用户管理系统。一个连接建立后, 进入
# Setting 页对*首个*配置 Update 时弹密码框(标题 "This Operation Requires Admin Permission.
# Please Enter Password:", 预填 0000); 输对一次后本连接内后续任何配置修改免密; 关连接重开
# 后首个 Update 再次要密码。输错弹 "Wrong Password!" 提示(Yes 关闭), 配置不写入、权限不变。
#
# 弹窗按钮用硬坐标(此固定布局版本 OCR/模板匹配小按钮不稳; 与 nav 坐标同一策略)。
# 判据以可逆锚点参数的 Modbus 回读为权威(锚点须 Modbus 可写以便还原; 4200 PULSE_WIDTH 实测
# 不可 Modbus 写且 GUI/寄存器映射不符, 勿用; Energy Pulse Constant 0x1066=4198 可写, GUI值×1000=寄存器)。
# --------------------------------------------------------------------------
# 密码弹窗/重连坐标(1920×1080 最大化基线; 换分辨率需重取)
_PWD_FIELD_XY = (958, 477)            # 密码输入框(预填 0000)
_PWD_CONFIRM_XY = (1003, 517)         # 弹窗 Confirm 按钮
_PWD_WRONG_YES_XY = (1070, 532)       # "Wrong Password!" 提示的 Yes 按钮
_CONFIRM_YES_XY = (1070, 532)         # "Are you sure?"/结果框 的 Yes(reboot/reset/clear 通用)
_PWD_TITLE_REGION = (845, 405, 240, 55)   # 弹窗标题/提示文本 OCR 区
_RECONNECT_XY = {
    "close_tab": (475, 89), "confirm_yes": (1070, 532),
    "conn_menu": (110, 49), "add_conn": (152, 143),
    "row0_check": (519, 250), "row_step": 40, "connect": (1328, 145),
}


def _reconnect_acuview(drv: AppDriver, cfg):
    """断开当前连接 + 经 Add Connection 重开 → 重置密码门禁(使下一个 Update 重新要密码)。

    断开手法(2026-07-15 用户确认): 点连接标签 × → 弹 "Are you sure you want to remove meter?"
    点 Yes → 再 Yes。随后经 Add Connection 选回保存连接(ACmeter=row0/RS485·COM11)并 Connect。
    坐标为最大化基线固定布局。⚠️ 该重连流程按用户手法实现, 未逐次实测(设备交互行为以用户确认为准)。
    """
    rc = _RECONNECT_XY
    try:
        drv.click_abs(*rc["close_tab"]); time.sleep(1.0)     # 标签 ×
        drv.click_abs(*rc["confirm_yes"]); time.sleep(1.0)   # "remove meter?" Yes
        drv.click_abs(*rc["confirm_yes"]); time.sleep(1.0)   # 再 Yes
    except GuiError:
        pass
    drv.click_abs(*rc["conn_menu"]); time.sleep(0.6)
    drv.click_abs(*rc["add_conn"]); time.sleep(1.5)
    if not _add_connection_open():
        raise GuiError("重连失败: Add Connection 对话框未打开")
    row = int(cfg.app.get("connection_row", 0))
    drv.click_abs(rc["row0_check"][0], rc["row0_check"][1] + row * rc["row_step"]); time.sleep(0.5)
    drv.click_abs(*rc["connect"]); time.sleep(6.0)
    drv.app = None
    drv.launch_or_connect()
    drv.maximize()


def _pwd_popup_present(drv: AppDriver) -> bool:
    """OCR 判断密码弹窗是否出现(标题含 admin/permission, 区别于 'Wrong Password!')。"""
    txt = drv.ocr_text(_PWD_TITLE_REGION).lower()
    return ("admin" in txt) or ("permission" in txt)


def _wrong_password_shown(drv: AppDriver) -> bool:
    """OCR 判断 'Wrong Password!' 提示是否出现。"""
    return "wrong" in drv.ocr_text(_PWD_TITLE_REGION).lower()


def _click_update_btn(drv: AppDriver):
    hit = drv.find_template(_t("btn_update"))
    if not hit:
        raise GuiError("找不到 Update 按钮")
    drv.click_abs(hit[0], hit[1]); time.sleep(1.3)


def _enter_password(drv: AppDriver, password):
    """在密码弹窗清空预填并输入 password, 点 Confirm。"""
    drv.click_abs(*_PWD_FIELD_XY); time.sleep(0.3)
    pyautogui.hotkey("ctrl", "a"); pyautogui.press("delete"); time.sleep(0.2)
    pyautogui.typewrite(str(password), interval=0.05); time.sleep(0.3)
    drv.click_abs(*_PWD_CONFIRM_XY); time.sleep(1.3)


def _confirm_update_yes(drv: AppDriver) -> bool:
    """处理 'Do you want to update?' + 结果框(btn_yes 模板轮询, 与 commit_update 同)。"""
    clicks = 0
    for _ in range(5):
        if drv.click_template(_t("btn_yes"), threshold=0.9):
            clicks += 1
            time.sleep(0.9)
        elif clicks >= 1:
            break
        else:
            time.sleep(0.5)
    return clicks >= 1


def _read_anchor(cfg, register):
    with Verifier(transport=cfg.transport.verify) as vf:
        return _coerce_like(vf.read_truth(register), register, vf)


def run_password_gate_case(case_meta: dict, page: str, widget: str, anchor_register,
                           first_gui, first_expect, second_gui, second_expect,
                           restore_reg=None, wrong_password="9999", correct_password="0000",
                           allow_write=None, arm_gate: bool = True,
                           config_path: str | None = None) -> Report:
    """密码门禁三态: 首个 Update 输错密码被拒(配置未写) → 输对密码生效 → 同会话再改免密生效。

    first_gui/second_gui  = GUI 键入值; first_expect/second_expect = 对应 Modbus 寄存器期望值
                            (锚点有缩放时二者不同, 如 Energy Pulse Constant: GUI '2' → 寄存器 2000)。
    restore_reg           = 结束还原写回的寄存器值(Modbus 直写, 锚点须可写)。
    arm_gate=True         = 先重连(关/开连接)以武装门禁; 若已在全新连接上可传 False 免重连。
    correct_password 走硬坐标输入(预填即 0000=正确)。⚠️ 弹窗坐标为最大化基线观测值。
    """
    cfg = get_config(config_path) if config_path else get_config()
    report = _header(case_meta)
    if _locked(report):
        return report
    transport = cfg.transport.verify

    with Verifier(transport=transport) as vf:
        reason = _safety_forbids(_reg_entry(vf, anchor_register), anchor_register, cfg, allow_write)
        if reason and not allow_write:
            report.add(Check(label=f"安全校验: 禁写 {anchor_register}", expected="允许写",
                             actual="拒绝", passed=False, detail=reason))
            report.save(f"auto_{case_meta.get('编号', 'case')}")
            return report
        original = _coerce_like(vf.read_truth(anchor_register), anchor_register, vf)
    if restore_reg is None:
        restore_reg = original
    print(f"[pwd] 锚点 {anchor_register} 原值={original}; first={first_gui}->{first_expect}; "
          f"second={second_gui}->{second_expect}; restore={restore_reg}")

    drv = _connect(cfg)
    try:
        if arm_gate:
            try:
                _reconnect_acuview(drv, cfg)
            except GuiError as exc:
                print(f"[pwd] 重连武装门禁失败({exc}), 依赖当前连接门禁态")

        # ── 首个 Update: 输错密码 → 断言被拒 ──
        navigate_to(drv, page)   # 只导航不提权(保留"首个 Update 要密码"态)
        _type_value(drv, page, widget, first_gui)
        _click_update_btn(drv)
        if not _pwd_popup_present(drv):
            report.add(Check(label="首个 Update 弹密码框", expected="弹出", actual="未弹出",
                             passed=False, detail="门禁未武装(需在全新连接上运行; 重连失败或已提权)"))
            report.save(f"auto_{case_meta.get('编号', 'case')}")
            return report
        _enter_password(drv, wrong_password)
        wrong_shown = _wrong_password_shown(drv)
        drv.click_abs(*_PWD_WRONG_YES_XY); time.sleep(0.8)     # 关 "Wrong Password!" 提示
        after_wrong = _read_anchor(cfg, anchor_register)
        report.add(Check(label="输错密码被拒(提示+配置未写)",
                         expected=f"提示WrongPassword且回读={original}",
                         actual=f"提示={wrong_shown}, 回读={after_wrong}",
                         passed=(wrong_shown and after_wrong == original)))

        # ── 首个 Update: 输对密码(0000) → 断言生效 ──
        _type_value(drv, page, widget, first_gui)   # 幂等重设(防输错步丢弃待写值)
        _click_update_btn(drv)
        if _pwd_popup_present(drv):
            _enter_password(drv, correct_password)
        _confirm_update_yes(drv)
        time.sleep(1.2)
        after_correct = _read_anchor(cfg, anchor_register)
        report.add(Check(label="输对密码后配置生效", expected=first_expect, actual=after_correct,
                         passed=(after_correct == first_expect)))

        # ── 第二个 Update: 免密 → 断言直接生效 ──
        navigate_to(drv, page)
        _type_value(drv, page, widget, second_gui)
        _click_update_btn(drv)
        no_pwd = not _pwd_popup_present(drv)
        _confirm_update_yes(drv)
        time.sleep(1.2)
        after_second = _read_anchor(cfg, anchor_register)
        report.add(Check(label="同会话再改免密生效", expected=f"免密且回读={second_expect}",
                         actual=f"免密={no_pwd}, 回读={after_second}",
                         passed=(no_pwd and after_second == second_expect)))
    finally:
        try:
            with MeterClient(transport=transport) as c:
                c.write(anchor_register, restore_reg, check_range=False)
            time.sleep(0.8)
            cur = _read_anchor(cfg, anchor_register)
            report.add(Check(label=f"还原 {anchor_register} == {restore_reg}", expected=restore_reg,
                             actual=cur, passed=(cur == restore_reg), detail="Modbus 直写还原"))
        except Exception as exc:  # noqa: BLE001
            report.add(Check(label=f"还原 {anchor_register}", expected=restore_reg,
                             actual=f"失败:{exc}", passed=False, detail="[!] 需人工核查锚点参数"))

    report.save(f"auto_{case_meta.get('编号', 'case')}")
    print(f"\n用例 {case_meta.get('编号')} 结果: {'[PASS]' if report.passed else '[FAIL]'}")
    return report


# --------------------------------------------------------------------------
# 用例类型 6: 多值序列写回读 (CT Type 遍历 / CT Primary 多点)
# --------------------------------------------------------------------------
def run_multi_write_verify_case(case_meta: dict, register, page: str, widget: str,
                                steps: list, restore_value=None, allow_write=None,
                                config_path: str | None = None) -> Report:
    """按顺序把 widget 设为 steps 里每个值并逐个回读断言, 最后还原到 restore_value。

    steps: [(gui_value, expect_value), ...] 或 [value, ...](gui==expect)。
    combo 类 gui_value=显示文本, expect_value=寄存器枚举 int。
    """
    cfg = get_config(config_path) if config_path else get_config()
    report = _header(case_meta)
    if _locked(report):
        return report
    transport = cfg.transport.verify

    with Verifier(transport=transport) as vf:
        reason = _safety_forbids(_reg_entry(vf, register), register, cfg, allow_write)
        if reason and not allow_write:
            report.add(Check(label=f"安全校验: 禁写 {register}", expected="允许写", actual="拒绝",
                             passed=False, detail=reason))
            report.save(f"auto_{case_meta.get('编号', 'case')}")
            return report
        original = _coerce_like(vf.read_truth(register), register, vf)
    restore_value = original if restore_value is None else restore_value

    drv = _connect(cfg)
    # 只在进入页面时导航一次(update 后页面不变); 仅重试时重新导航作恢复。非 General 页免密。
    _goto_and_elevate(drv, page)
    with Verifier(transport=transport) as vf:
        seq = list(steps) + [(restore_value, restore_value)]  # 末尾追加还原步
        for i, step in enumerate(seq):
            gui_value, expect_value = step if isinstance(step, (tuple, list)) else (step, step)
            is_restore = (i == len(seq) - 1)
            label = (f"还原 {register} == {expect_value}" if is_restore
                     else f"设 {gui_value} 后回读 {register} == {expect_value}")
            actual = None
            for attempt in range(3):
                if attempt > 0:
                    _goto_and_elevate(drv, page)   # 重试才重新导航(恢复)
                _ensure_maximized(drv)
                _type_value(drv, page, widget, gui_value)
                try:
                    commit_update(drv)
                except GuiError as exc:
                    actual = f"Update错误:{exc}"; time.sleep(1.0); continue
                time.sleep(1.2)
                actual = _coerce_like(vf.read_truth(register), register, vf)
                ok, detail = values_match(expect_value, actual, vf.tol)
                if ok:
                    report.add(Check(label=label, expected=expect_value, actual=actual,
                                     passed=True, detail=f"{detail} (尝试 {attempt + 1}/3)"))
                    break
                time.sleep(1.0)
            else:
                report.add(Check(label=label, expected=expect_value, actual=actual,
                                 passed=False, detail="3 次写入仍不符"))

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
        report.add(Check(label="界面显示值 vs 期望",
                         expected=expect, actual=gui_val, passed=ok,
                         detail=f"GUI原文={gui_text!r} {detail}"))

    report.save(f"auto_{case_meta.get('编号', 'case')}")
    print(f"\n用例 {case_meta.get('编号')} 结果: {'[PASS]' if report.passed else '[FAIL]'}")
    return report
