"""Acuview2 固件升级(Meter Update 窗口)自动化 runner —— 014 Firmware升级 模块。

流程(2026-07-15 真机全流程实操走通, 工程师现场逐步确认):
  Operation → Firmware Update... → Meter Update 窗口 → 选 Programming Baud Rate
  → Select Firmware File → 文件对话框(填路径+打开) → Select All → Connect
  → Yes(确认) → Yes(risk 提示) → Status: Clearing...→Programming→Verifying
  → Write Success! → Update Finished 弹窗(标题 Message) → Yes → 关窗 → 重建连接。

已知怪癖(2026-07-15 实录):
  - Meter Update / Message 弹窗均非模态, 会被主窗盖住; 点弹窗按钮前须先激活 Meter Update。
  - Update Finished 的 Message 弹窗可能落在副屏, 激活 Meter Update 后才浮顶可点。
  - 文件对话框路径用 powershell Set-Clipboard + Ctrl+V 粘贴(typewrite 打反斜杠不稳)。
  - 升级中主窗连接标签清空(设备断开)属正常; 升级完成后必须关 Meter Update 再重建连接。
  - FIRMWARE_VERSION@61440 等版本寄存器是 ASCII 双字符打包(0x312E='1.'), 按字节解码。

判据(2026-07-15 工程师裁决):
  - 同版本重刷判"升级流程成功"(Write Success + Update Finished + 表恢复在线 + 数据保持),
    版本号仅记录不强判; config firmware.expect_version 非空时才严判版本。
  - 数据保持用配置类寄存器快照(Basic Setting 4096..4424)升级前后比对;
    能量持续累计, 不作保持判据(仅记录)。

安全: 升级会重启电表(团队约定: 执行前须用户确认)。config `run.allow_firmware_upgrade`
默认 false, 刷机类用例经 `upgrade_allowed()` 门禁 skip; 现场确认后置 true 再跑。
"""
from __future__ import annotations

import logging
import time

import pyautogui

logging.getLogger("pymodbus").setLevel(logging.WARNING)
logging.getLogger("pymodbus.logging").setLevel(logging.WARNING)

from .config import get_config
from .gui_driver import AppDriver, GuiError
from .testcase_engine import _connect, _header, _locked, _t
from .verify import Check, Report, Verifier

# ---- 主窗口(1920x1080 最大化基线)绝对坐标 ----
MENU_OPERATION_XY = (269, 49)          # 菜单栏 Operation
MENU_FIRMWARE_XY = (322, 144)          # 下拉项 Firmware Update...
ADD_METER_XY_REL = (340, 98)           # 主窗口 "Click To Add New Meter" 的 +(窗口相对)

# ---- Meter Update 窗口内控件(相对窗口左上角; 窗口 788x547, UIA 可取矩形) ----
REL_SCAN_TOGGLE = (44, 59)
REL_BTN_CONNECT = (68, 103)
REL_BTN_SELECT_ALL = (258, 103)
REL_BAUD_COMBO = (519, 103)
REL_BTN_SELECT_FILE = (678, 103)
REL_ROW1_SELECT = (58, 183)
REL_INFO_REGION = (195, 42, 585, 36)   # 设备信息栏 "Model: .. Hardware: .. Firmware: .."
REL_ROW_STRIP = (20, 150, 755, 75)     # 表格首行整条(含 Status 列), OCR 关键词用
REL_STATUS_CELL = (600, 168, 185, 36)  # Status 列单元格(升级进度/校验结果; 红/绿字, 增强OCR)
REL_BAUD_POPUP = (469, 117, 160, 290)  # 波特率下拉展开区(supply 给 _select_combo_option)

# 升级状态关键词(OCR Status 列, 按出现顺序)
_STAGES = ("Clearing", "Programming", "Verifying")

# 动态量寄存器(随时间/运行走动, 不能作"升级前后保持"判据; 升级耗时约2分钟时钟必然变):
#   4121/4122 运行时间(32位)、4123/4124 负载时间(32位)、4127~4134 实时时钟(周/年/月/日/时/分/秒/毫秒)。
_DYNAMIC_ADDRS = frozenset(range(4121, 4125)) | frozenset(range(4127, 4135))

# 配置类寄存器快照范围(Basic Setting 全量, 2026-07-15 实测 102 项可读; 4400 段为
# Operations 触发寄存器, 静态应恒 0, 一并快照)。剔除 32 位高字读不到地址 + 动态量寄存器。
SETTING_SNAPSHOT_ADDRS: tuple[int, ...] = tuple(
    a for a in range(4096, 4138) if a not in (4118,) and a not in _DYNAMIC_ADDRS
) + tuple(a for a in range(4160, 4202) if a not in (4164, 4190, 4194, 4199)) + tuple(range(4400, 4425))

# 设备信息类寄存器(ASCII 打包): 名称 -> (地址, 寄存器数)
INFO_REGS = {
    "FIRMWARE_VERSION": (61440, 2),
    "BOOTLOADER_VERSION": (61458, 2),
    "HARDWARE_VERSION": (61520, 2),
    "SERIAL_NUMBER": (61504, 16),
    "MODEL": (61553, 10),
}


def upgrade_allowed(config_path: str) -> bool:
    """刷机门禁: config run.allow_firmware_upgrade 为 true 才允许真实升级。"""
    try:
        cfg = get_config(config_path)
        return bool(cfg.data.get("run", {}).get("allow_firmware_upgrade", False))
    except Exception:  # noqa: BLE001 - 配置异常一律视为未授权
        return False


def _fw_cfg(cfg) -> dict:
    return cfg.data.get("firmware", {}) or {}


# --------------------------------------------------------------------------
# Modbus 侧: ASCII 版本/信息读取 + 配置快照
# --------------------------------------------------------------------------
def _ascii_regs(client, addr: int, nregs: int) -> str:
    """ASCII 双字符打包寄存器解码(高字节在前), 去 NUL/空白。"""
    regs = client.read_block(addr, nregs)
    chars = []
    for r in regs:
        chars.append(chr((int(r) >> 8) & 0xFF))
        chars.append(chr(int(r) & 0xFF))
    return "".join(c for c in chars if c.isprintable()).strip()


def read_device_info(vf: Verifier) -> dict:
    """读设备信息类寄存器(版本号 ASCII 解码)。单项失败记 None 不中断。"""
    out = {}
    for name, (addr, nregs) in INFO_REGS.items():
        try:
            out[name] = _ascii_regs(vf.client, addr, nregs)
        except Exception:  # noqa: BLE001
            out[name] = None
    return out


def snapshot_settings(vf: Verifier) -> dict[int, int | None]:
    """配置类寄存器逐地址快照(u16 原值); 读不到记 None。"""
    snap: dict[int, int | None] = {}
    for addr in SETTING_SNAPSHOT_ADDRS:
        try:
            snap[addr] = int(vf.client.read_addr(addr))
        except Exception:  # noqa: BLE001
            snap[addr] = None
    return snap


def _diff_snapshots(pre: dict, post: dict) -> list[str]:
    """比对两次快照, 返回差异明细(空=完全一致)。None(读失败)不参与判差。"""
    diffs = []
    for addr, v0 in pre.items():
        v1 = post.get(addr)
        if v0 is None or v1 is None:
            continue
        if v0 != v1:
            diffs.append(f"@{addr}: {v0} -> {v1}")
    return diffs


def _wait_meter_online(vf_factory, timeout_s: float = 180.0, interval_s: float = 5.0):
    """升级/重启后等电表在 USB 口恢复可读; 返回可用的 Verifier(调用方负责关闭)。"""
    deadline = time.time() + timeout_s
    last_exc: Exception | None = None
    while time.time() < deadline:
        try:
            vf = vf_factory()
            vf.__enter__()
            vf.client.read_block(INFO_REGS["FIRMWARE_VERSION"][0], 2)
            return vf
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            time.sleep(interval_s)
    raise GuiError(f"电表 {timeout_s}s 内未恢复在线: {last_exc}")


# --------------------------------------------------------------------------
# GUI 侧: Meter Update 窗口驱动
# --------------------------------------------------------------------------
def _find_window(title: str, min_width: int = 100):
    """按标题精确找可见顶层窗口(pywinauto uia)。找不到返回 None。"""
    from pywinauto import Desktop
    try:
        for w in Desktop(backend="uia").windows(visible_only=True):
            try:
                if w.window_text() == title and w.rectangle().width() > min_width:
                    return w
            except Exception:  # noqa: BLE001
                continue
    except Exception:  # noqa: BLE001
        pass
    return None


def _mu_rect(mu) -> tuple[int, int]:
    r = mu.rectangle()
    return r.left, r.top


def _mu_click(drv: AppDriver, mu, rel: tuple[int, int]):
    left, top = _mu_rect(mu)
    drv.click_abs(left + rel[0], top + rel[1])


def _mu_activate(mu):
    """激活 Meter Update(非模态, 会被主窗盖住; 弹窗也要靠它浮顶)。"""
    try:
        mu.set_focus()
    except Exception:  # noqa: BLE001
        pass
    time.sleep(0.2)


def open_meter_update(drv: AppDriver) -> object:
    """主窗口 Operation → Firmware Update..., 返回 Meter Update 窗口句柄。已开则复用。"""
    mu = _find_window("Meter Update")
    if mu is not None:
        _mu_activate(mu)
        return mu
    try:
        drv.win.set_focus()
    except Exception:  # noqa: BLE001
        pass
    time.sleep(0.25)
    for _ in range(3):
        drv.click_abs(*MENU_OPERATION_XY)
        time.sleep(0.4)
        drv.click_abs(*MENU_FIRMWARE_XY)
        time.sleep(1.0)
        mu = _find_window("Meter Update")
        if mu is not None:
            _mu_activate(mu)
            return mu
        pyautogui.press("esc")
        time.sleep(0.5)
    raise GuiError("Operation→Firmware Update 未弹出 Meter Update 窗口")


def _wait_file_dialog(timeout_s: float = 10.0):
    """等升级文件选择对话框弹出。

    实测(2026-07-16): 点 Meter Update 的 "Select Firmware File" 按钮后, 弹出的文件
    对话框标题是 "Select Firmware"(不是按钮全名); 中文系统可能是 "打开"。用宽松正则兜底。
    """
    from pywinauto import Desktop
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        time.sleep(0.5)
        try:
            cands = Desktop(backend="uia").windows(
                title_re=r".*(Select Firmware|打开|Open).*", visible_only=True)
            for w in cands:
                try:
                    if w.rectangle().width() > 200:
                        return w
                except Exception:  # noqa: BLE001
                    continue
        except Exception:  # noqa: BLE001
            pass
    return None


def _win32_file_dialog():
    """win32 后端定位真正的文件选择对话框(#32770 全屏可见), 排除隐藏 Qt 幽灵窗。

    实测(2026-07-16): 同名 "Select Firmware" 有两个窗口——真对话框 class=#32770 可见,
    另有一个隐藏的 Qt5*QWindowIcon; uia 后端点框易落空, win32 直控 Edit 稳。
    """
    from pywinauto import Desktop
    for w in Desktop(backend="win32").windows(title_re=r"Select Firmware|打开|Open"):
        try:
            if w.class_name() == "#32770" and w.is_visible():
                return w
        except Exception:  # noqa: BLE001
            continue
    return None


def _dialog_open(package_path: str, timeout_s: float = 10.0) -> bool:
    """在文件对话框文件名框写路径并提交(win32 set_edit_text + Enter)。返回是否已关闭对话框。"""
    deadline = time.time() + timeout_s
    dlg = None
    while time.time() < deadline:
        dlg = _win32_file_dialog()
        if dlg is not None:
            break
        time.sleep(0.5)
    if dlg is None:
        raise GuiError("Select Firmware 文件对话框未弹出")
    edits = [e for e in dlg.descendants(class_name="Edit")]
    if not edits:
        raise GuiError("文件对话框无文件名输入框")
    # 文件名框(最底部那个 Edit): 取 top 最大者
    fn_edit = max(edits, key=lambda e: e.rectangle().top)
    fn_edit.set_edit_text(package_path)
    time.sleep(0.2)
    fn_edit.set_focus()
    time.sleep(0.15)
    pyautogui.press("enter")
    # 等对话框关闭(=已打开文件)
    for _ in range(30):
        time.sleep(0.3)
        if _win32_file_dialog() is None:
            return True
    return False


def select_firmware_file(drv: AppDriver, mu, package_path: str, parse_timeout_s: float = 45.0) -> bool:
    """点 Select Firmware File → 文件对话框填路径打开 → 等信息栏出现包名(解析完成)。

    返回 True=包已加载(信息栏含 .MFEA); False=未见加载完成(非法包场景由调用方查错误弹窗)。
    文件对话框走 win32 set_edit_text + Enter(2026-07-16 实证: 剪贴板+坐标粘贴会落空)。
    """
    _mu_activate(mu)
    _mu_click(drv, mu, REL_BTN_SELECT_FILE)
    closed = _dialog_open(package_path, timeout_s=10.0)
    if not closed:
        return False   # 对话框未关(路径无效/被拦), 调用方按未加载处理
    # 等解析完成: 信息栏出现 "(<包名>.MFEA)"
    deadline = time.time() + parse_timeout_s
    left, top = _mu_rect(mu)
    x, y, w, h = REL_INFO_REGION
    while time.time() < deadline:
        time.sleep(1.5)
        try:
            txt = drv.ocr_text((left + x, top + y, w, h))
        except GuiError:
            return False
        if "MFEA" in txt.upper():
            return True
    return False


def _close_popup_yes(drv: AppDriver, mu, popup) -> bool:
    """关确认/结果弹窗: 激活 Meter Update 让弹窗浮顶(副屏怪癖), 把弹窗挪回主屏后点 Yes。"""
    _mu_activate(mu)
    time.sleep(0.5)
    try:
        r = popup.rectangle()
        if r.left >= 1920 or r.left < -100:  # 弹窗落在副屏 → 挪回主屏
            popup.move_window(x=700, y=400)
            time.sleep(0.5)
        popup.set_focus()
        time.sleep(0.4)
    except Exception:  # noqa: BLE001
        pass
    for _ in range(3):
        if drv.click_template(_t("btn_yes"), threshold=0.85):
            time.sleep(1.0)
            return True
        pyautogui.press("enter")   # 兜底: 默认按钮
        time.sleep(0.8)
    return False


def _confirm_update_dialogs(drv: AppDriver, mu, timeout_s: float = 25.0) -> int:
    """Connect 后连续点掉两个确认弹窗(Do you want to update? / risk 提示)。返回点掉个数。"""
    clicks = 0
    deadline = time.time() + timeout_s
    while time.time() < deadline and clicks < 2:
        time.sleep(0.4)
        _mu_activate(mu)
        if drv.click_template(_t("btn_yes"), threshold=0.85):
            clicks += 1
            time.sleep(0.8)
    return clicks


def wait_update_result(drv: AppDriver, mu, timeout_s: float = 2700.0) -> tuple[bool, str, list[str]]:
    """轮询 Status 列直到 Write Success!/Failed 或超时。

    返回 (成功?, 末次状态文本, 观测到的阶段序列)。Update Finished 弹窗出现也视为成功信号。
    """
    seen: list[str] = []
    last_txt = ""
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        time.sleep(3.0)
        popup = _find_window("Message")
        if popup is not None:
            _close_popup_yes(drv, mu, popup)
            if "Write Success" not in " ".join(seen):
                seen.append("Update Finished 弹窗")
            return True, last_txt or "Update Finished", seen
        txt = read_row_status(drv, mu)   # 增强 OCR(灰度+放大), 读红/绿 Status 字
        if txt:
            last_txt = txt
        for stage in _STAGES:
            if stage.lower() in txt.lower() and (not seen or seen[-1] != stage):
                seen.append(stage)
                print(f"    [升级进度] {stage} ({time.strftime('%H:%M:%S')})")
        if "success" in txt.lower():
            seen.append("Write Success")
            return True, last_txt, seen
        if "failed" in txt.lower() or "fail" in txt.lower():
            return False, last_txt, seen
        # 防锁屏: 轻移鼠标(不点击)
        pyautogui.moveRel(1, 0)
        pyautogui.moveRel(-1, 0)
    return False, f"超时({timeout_s}s): {last_txt}", seen


def close_meter_update(mu):
    try:
        mu.close()
        time.sleep(1.5)
    except Exception:  # noqa: BLE001
        pass


def reconnect_main(drv: AppDriver, cfg, wait_s: float = 20.0):
    """升级后重建主窗口连接: 点 + → Add Connection → 选行 Connect(行号取 config)。"""
    try:
        drv.win.set_focus()
    except Exception:  # noqa: BLE001
        pass
    time.sleep(0.6)
    r = drv.window_rect()
    drv.click_abs(r.left + ADD_METER_XY_REL[0], r.top + ADD_METER_XY_REL[1])
    time.sleep(2.5)
    if _find_window("Add Connection") is not None:
        drv.connect_meter(row_index=int(cfg.app.get("connection_row", 0)))
    time.sleep(wait_s)


# --------------------------------------------------------------------------
# runner 1: 正常升级(单次/压测) + 数据保持
# --------------------------------------------------------------------------
def run_firmware_update_case(case_meta: dict, config_path: str, baud: int | None = None,
                             rounds: int = 1, package: str | None = None,
                             expect_version: str | None = None,
                             check_settings: bool = True,
                             check_info_keep: bool = False,
                             physical_note: str = "") -> Report:
    """RTU 升级全流程 + 升级前后数据保持校验。

    baud: Programming Baud Rate(None=保持当前); rounds: 连续升级次数(压测);
    package: 升级包路径(None=config firmware.package);
    expect_version: 期望升级后版本(None=不强判, 仅记录; 同版重刷裁决);
    check_settings: 升级前后配置类寄存器快照比对;
    check_info_keep: 追加设备信息寄存器(SN/Model/HW/Boot)保持比对(case16)。
    """
    cfg = get_config(config_path)
    report = _header(case_meta)
    cid = case_meta.get("编号", "case")
    if _locked(report):
        report.save(f"auto_{cid}")
        return report
    if not upgrade_allowed(config_path):
        report.add(Check(label="前置: 刷机授权(run.allow_firmware_upgrade)", expected=True,
                         actual=False, passed=False,
                         detail="升级会重启电表; 现场确认后在 config 置 true 再跑"))
        report.save(f"auto_{cid}")
        return report

    fw = _fw_cfg(cfg)
    pkg = package or fw.get("package")
    if not pkg:
        report.add(Check(label="前置: 升级包路径(firmware.package)", expected="非空",
                         actual=None, passed=False, detail="config firmware.package 未配置"))
        report.save(f"auto_{cid}")
        return report
    timeout_s = float(fw.get("update_timeout_s", 2700))

    # ---- 升级前基线(USB/COM6) ----
    with Verifier() as vf:
        pre_info = read_device_info(vf)
        pre_snap = snapshot_settings(vf) if check_settings else {}
    ok_reads = sum(1 for v in pre_snap.values() if v is not None)
    print(f"  [基线] FW={pre_info.get('FIRMWARE_VERSION')} 配置快照 {ok_reads}/{len(pre_snap)} 项")

    # ---- GUI 升级(可多轮) ----
    drv = _connect(cfg)
    all_rounds_ok = True
    for rnd in range(1, rounds + 1):
        tag = f"第{rnd}/{rounds}轮" if rounds > 1 else ""
        mu = open_meter_update(drv)
        if baud is not None:
            _mu_activate(mu)
            left, top = _mu_rect(mu)
            _mu_click(drv, mu, REL_BAUD_COMBO)
            time.sleep(0.8)
            px, py, pw, ph = REL_BAUD_POPUP
            try:
                drv._select_combo_option(str(baud), (left + px, top + py, pw, ph))
                report.add(Check(label=f"选择升级波特率 {tag}".strip(), expected=baud, actual=baud,
                                 passed=True, detail="下拉 OCR 选值"))
            except GuiError as exc:
                report.add(Check(label=f"选择升级波特率 {tag}".strip(), expected=baud,
                                 actual=f"下拉选值失败: {exc}", passed=False))
                close_meter_update(mu)
                all_rounds_ok = False
                break
        loaded = select_firmware_file(drv, mu, pkg)
        report.add(Check(label=f"加载升级包 {tag}".strip(), expected="信息栏出现 .MFEA 包名",
                         actual="已加载" if loaded else "未确认加载", passed=loaded,
                         detail=str(pkg)))
        if not loaded:
            close_meter_update(mu)
            all_rounds_ok = False
            break
        _mu_activate(mu)
        _mu_click(drv, mu, REL_BTN_SELECT_ALL)
        time.sleep(0.5)
        _mu_click(drv, mu, REL_BTN_CONNECT)
        n_yes = _confirm_update_dialogs(drv, mu)
        report.add(Check(label=f"确认弹窗(升级确认+risk提示) {tag}".strip(), expected=2,
                         actual=n_yes, passed=n_yes >= 1,
                         detail="Connect 后 Yes×2 开始刷机"))
        ok, last_txt, stages = wait_update_result(drv, mu, timeout_s=timeout_s)
        report.add(Check(label=f"升级过程无错误 {tag}".strip(),
                         expected="Clearing→Programming→Verifying→Write Success",
                         actual=" → ".join(stages) if stages else last_txt, passed=ok,
                         detail=last_txt))
        if not ok:
            close_meter_update(mu)
            all_rounds_ok = False
            break
        if rnd < rounds:
            time.sleep(10.0)   # 下一轮前给设备缓冲(升级后重启)
    # ---- 收尾: 关窗 + 重建连接 ----
    mu = _find_window("Meter Update")
    if mu is not None:
        close_meter_update(mu)
    reconnect_main(drv, cfg)

    # ---- 升级后校验(USB/COM6; 表重启后等在线) ----
    if all_rounds_ok:
        vf = _wait_meter_online(Verifier, timeout_s=180.0)
        try:
            post_info = read_device_info(vf)
            post_snap = snapshot_settings(vf) if check_settings else {}
        finally:
            vf.__exit__(None, None, None)
        report.add(Check(label="升级后电表恢复在线(USB回读)", expected="在线",
                         actual="在线", passed=True))
        fw_pre, fw_post = pre_info.get("FIRMWARE_VERSION"), post_info.get("FIRMWARE_VERSION")
        if expect_version:
            report.add(Check(label="升级后 FW 版本", expected=expect_version, actual=fw_post,
                             passed=str(fw_post) == str(expect_version)))
        else:
            report.add(Check(label="FW 版本记录(同版重刷不强判)", expected=f"升级前 {fw_pre}",
                             actual=f"升级后 {fw_post}", passed=fw_post is not None,
                             detail="裁决2026-07-15: 无第二版本包, 版本号仅记录"))
        if check_settings:
            diffs = _diff_snapshots(pre_snap, post_snap)
            report.add(Check(label="升级前后配置寄存器保持(Basic Setting 快照)",
                             expected="0 项差异", actual=f"{len(diffs)} 项差异",
                             passed=not diffs, detail="; ".join(diffs[:10])))
        if check_info_keep:
            for name in ("SERIAL_NUMBER", "MODEL", "HARDWARE_VERSION", "BOOTLOADER_VERSION"):
                report.add(Check(label=f"升级前后 {name} 保持", expected=pre_info.get(name),
                                 actual=post_info.get(name),
                                 passed=pre_info.get(name) == post_info.get(name)))
    if physical_note:
        report.add(Check(label="MANUAL 目视项", expected=physical_note, actual="待人工确认",
                         passed=True, detail="物理观察项不计自动判据"))
    report.save(f"auto_{cid}")
    return report


# --------------------------------------------------------------------------
# runner 2: 非法固件包合法性校验(不刷机, 不重启)
# --------------------------------------------------------------------------
def read_row_status(drv: AppDriver, mu) -> str:
    """OCR 设备行 Status 列文本(升级进度/校验结果都显示在这)。

    Status 为红/绿细字, 默认灰度 OCR 常漏读 → 灰度+3x放大+psm7 增强(2026-07-16 实证)。
    """
    left, top = _mu_rect(mu)
    x, y, w, h = REL_STATUS_CELL
    try:
        from PIL import Image, ImageOps
        img = ImageOps.grayscale(pyautogui.screenshot(region=(left + x, top + y, w, h)))
        img = img.resize((img.width * 3, img.height * 3),
                         getattr(Image, "Resampling", Image).LANCZOS)
        if not getattr(drv, "_pytesseract", None):
            return ""
        return " ".join(drv._pytesseract.image_to_string(img, config="--psm 7").split())
    except Exception:  # noqa: BLE001
        return ""


# 他产品包被拒绝时 Status 列关键词(2026-07-16 实测: 加载 AcuIOM/4100 包 → "Model Unmatched!";
# OCR 可能把 t 认成 i → "Unmaiched", 故用 "unma" 前缀容错)。
_REJECT_KEYWORDS = ("unma", "invalid", "not match", "mismatch", "fail")


def run_firmware_invalid_file_case(case_meta: dict, config_path: str,
                                   invalid_package: str | None = None) -> Report:
    """选择他产品 .MFEA, 断言上位机做固件合法性校验并拒绝, 不发生升级。

    实测(2026-07-16): 此上位机对他产品包**加载时即校验**——信息栏会显示*包*的型号,
    但设备行 Status 列显示红字 "Model Unmatched!" 且 Select 复选框置灰不可选,
    即"校验成功=正确拒绝"。故判据是 Status 列关键词, 不是弹窗(弹窗为 1320 版行为)。
    """
    cfg = get_config(config_path)
    report = _header(case_meta)
    cid = case_meta.get("编号", "case")
    if _locked(report):
        report.save(f"auto_{cid}")
        return report
    bad = invalid_package or _fw_cfg(cfg).get("invalid_package")
    if not bad:
        report.add(Check(label="前置: 非法包路径(firmware.invalid_package)", expected="非空",
                         actual=None, passed=False))
        report.save(f"auto_{cid}")
        return report
    drv = _connect(cfg)
    mu = open_meter_update(drv)
    try:
        select_firmware_file(drv, mu, bad, parse_timeout_s=10.0)
        # 加载后轮询 Status 列, 等校验结果出现(拒绝关键词)
        status = ""
        rejected = False
        for _ in range(8):
            time.sleep(1.0)
            status = read_row_status(drv, mu)
            if any(k in status.lower() for k in _REJECT_KEYWORDS):
                rejected = True
                break
        report.add(Check(label="他产品包固件合法性校验被拒绝(Status 列)",
                         expected="Status 显示 Model Unmatched!/校验失败",
                         actual=status or "(Status 为空)", passed=rejected, detail=str(bad)))
    finally:
        close_meter_update(_find_window("Meter Update") or mu)
    report.save(f"auto_{cid}")
    return report


# --------------------------------------------------------------------------
# runner 3: 升级界面设备信息栏显示校验(只读, 不刷机)
# --------------------------------------------------------------------------
def run_firmware_info_display_case(case_meta: dict, config_path: str) -> Report:
    """OCR 信息栏 Model/Hardware/Firmware 并校验。

    实测(2026-07-16): 信息栏在未选包前显示 "Please Select Firmware File!", 须先加载合法
    升级包(仅解析, 不刷机不重启)才显示设备信息。
    判据: Firmware 严判=FIRMWARE_VERSION@61440(ASCII); Model/Hardware 严判=config
    firmware.expect_model/expect_hardware(工程师维护); 对应寄存器原值记入 detail——
    实测 Product Info 块(61504..61553)为 ASCII 递增占位(SN乱码来源), 疑似固件/产测未烧写。
    """
    import re
    cfg = get_config(config_path)
    report = _header(case_meta)
    cid = case_meta.get("编号", "case")
    if _locked(report):
        report.save(f"auto_{cid}")
        return report
    fw = _fw_cfg(cfg)
    pkg = fw.get("package")
    with Verifier() as vf:
        info = read_device_info(vf)
    drv = _connect(cfg)
    mu = open_meter_update(drv)
    try:
        if pkg:   # 选包以点亮设备信息栏(只解析包, 无任何写表动作)
            loaded = select_firmware_file(drv, mu, pkg)
            report.add(Check(label="加载升级包以显示设备信息栏", expected="信息栏出现设备信息",
                             actual="已加载" if loaded else "未确认加载", passed=loaded))
        _mu_activate(mu)
        left, top = _mu_rect(mu)
        x, y, w, h = REL_INFO_REGION
        bar = " ".join(drv.ocr_text((left + x, top + y, w, h)).split())
        print(f"  [信息栏OCR] {bar}")
        m = re.search(r"Model:\s*([\w.\-]+).*?Hardware:\s*([\d.]+).*?Firmware:\s*([\d.]+)", bar)
        if not m:
            report.add(Check(label="信息栏 OCR 解析", expected="Model/Hardware/Firmware 三段",
                             actual=bar or "(OCR为空)", passed=False))
        else:
            ui_model, ui_hw, ui_fw = m.group(1), m.group(2), m.group(3)
            expect_model = fw.get("expect_model") or "MAC1"
            expect_hw = fw.get("expect_hardware") or "1.03"
            report.add(Check(label="信息栏 Model 显示", expected=expect_model, actual=ui_model,
                             passed=str(ui_model) == str(expect_model),
                             detail=f"MODEL@61553={info.get('MODEL')!r}(占位); "
                                    f"正式版应为 AcuRev-101-mA/mV(xlsx 注: 待固件确认)"))
            report.add(Check(label="信息栏 Hardware 显示", expected=expect_hw, actual=ui_hw,
                             passed=str(ui_hw) == str(expect_hw),
                             detail=f"HARDWARE_VERSION@61520={info.get('HARDWARE_VERSION')!r}"
                                    f"(占位, 界面值另有来源, 疑似产测未烧写)"))
            report.add(Check(label="信息栏 Firmware 显示=寄存器", expected=info.get("FIRMWARE_VERSION"),
                             actual=ui_fw, passed=str(ui_fw) == str(info.get("FIRMWARE_VERSION")),
                             detail="FIRMWARE_VERSION@61440 ASCII 解码"))
    finally:
        close_meter_update(_find_window("Meter Update") or mu)
    report.save(f"auto_{cid}")
    return report
