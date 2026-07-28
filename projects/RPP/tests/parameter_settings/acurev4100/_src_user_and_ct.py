# -*- coding: utf-8 -*-
"""
_src_user_and_ct.py — AcuRev4100 "User and CT" 页面操作 + 接线方式(Wiring)校验

页面路径：Physical Devices → <设备> → Settings → User and CT（顶部 Settings 下拉子项）
页面三区块：Wiring Configuration / Current Input Channel(24 行) / User and Channel Mapping(12 行)

服务用例：TestCase_AcuHMI_001_07_case38~42（接入设备参数设置 → AcuRev4100 接线方式验证）

寄存器来源：AcuRev4100 Modbus Address Table v1.02 20260202.xlsx（Basic Setting sheet）
  Service Configuration(Wiring)   0x1042 = 4162   0:1E2W 1:2E3W1P 2:2E3W_Delta 3:2E3W_Network 4:3E4WY
  channel N voltage assignment    0x104A = 4170 起，步长 5：REG = 4170 + (N-1)*5
                                  编码 0:Va 1:Vb 2:Vc 3:Vab 4:Vbc 5:Vca
Modbus 读取/校验复用 helpers_4100（连接常量由 acurev4100/conftest.py 会话级绑定后即时生效）。

真机行为（现场探查实测，详见 knowledge/meters/RPP/requirements/context/UserAndCT_context.md）：
  - 切换 Wiring **自动保存**（无需点 Save），触发约 5s "Running" 忙态；
    可靠判据：轮询底部 Save 按钮文案，由 "Running" 变回 "Save" 即结束。
  - 各模式 Voltage Assignment 列规律 / 是否可编辑 / User Channel 区块显隐见各 verify_* 与 test 用例。
"""
import re
import time
import logging

import allure
from playwright.sync_api import Page

# Modbus 读取/校验 + step 复用 helpers_4100（其连接常量由 conftest 会话级绑定后即时生效）
from helpers_4100 import modbus_read, verify_modbus, step

_log = logging.getLogger("acurev4100_test")

# ── 寄存器 ────────────────────────────────────────────────────────────
REG_SERVICE_CONFIG = 4162      # 0x1042  接线方式(Service Configuration)
REG_CH1_VOLT_ASSIGN = 4170     # 0x104A  通道 1 Voltage Assignment；通道 N = 4170+(N-1)*5
VA_STRIDE = 5
CH_COUNT = 24

# 接线方式：页面文案 → 寄存器编码
WIRING_ENCODE = {
    "1 Element 2 Wire":         0,
    "2 Element 3 Wire 1 Phase": 1,
    "2 Element 3 Wire Delta":   2,
    "2 Element 3 Wire Network": 3,
    "3 Element 4 Wire Y":       4,
}
# Voltage Assignment：页面文案 → 寄存器编码
VA_ENCODE = {"Va": 0, "Vb": 1, "Vc": 2, "Vab": 3, "Vbc": 4, "Vca": 5}

DEFAULT_WIRING = "3 Element 4 Wire Y"


def va_reg(channel: int) -> int:
    """通道(1-based) Voltage Assignment 寄存器地址。"""
    return REG_CH1_VOLT_ASSIGN + (channel - 1) * VA_STRIDE


def expected_va_encode(mode: str, channel: int) -> int:
    """按接线方式返回通道(1-based)期望的 Voltage Assignment 编码。

    实测规律：
      1E2W           : 全部 Va
      2E3W1P         : 奇数通道 Va, 偶数通道 Vc
      2E3W Delta     : 奇数通道 Vab, 偶数通道 Vbc
      2E3W Network   : Va/Vb/Vc 三通道循环
      3E4WY          : Va/Vb/Vc 三通道循环
    """
    if mode == "1 Element 2 Wire":
        return VA_ENCODE["Va"]
    if mode == "2 Element 3 Wire 1 Phase":
        return VA_ENCODE["Va"] if channel % 2 == 1 else VA_ENCODE["Vc"]
    if mode == "2 Element 3 Wire Delta":
        return VA_ENCODE["Vab"] if channel % 2 == 1 else VA_ENCODE["Vbc"]
    # 2E3W Network / 3E4WY：Va/Vb/Vc 三行循环（0=Va,1=Vb,2=Vc）
    return (channel - 1) % 3


# ── 页面元素定位 ──────────────────────────────────────────────────────

def _save_button(page: Page):
    """底部固定操作栏中的 Save 按钮（文案在 Save / Running 间切换）。"""
    btn = page.locator(".buttonFixed.c_common_button_fix button")
    if btn.count() == 0:
        btn = page.locator("button:has-text('Save')")
    return btn.first


def _input_table(page: Page):
    """Current Input Channel 表（.custom-table 第 1 个）。"""
    return page.locator(".custom-table").nth(0)


def _user_table(page: Page):
    """User and Channel Mapping 表（.custom-table 第 2 个）。"""
    return page.locator(".custom-table").nth(1)


def _va_wrapper(page: Page, row_idx: int):
    """Current Input Channel 第 row_idx(0-based) 行的 Voltage Assignment 下拉（td 第 6 列）。"""
    return (_input_table(page).locator("tbody tr").nth(row_idx)
            .locator("td").nth(5).locator(".el-select__wrapper"))


def _phase_wrapper(page: Page, uc_idx: int, phase: str):
    """User and Channel Mapping 第 uc_idx(0-based) 行的 Phase A/B/C 下拉。"""
    col = {"A": 2, "B": 3, "C": 4}[phase]
    return (_user_table(page).locator("tbody tr").nth(uc_idx)
            .locator("td").nth(col).locator(".el-select__wrapper"))


# ── 读取 / 判定 ───────────────────────────────────────────────────────

def read_va_cell_text(page: Page, row_idx: int) -> str:
    """读取某行 Voltage Assignment 当前显示值（如 'Va'/'Vab'）。"""
    return _va_wrapper(page, row_idx).inner_text().strip()


def va_is_disabled(page: Page, row_idx: int) -> bool:
    """某行 Voltage Assignment 下拉是否为 disabled 只读。"""
    cls = _va_wrapper(page, row_idx).get_attribute("class") or ""
    return "is-disabled" in cls


def open_va_options(page: Page, row_idx: int) -> list[str]:
    """点开某行 Voltage Assignment 下拉，返回可选项文字列表（读完 Esc 关闭）。"""
    w = _va_wrapper(page, row_idx)
    w.scroll_into_view_if_needed()
    w.click()
    page.wait_for_timeout(500)
    items = page.locator(".el-select-dropdown__item:visible")
    opts = [items.nth(i).inner_text().strip() for i in range(items.count())]
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)
    _log.info("[VA OPTIONS] row%d → %s", row_idx + 1, opts)
    return opts


def user_mapping_visible(page: Page) -> bool:
    """User and Channel Mapping 区块是否**真正可见**（1E2W 下该表隐藏但仍在 DOM，
    故须判 is_visible，而非仅按 .custom-table 计数）。"""
    tbl = page.locator(".custom-table").nth(1)
    return tbl.count() > 0 and tbl.is_visible()


def user_channel_row_count(page: Page) -> int:
    """User and Channel Mapping 表可见行数（预期 12）。"""
    if not user_mapping_visible(page):
        return 0
    return _user_table(page).locator("tbody tr").count()


def phase_is_disabled(page: Page, uc_idx: int, phase: str) -> bool:
    """某 User Channel 行的 Phase A/B/C 下拉是否 disabled/固定。"""
    cls = _phase_wrapper(page, uc_idx, phase).get_attribute("class") or ""
    return "is-disabled" in cls


# ── 页面操作 ──────────────────────────────────────────────────────────

@allure.step("导航到 User and CT 页面")
def nav_to_user_and_ct(page: Page) -> None:
    """Physical Devices → <设备> → Settings → User and CT。

    设备名取 helpers_4100.DEVICE_NAME（conftest 会话级动态发现后绑定）。
    等待 'Wiring Configuration' 文本出现即表示页面加载完毕。
    """
    from helpers_4100 import DEVICE_NAME  # 调用时读取已绑定值

    step("Click Physical Devices")
    page.locator(
        "a:has-text('Physical Devices'), "
        "span:has-text('Physical Devices'), "
        "li:has-text('Physical Devices')"
    ).first.click()
    page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)

    step(f"Click device: {DEVICE_NAME}")
    page.locator(f"text={DEVICE_NAME}").first.click()
    page.wait_for_selector("text=Settings", timeout=10000)

    _settings_loc = page.locator(
        "span:has-text('Settings'), a:has-text('Settings'), button:has-text('Settings')"
    )
    _uandct_loc = page.locator(
        "a:has-text('User and CT'), li:has-text('User and CT'), span:has-text('User and CT')"
    )
    step("Open Settings menu → User and CT")
    clicked = False
    for _ in range(4):
        _settings_loc.first.click()
        page.wait_for_timeout(700)
        if _uandct_loc.first.is_visible():
            try:
                _uandct_loc.first.click(timeout=3000)
                clicked = True
                break
            except Exception:
                pass  # 菜单在 click 前收起，继续重试
    if not clicked:
        _log.warning("nav: Settings 4 次未展开 → reload 重置页面")
        page.reload()
        page.wait_for_timeout(1500)
        page.wait_for_selector(f"text={DEVICE_NAME}", timeout=10000)
        page.locator(f"text={DEVICE_NAME}").first.click()
        page.wait_for_selector("text=Settings", timeout=10000)
        page.wait_for_timeout(500)
        _settings_loc.first.click()
        page.wait_for_timeout(700)
        if _uandct_loc.first.is_visible():
            _uandct_loc.first.click()
        else:
            _uandct_loc.first.click(force=True, timeout=5000)

    page.wait_for_selector("text=Wiring Configuration", timeout=12000)
    page.wait_for_timeout(800)
    step("Ready: User and CT")


def wait_running_done(page: Page, timeout_ms: int = 25000) -> None:
    """切换 Wiring 后轮询底部 Save 按钮文案，由 'Running' 变回 'Save' 视为结束。

    实测切换重算耗时稳定 ~5s。仅用按钮文案判据（绿色✅图标转瞬即逝不可靠；
    VA 下拉可交互性在 1E2W/Delta 下永久 disabled，也不能作判据）。
    """
    btn = _save_button(page)
    deadline = time.time() + timeout_ms / 1000.0
    txt = ""
    while time.time() < deadline:
        try:
            txt = btn.inner_text().strip()
        except Exception:
            txt = ""
        if "Running" not in txt:
            page.wait_for_timeout(800)  # 额外沉降，等表格重算落定
            return
        page.wait_for_timeout(500)
    _log.warning("wait_running_done: 超时，Save 按钮文案仍为 %r", txt)


@allure.step("切换接线方式 → {mode}")
def switch_wiring(page: Page, mode: str) -> None:
    """打开 Wiring Configuration 下拉选中 mode，等待 Running 重算结束（自动保存，无需点 Save）。"""
    step(f"Switch Wiring → {mode}")
    card = page.locator(".c_card_layout").filter(has_text="Wiring Configuration")
    wrapper = card.locator(".el-select__wrapper").first
    wrapper.scroll_into_view_if_needed()
    wrapper.click()
    page.wait_for_timeout(500)

    target = page.locator(".el-select-dropdown__item:visible").get_by_text(mode, exact=True)
    if target.count() == 0:
        items = page.locator(".el-select-dropdown__item:visible")
        opts = [items.nth(i).inner_text().strip() for i in range(items.count())]
        page.keyboard.press("Escape")
        raise RuntimeError(f"switch_wiring: 选项 {mode!r} 未找到，可见项：{opts}")
    target.first.click()
    page.wait_for_timeout(1000)  # 给忙态一点时间进入 Running
    wait_running_done(page)
    wait_modbus_ready()          # 切换后设备 Modbus 可能短暂复位，等其可读再继续
    step(f"Wiring switched: {mode}")


def wait_modbus_ready(rounds: int = 5) -> None:
    """切接线后设备 Modbus 常短暂复位/拒读（尤以 1E2W 大拓扑切换为甚），
    轮询 Service Config 直到可读或轮次用尽。每轮内 modbus_read 自带 5×3s 重试。"""
    for i in range(rounds):
        try:
            modbus_read(REG_SERVICE_CONFIG, count=1)
            return
        except Exception as exc:  # noqa: BLE001  就绪探测，容忍复位期异常
            _log.warning("wait_modbus_ready: round %d/%d 仍不可读 (%s)", i + 1, rounds, exc)
    _log.warning("wait_modbus_ready: %d 轮后仍不可读，交由后续断言暴露", rounds)


# ── Modbus 校验 ───────────────────────────────────────────────────────

@allure.step("校验接线方式寄存器 = {mode}")
def verify_wiring_reg(mode: str) -> None:
    """Modbus 回读 Service Configuration，确认接线方式已生效。"""
    verify_modbus(REG_SERVICE_CONFIG, WIRING_ENCODE[mode], label=f"ServiceConfig({mode})")


def verify_va_pattern(mode: str) -> None:
    """对全部 24 通道，按接线方式期望校验 Voltage Assignment 寄存器（Modbus 回读）。"""
    bad = []
    for ch in range(1, CH_COUNT + 1):
        exp = expected_va_encode(mode, ch)
        regs = modbus_read(va_reg(ch), count=1)
        ok = regs[0] == exp
        _log.info("[VERIFY VA] ch%-2d  expected=%-2s  actual=%-2s  → %s",
                  ch, exp, regs[0], "✓" if ok else "✗")
        if not ok:
            bad.append((ch, exp, regs[0]))
    assert not bad, (f"VA 寄存器不符 ({mode}): "
                     + ", ".join(f"ch{c}:期望{e}实得{g}" for c, e, g in bad))


# ════════════════════════════════════════════════════════════════════
# User Channel Description（组 B：User Channel 命名，032_002_003~015）
# ════════════════════════════════════════════════════════════════════

REG_UC1_DESC = 4608     # 0x1200  User Channel 1 Description；UserN = 4608+(N-1)*10
UC_DESC_STRIDE = 10     # 每个 UC 10 个寄存器（ASCII 20 字节）
UC_COUNT = 12
UC_DESC_MAXLEN = 20     # 名称最大长度（>20 保存被拒）

# 失败提示文案片段（Element Plus 顶部 .el-message--warning toast，实测）
ERR_DESC_ASCII = "must contain only ASCII"
ERR_DESC_LENGTH = "less than 20"


def uc_desc_reg(uc: int) -> int:
    """User Channel(1-based) Description 首寄存器地址。"""
    return REG_UC1_DESC + (uc - 1) * UC_DESC_STRIDE


def _desc_input(page: Page, uc_idx: int):
    """User and Channel Mapping 第 uc_idx(0-based) 行的 Description 输入框。"""
    return (_user_table(page).locator("tbody tr").nth(uc_idx)
            .locator("td").nth(1).locator("input"))


def ensure_user_mapping(page: Page) -> None:
    """确保 User and Channel Mapping 区块可见（组 B 需要）；不可见时切到默认接线方式。"""
    if not user_mapping_visible(page):
        _log.info("[UC DESC] User Channel 区块不可见 → 切到 %s", DEFAULT_WIRING)
        switch_wiring(page, DEFAULT_WIRING)


def set_description(page: Page, uc_idx: int, text: str) -> None:
    """清空并填入某 User Channel 行的 Description（不点 Save）。"""
    step(f"Set UC{uc_idx + 1} Description = {text!r}")
    inp = _desc_input(page, uc_idx)
    inp.scroll_into_view_if_needed()
    inp.click()
    inp.press("Control+a")
    inp.press("Delete")
    if text:
        inp.type(text, delay=30)
    _log.info("  UC%d Description typed %r", uc_idx + 1, text)


def read_description_input(page: Page, uc_idx: int) -> str:
    """读取某 User Channel 行 Description 输入框当前值。"""
    return _desc_input(page, uc_idx).input_value()


def read_description_modbus(uc: int) -> str:
    """Modbus 回读 User Channel(1-based) Description（10 寄存器 ASCII，去尾部 NUL/空白）。

    每寄存器 2 字节，高字节在前（big-endian）。跨传输权威校验，替代仅看页面回显。
    """
    regs = modbus_read(uc_desc_reg(uc), count=UC_DESC_STRIDE)
    chars = []
    for reg in regs:
        chars.append((reg >> 8) & 0xFF)
        chars.append(reg & 0xFF)
    # 实测该字段存储会夹杂 NUL 填充（首/中/尾均可能），名称本身不含 NUL，
    # 故滤除所有 0x00 后再解 ASCII（与页面显示的有效字符一致）。
    raw = bytes(b for b in chars if b != 0)
    return raw.decode("ascii", errors="replace").rstrip()


@allure.step("点击 Save（User and CT）并返回结果")
def save_user_and_ct(page: Page) -> tuple[bool, str]:
    """点击底部 Save，返回 (是否出现警告/失败, 提示文案)。

    成功与否以 Modbus 回读为准（成功 toast class 未固定）；本函数只负责捕获
    失败/警告 toast（`.el-message--warning`，实测非法保存与 'No change to save' 都是该类）。
    返回 warning=True 表示出现了警告类 toast（含校验失败与无变化）。
    """
    step("Click Save (User and CT)")
    _save_button(page).click()
    page.wait_for_timeout(1500)
    msg = ""
    warning = False
    for sel in [".el-message--warning", ".el-message--error"]:
        loc = page.locator(sel)
        if loc.count() > 0 and loc.first.is_visible():
            msg = loc.first.inner_text().strip()
            warning = True
            _log.info("[SAVE] warning/error toast: %r", msg)
            break
    if warning and "no change" in msg.lower():
        # "No change to save"：值已是目标值，非校验失败；视为成功（后续回读仍校验实际值）
        _log.info("[SAVE] 'No change to save' → 视为成功（值已就绪）")
        warning = False
    if not warning:
        # 可能是成功 toast（class 未固定）或无 toast——记录后交由 Modbus 回读判定
        if not msg:
            succ = page.locator(".el-message--success, .el-message")
            if succ.count() > 0 and succ.first.is_visible():
                msg = succ.first.inner_text().strip()
                _log.info("[SAVE] toast: %r", msg)
            else:
                _log.info("[SAVE] 无 toast")
    page.wait_for_timeout(500)
    return warning, msg


@allure.step("校验 UC{uc} Description 寄存器 = {expected!r}")
def verify_description_modbus(uc: int, expected: str) -> None:
    """Modbus 回读并断言 User Channel Description。"""
    actual = read_description_modbus(uc)
    ok = actual == expected
    _log.info("[VERIFY DESC] UC%-2d  expected=%-24r actual=%-24r → %s",
              uc, expected, actual, "✓ PASS" if ok else "✗ FAIL")
    assert ok, f"UC{uc} Description: 期望 {expected!r}, 实得 {actual!r}"


def displayed_name(page: Page, uc_idx: int) -> str:
    """某 User Channel 行 Description 为空时页面回退显示的名称（输入框 placeholder 或值）。

    实测：Description 空值保存后，界面回退显示 'User Channel N'。优先取输入框当前值，
    为空则取 placeholder。
    """
    inp = _desc_input(page, uc_idx)
    val = (inp.input_value() or "").strip()
    if val:
        return val
    return (inp.get_attribute("placeholder") or "").strip()


# ════════════════════════════════════════════════════════════════════
# 跨页面显示（_012 立即生效 / _013 跨界面传播）
#   Metering → Realtime/Demand/Energy/Sequence：Meter Point 下拉显示
#     "User Channel N:<Description>"（改名随之切换）
#   Metering → THD/Harmonics：Input Channel 维度，不含 User Channel 名（"除外"）
# ════════════════════════════════════════════════════════════════════

# 会展示 User Channel 名的 Metering 子页（Meter Point 下拉）
METERING_UC_PAGES = ["Realtime", "Demand", "Energy", "Sequence"]
# 不展示 User Channel 名的 Metering 子页（Input Channel 维度）
METERING_INPUT_PAGES = ["THD", "Harmonics"]


@allure.step("导航到 Metering → {leaf}")
def nav_to_metering(page: Page, leaf: str) -> None:
    """从设备详情顶部菜单进入 Metering → <leaf> 子页。"""
    step(f"Metering → {leaf}")
    navbar = page.locator(".c_top_navbar")
    navbar.locator(".el-sub-menu__title").filter(has_text="Metering").first.click()
    page.wait_for_timeout(500)
    page.locator(".el-menu--popup:visible .el-menu-item:visible").filter(
        has_text=leaf).first.click()
    page.wait_for_selector(".el-select__wrapper", timeout=10000)
    page.wait_for_timeout(600)


def select_meter_point_uc(page: Page, uc: int) -> str:
    """在 Metering 页把 Meter Point 下拉切到 User Channel(1-based) uc，返回下拉当前显示文本。

    显示格式 "User Channel N:<Description>"。用 aria-controls 精确定位该下拉自己的
    popper，避免误取页面其它下拉的隐藏项。
    """
    mp = page.locator(".el-select__wrapper").filter(has_text="User Channel").first
    mp.scroll_into_view_if_needed()
    mp.click()
    page.wait_for_timeout(400)
    aria = mp.locator(".el-select__input").get_attribute("aria-controls")
    listbox = page.locator(f"[id='{aria}']") if aria else page
    # 锚定正则：命名时文本为 "User Channel N:desc"，空值时为 "User Channel N"；
    # (:|$) 收尾可区分 UC1 与 UC12（避免 "User Channel 1" 误命中 "User Channel 12"）。
    listbox.locator(".el-select-dropdown__item").filter(
        has_text=re.compile(rf"User Channel {uc}(:|$)")).first.click()
    page.wait_for_timeout(500)
    txt = mp.inner_text().strip()
    _log.info("[METER POINT] UC%d 显示 → %r", uc, txt)
    return txt


def read_input_channel_select_text(page: Page) -> str:
    """THD/Harmonics 页读取 Input Channel 下拉当前文本（应为 'Input Channel N'，不含 UC 名）。"""
    return page.locator(".el-select__wrapper").first.inner_text().strip()


# 让 from _src_user_and_ct import * 包含单下划线工具名（须放文件最末）
__all__ = [_n for _n in globals() if not _n.startswith('__')]
