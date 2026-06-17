# -*- coding: utf-8 -*-
"""
factory_reset.py — FactoryReset 操作共用辅助函数

供 system_settings/ 下所有 005_06_caseXX.py 使用。
"""
import sys
from pathlib import Path

from playwright.sync_api import Page

_REPO_ROOT = str(Path(__file__).parents[4])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from projects.AcuHMI_1_7.settings import HMI_URL, DEFAULT_USERNAME, DEFAULT_PASSWORD  # noqa: E402


# ── 导航辅助 ─────────────────────────────────────────────────────────────────

def nav_to(page: Page, hash_path: str) -> None:
    """跳转到指定 hash 路径并等待页面稳定。"""
    base = page.url.split("#")[0]
    page.goto(base + hash_path)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


# ── FactoryReset 触发 ────────────────────────────────────────────────────────

def trigger_factory_reset(page: Page) -> None:
    """在 Maintenance 页面点击 Factory Reset 并在确认框中确认。

    流程：
      1. 导航到 Maintenance 页（hash: #/systemSettings/maintenance 或菜单点击）
      2. 找到 Factory Reset 按钮并点击
      3. 在弹出确认框中点击 Yes / Confirm / 确认

    重要：Factory Reset 会触发设备重启，调用方需在此函数后等待设备重连。
    选择器标注：待真机校验 — 基于 005_01_case01 的 Reboot System 模式推断。
    """
    base = page.url.split("#")[0]
    # 先跳到 dateTime，再通过菜单项导航到 Maintenance（避免直接 goto 时菜单未展开）
    page.goto(base + "#/systemSettings/dateTime")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(600)

    # 点击左侧导航 Maintenance — 选择器待真机校验
    maintenance_item = page.locator(".left-nav-item").filter(has_text="Maintenance").first
    if maintenance_item.count() == 0:
        maintenance_item = page.locator(".el-menu-item").filter(has_text="Maintenance").first
    maintenance_item.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # 点击 Factory Reset 按钮 — 选择器待真机校验
    reset_btn = page.locator("button").filter(has_text="Factory Reset").first
    assert reset_btn.count() > 0, "未找到 Factory Reset 按钮（选择器待真机校验）"
    reset_btn.click()
    page.wait_for_timeout(1000)

    # 确认对话框 — 选择器待真机校验
    for btn_name in ["Yes", "Yes, continue", "Confirm", "确认", "OK"]:
        btn = page.get_by_role("button", name=btn_name)
        if btn.count() > 0 and btn.first.is_visible():
            btn.first.click()
            break
    page.wait_for_timeout(2000)


# ── 重启等待与重登录 ──────────────────────────────────────────────────────────

def wait_for_device_and_relogin(page: Page) -> None:
    """等待设备重启完成并重新登录。

    复用 test_TestCase_AcuHMI_005_01_case01 中已验证的重连模式：
      1. 等 60 秒让设备完成重启
      2. 轮询 goto(HMI_URL)，最多 30 次 × 10 秒 = 5 分钟
      3. 若落在登录页则自动填写凭据并登录
      4. 处理登录后可能弹出的"修改默认密码"对话框
    """
    base = HMI_URL

    # 等设备启动
    page.wait_for_timeout(60_000)

    for _ in range(30):
        try:
            page.goto(base, timeout=20_000)
            page.wait_for_load_state("networkidle", timeout=20_000)
            break
        except Exception:
            page.wait_for_timeout(10_000)

    try:
        page.wait_for_selector(".el-loading-mask", state="hidden", timeout=60_000)
    except Exception:
        pass
    page.wait_for_timeout(2000)

    def _needs_login(p: Page) -> bool:
        return "login" in p.url.lower() or p.locator("input[type='password']").count() > 0

    if _needs_login(page):
        try:
            page.wait_for_selector("button:has-text('Sign In')", state="visible", timeout=30_000)
        except Exception:
            pass
        page.wait_for_timeout(1000)
        page.get_by_role("textbox", name="Enter User Name").fill(DEFAULT_USERNAME)
        page.get_by_role("textbox", name="Enter User Name").press("Tab")
        page.get_by_role("textbox", name="Enter Password").fill(DEFAULT_PASSWORD)
        try:
            page.wait_for_selector(".el-loading-mask", state="hidden", timeout=30_000)
        except Exception:
            pass
        page.get_by_role("button", name="Sign In").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)
        # 处理"修改默认密码"对话框
        try:
            page.get_by_role("button", name="Cancel").click(timeout=3000)
        except Exception:
            pass
        try:
            page.wait_for_selector(".el-loading-mask", state="hidden", timeout=30_000)
        except Exception:
            pass
        page.wait_for_timeout(2000)
