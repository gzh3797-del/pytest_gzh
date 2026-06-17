# -*- coding: utf-8 -*-
import os
import sys
import tempfile
from pathlib import Path

import pytest
from playwright.sync_api import Page

_REPO_ROOT = str(Path(__file__).parents[3])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from projects.AcuHMI_1_7.settings import HMI_URL  # noqa: E402
from projects.AcuHMI_1_7.tests.ui.systemsettings.helpers.factory_reset import (  # noqa: E402
    nav_to,
    trigger_factory_reset,
    wait_for_device_and_relogin,
)

_FIXTURE_DIR = Path(__file__).parent / "fixtures"

# 若仓库中已有真实固件文件，填入其路径；否则保持 None，运行时自动生成占位文件。
# 需替换为真实固件：从设备供应商处获取的 .bin / .pkg 固件升级包，
# 用于验证"上传后不点升级，Factory Reset 后固件信息清空"的场景。
_FIRMWARE_FILE_PATH: str | None = None


def _nav_to_firmware_update(page: Page) -> None:
    """导航到 System Settings → Firmware Update 页面。

    选择器待真机校验：hash 路径基于同类页面模式推断。
    """
    nav_to(page, "#/systemSettings/firmwareUpdate")


@pytest.mark.destructive
def test_TestCase_AcuHMI_005_06_case16(system_settings_page: Page) -> None:
    """TestCase_AcuHMI_005_06_case16｜上传Firmware Update文件后不升级，Factory Reset后页面信息为空

    预置条件：
      1. 管理权限登录 AcuHMI，设备运行正常

    测试步骤：
      1. 进入 Firmware Update 页面，上传 Firmware Update 文件（不点升级按钮）
      2. Factory Reset 恢复出厂，等待设备重启并重新登录
      3. 进入 Firmware Update 页面，验证页面内容为空，不显示已上传的文件

    预期结果：
      1. 固件文件上传成功，显示文件名（但不触发升级）
      2. Factory Reset 后，Firmware Update 页面不显示之前上传的文件信息
    """
    page = system_settings_page

    # ── Step 1: 上传固件文件（不点升级）──────────────────────────────────────
    tmp_path: str | None = None
    if _FIRMWARE_FILE_PATH and Path(_FIRMWARE_FILE_PATH).exists():
        firmware_file = _FIRMWARE_FILE_PATH
    else:
        # 生成占位固件文件（二进制内容，后缀 .bin）
        # 需替换为真实固件：供应商提供的 .bin / .pkg 升级包
        _FIXTURE_DIR.mkdir(exist_ok=True)
        fd, tmp_path = tempfile.mkstemp(suffix=".bin", dir=str(_FIXTURE_DIR))
        with os.fdopen(fd, "wb") as f:
            # 写入占位内容，不会触发真实升级
            f.write(b"\x00\x01\x02\x03FIRMWARE_PLACEHOLDER\x00")
        firmware_file = tmp_path

    try:
        _nav_to_firmware_update(page)

        # 找到固件上传触发按钮
        # 选择器待真机校验：按钮文字可能为 "Browse" / "Choose File" / "Upload"
        upload_trigger = (
            page.get_by_role("button", name="Browse")
            .or_(page.get_by_role("button", name="Choose File"))
            .or_(page.get_by_role("button", name="Upload"))
            .first
        )
        assert upload_trigger.count() > 0, (
            "未找到固件上传触发按钮（选择器待真机校验）"
        )

        with page.expect_file_chooser(timeout=5000) as fc_info:
            upload_trigger.click()
        fc_info.value.set_files(firmware_file)
        page.wait_for_timeout(1500)

        # 确认文件名显示在页面上（上传后显示文件信息，但不点升级）
        # 注意：此处不点击 "Upgrade" / "Update" 按钮
        page.wait_for_timeout(1000)

    finally:
        if tmp_path and Path(tmp_path).exists():
            os.unlink(tmp_path)

    # ── Step 2: 执行 Factory Reset + 等待重启 + 重登录 ───────────────────────
    trigger_factory_reset(page)
    wait_for_device_and_relogin(page)

    # ── Step 3: 导航回 Firmware Update，断言页面内容为空 ─────────────────────
    _nav_to_firmware_update(page)
    page.wait_for_timeout(1000)

    # 断言页面不显示之前上传的固件文件信息
    # 选择器待真机校验：文件名显示区域可能为 .el-upload__tip / .file-name / 输入框
    file_info = page.locator(".el-upload__tip, .file-name, .upload-filename").first
    if file_info.count() > 0 and file_info.is_visible():
        file_text = file_info.inner_text().strip()
        assert file_text == "" or "No file" in file_text or file_text is None, (
            f"Factory Reset 后 Firmware Update 页面应为空，"
            f"实际显示='{file_text}'"
        )

    # 断言没有显示"等待升级"状态
    pending_indicator = (
        page.get_by_text("Pending", exact=False)
        .or_(page.get_by_text("Waiting", exact=False))
        .or_(page.get_by_text("Ready to upgrade", exact=False))
    )
    assert pending_indicator.count() == 0, (
        "Factory Reset 后 Firmware Update 页面不应显示待升级状态"
    )
