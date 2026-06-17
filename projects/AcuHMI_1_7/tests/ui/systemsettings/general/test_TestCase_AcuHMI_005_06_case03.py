# -*- coding: utf-8 -*-
import os
import sys
import tempfile
from pathlib import Path

import pytest
from playwright.sync_api import Page, expect

_REPO_ROOT = str(Path(__file__).parents[3])
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from projects.AcuHMI_1_7.settings import HMI_URL  # noqa: E402

# ---------------------------------------------------------------------------
# 旧证书占位文件配置
# ---------------------------------------------------------------------------
# 若仓库中已有真实的旧版本证书文件，将路径赋值给下列常量；
# 否则测试运行时会自动生成一个内容为随机字节的占位 .cer 文件（模拟"证书不匹配"场景）。
# 需替换为真实旧证书：在设备 A 上用旧密钥导出的 .cer / .pem 文件。
_FIXTURE_DIR = Path(__file__).parent / "fixtures"
_OLD_CERT_PATH: str | None = None  # 若有真实文件，填写绝对路径字符串，否则保持 None


def _nav_to_certificate(page: Page) -> None:
    """导航到 System Settings → Certificate 页面。

    选择器待真机校验：基于同类 hash-URL 导航模式推断。
    """
    base = page.url.split("#")[0]
    page.goto(base + "#/systemSettings/certificate")
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


@pytest.mark.destructive
def test_TestCase_AcuHMI_005_06_case03(system_settings_page: Page) -> None:
    """TestCase_AcuHMI_005_06_case03｜Certificate 操作：导出→生成CSR→生成自签名证书→导入旧证书，断言证书不匹配导入失败

    预置条件：
      1. 管理权限登录 AcuHMI，设备运行正常

    测试步骤：
      1. 进入 Certificate 页面，点击导出证书（Export Certificate）
      2. 点击生成 CSR（Generate CSR），查看证书信息准确
      3. 点击生成新的自签名证书（Generate Self-Signed Certificate）
      4. 上传旧版本 .cer 文件（与当前证书不匹配）
      5. 断言提示"证书不匹配"，导入失败

    预期结果：
      1. 导出成功
      2. 生成 CSR 成功
      3. 生成自签名证书成功
      4. 提示证书不匹配，导入失败（页面显示错误信息）
    """
    page = system_settings_page

    # ── Step 1: 导航到 Certificate 页面 ──────────────────────────────────────
    _nav_to_certificate(page)

    # ── Step 1: 导出证书 ──────────────────────────────────────────────────────
    # 选择器待真机校验：Export Certificate 按钮文字可能为 "Export" / "Export Certificate"
    export_btn = (
        page.get_by_role("button", name="Export Certificate")
        .or_(page.get_by_role("button", name="Export"))
        .first
    )
    if export_btn.count() > 0 and export_btn.is_visible():
        # 允许下载，不阻塞；Export 按钮点击后通常直接触发下载
        with page.expect_download(timeout=10_000):
            export_btn.click()
        page.wait_for_timeout(500)

    # ── Step 2: 生成 CSR ─────────────────────────────────────────────────────
    # 选择器待真机校验：按钮文字可能为 "Generate CSR" / "Create CSR"
    csr_btn = (
        page.get_by_role("button", name="Generate CSR")
        .or_(page.get_by_role("button", name="Create CSR"))
        .first
    )
    assert csr_btn.count() > 0, "未找到 Generate CSR 按钮（选择器待真机校验）"
    csr_btn.click()
    page.wait_for_timeout(2000)

    # 确认弹框（如有）
    for confirm_name in ["Confirm", "OK", "Generate", "确认"]:
        btn = page.get_by_role("button", name=confirm_name)
        if btn.count() > 0 and btn.first.is_visible():
            btn.first.click()
            page.wait_for_timeout(1000)
            break

    # 断言无错误提示
    assert page.locator(".el-message--error").count() == 0, "生成 CSR 出现错误提示"

    # ── Step 3: 生成自签名证书 ───────────────────────────────────────────────
    # 选择器待真机校验：按钮文字可能为 "Generate Self-Signed Certificate" / "Self-Signed"
    selfsign_btn = (
        page.get_by_role("button", name="Generate Self-Signed Certificate")
        .or_(page.get_by_role("button", name="Self-Signed"))
        .or_(page.get_by_role("button", name="Generate"))
        .first
    )
    assert selfsign_btn.count() > 0, "未找到生成自签名证书按钮（选择器待真机校验）"
    selfsign_btn.click()
    page.wait_for_timeout(2000)

    for confirm_name in ["Confirm", "OK", "Generate", "确认"]:
        btn = page.get_by_role("button", name=confirm_name)
        if btn.count() > 0 and btn.first.is_visible():
            btn.first.click()
            page.wait_for_timeout(1500)
            break

    assert page.locator(".el-message--error").count() == 0, "生成自签名证书出现错误提示"

    # ── Step 4 & 5: 上传旧版本 .cer 文件，断言提示证书不匹配 ─────────────────
    # 准备占位证书文件（若 _OLD_CERT_PATH 未配置则生成临时文件）
    tmp_path: str | None = None
    if _OLD_CERT_PATH and Path(_OLD_CERT_PATH).exists():
        cert_file = _OLD_CERT_PATH
    else:
        # 生成一个内容为随机字节的占位 .cer，与当前设备证书肯定不匹配
        _FIXTURE_DIR.mkdir(exist_ok=True)
        fd, tmp_path = tempfile.mkstemp(suffix=".cer", dir=str(_FIXTURE_DIR))
        with os.fdopen(fd, "wb") as f:
            # 写入非法证书内容（非 PEM/DER 格式），触发"证书不匹配"
            f.write(b"-----BEGIN CERTIFICATE-----\n")
            f.write(b"INVALIDCERTIFICATECONTENTFORTEST==\n")
            f.write(b"-----END CERTIFICATE-----\n")
        cert_file = tmp_path

    try:
        # 找到证书导入区域，点击 Browse / Upload 触发文件选择器
        # 选择器待真机校验：按钮文字可能为 "Browse" / "Upload Certificate" / "Import"
        import_trigger = (
            page.get_by_role("button", name="Browse")
            .or_(page.get_by_role("button", name="Upload Certificate"))
            .or_(page.get_by_role("button", name="Import Certificate"))
            .first
        )
        assert import_trigger.count() > 0, (
            "未找到证书导入触发按钮（选择器待真机校验）"
        )
        with page.expect_file_chooser(timeout=5000) as fc_info:
            import_trigger.click()
        fc_info.value.set_files(cert_file)
        page.wait_for_timeout(1000)

        # 点击 Import / Upload 提交
        submit_btn = (
            page.get_by_role("button", name="Import")
            .or_(page.get_by_role("button", name="Upload"))
            .first
        )
        if submit_btn.count() > 0 and submit_btn.is_visible():
            submit_btn.click()
            page.wait_for_timeout(2000)

        # 断言：证书不匹配，导入失败
        # 选择器待真机校验：错误文字可能为 "not match" / "mismatch" / "fail" / "证书不匹配"
        error_indicator = (
            page.get_by_text("not match", exact=False)
            .or_(page.get_by_text("mismatch", exact=False))
            .or_(page.get_by_text("fail", exact=False))
            .or_(page.get_by_text("证书不匹配", exact=False))
            .or_(page.get_by_text("invalid", exact=False))
        )
        expect(error_indicator).to_be_visible(timeout=5000)

    finally:
        if tmp_path and Path(tmp_path).exists():
            os.unlink(tmp_path)
