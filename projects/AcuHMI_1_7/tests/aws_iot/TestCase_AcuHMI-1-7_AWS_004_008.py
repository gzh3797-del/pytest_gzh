"""
FTS编号: FTS_AcuHMI_AWS_004_008
用例标题: 未勾选设备时无法保存并提示
用例级别: LV2

预置条件:
  - Enable 已开启，合法 URL/Topic/证书已配置

测试步骤:
  1. 先配置合法参数并勾选一台设备（确保存在已勾选状态）
  2. 取消勾选全部设备
  3. 点击保存

预期结果:
  - 阻止保存，提示 "Please provide Devices Selection" 或表单校验错误
"""

from pathlib import Path

import pytest

_THIS_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent


def _uncheck_all_devices(page):
    rows = page.page.locator(page.DEVICE_TABLE_ROWS).all()
    for row in rows:
        cbs = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
        if cbs and "is-checked" in (cbs[0].get_attribute("class") or ""):
            cbs[0].evaluate("el => el.click()")
            page.page.wait_for_timeout(300)


def _get_page_message(page) -> str:
    """返回 toast/notification 消息或 Element UI 表单行内校验错误文字。"""
    try:
        el = page.page.locator(page.RESULT_MSG).first
        el.wait_for(state="visible", timeout=3000)
        txt = el.inner_text().strip()
        if txt:
            return txt
    except Exception:
        pass
    try:
        errors = page.page.locator("css=.el-form-item__error").all()
        for err in errors:
            if err.is_visible():
                txt = err.inner_text().strip()
                if txt:
                    return txt
    except Exception:
        pass
    try:
        err_inputs = page.page.locator("css=.el-form-item.is-error").all()
        if any(inp.is_visible() for inp in err_inputs):
            return "表单校验错误"
    except Exception:
        pass
    return ""


class TestCase_AcuHMI_1_7_AWS_004_008:

    @pytest.mark.aws_iot
    def test_save_blocked_without_device(self, aws_page, aws_cfg):
        aws = aws_cfg["aws_iot"]
        # 先建立合法状态：Enable + 证书 + 选一台设备
        aws_page.ensure_enabled()
        aws_page.set_url(aws["url"])
        aws_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
        aws_page.select_only_device()  # 随机选一台物理设备
        aws_page.save()
        # 再取消全部设备勾选
        _uncheck_all_devices(aws_page)
        aws_page.save()
        msg = _get_page_message(aws_page)
        assert msg, "未勾选设备时 Save 应被阻止并给出提示"
