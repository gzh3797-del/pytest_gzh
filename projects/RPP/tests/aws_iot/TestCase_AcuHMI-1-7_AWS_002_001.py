"""
FTS编号: FTS_WEB2_AWS_002_001
用例标题: URL 参数校验与合法值校验
用例级别: LV1

预置条件: 网关已登录，已导航到 AWS IoT 配置页面，AWS IoT 处于 Enable 状态

测试步骤:
  1. 输入含非法字符的 URL（如 TEST_url@iot...），点击 Save，验证被拦截
  2. 输入超过 128 字符的 URL，点击 Save，验证被拦截（合法范围 20~128 字符）
  3. 输入合法 URL，点击 Save，验证保存成功
  4. 清空 URL，点击 Save，验证被拦截

预期结果:
  - 含非法字符的 URL 无法保存，页面给出提示
  - 超过 128 字符的 URL 无法保存
  - 合法 URL 保存成功
  - 空 URL 无法保存
"""

import pytest


ILLEGAL_CHAR_URL = "TEST_url@iot.invalid.com"
OVER_LENGTH_URL  = "a" * 129 + ".iot.amazonaws.com"


def _get_page_message(page) -> str:
    """返回 toast/notification 消息或 Element UI 表单行内校验错误文字（非空即代表有提示）。"""
    # 1) toast / notification / message-box
    try:
        el = page.page.locator(page.RESULT_MSG).first
        el.wait_for(state="visible", timeout=3000)
        txt = el.inner_text().strip()
        if txt:
            return txt
    except Exception:
        pass
    # 2) 表单行内校验错误（el-form-item__error）
    try:
        errors = page.page.locator("css=.el-form-item__error").all()
        for err in errors:
            if err.is_visible():
                txt = err.inner_text().strip()
                if txt:
                    return txt
    except Exception:
        pass
    # 3) 输入框 is-error 状态（说明存在校验失败，但无文字提示）
    try:
        err_inputs = page.page.locator("css=.el-form-item.is-error").all()
        if any(inp.is_visible() for inp in err_inputs):
            return "表单校验错误"
    except Exception:
        pass
    return ""


class TestCase_RPP_AWS_002_001:

    @pytest.mark.aws_iot
    def test_url_validation(self, aws_page, aws_cfg):
        aws_page.ensure_enabled()

        # 场景1：含非法字符的 URL 应被拦截
        aws_page.set_url(ILLEGAL_CHAR_URL)
        aws_page.save()
        assert _get_page_message(aws_page), "含非法字符的 URL 应被拦截并给出提示"

        # 场景2：超过 128 字符的 URL 应被拦截
        aws_page.set_url(OVER_LENGTH_URL)
        aws_page.save()
        assert _get_page_message(aws_page), "超过 128 字符的 URL 应被拦截"

        # 场景3：合法 URL 应保存成功
        aws_page.set_url(aws_cfg["aws_iot"]["url"])
        aws_page.save()
        assert "login" not in aws_page.page.url.lower(), "合法 URL 应能保存成功"

        # 场景4：空 URL 应被拦截
        aws_page.page.locator(aws_page.URL_INPUT).fill("")
        aws_page.save()
        assert _get_page_message(aws_page), "空 URL 应被拦截"
