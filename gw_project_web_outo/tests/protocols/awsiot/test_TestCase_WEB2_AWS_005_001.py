import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_WEB2_AWS_005_001
# 用例标题：AWS IoT 24小时重连测试
# 用户确认：此测试不实现自动化，为专项自动化测试，L列(自动化)应设为否
@pytest.mark.skip(reason="用户确认不实现自动化：AWS IoT 24小时重连测试为专项长时间测试，不适合自动化脚本")
def test_TestCase_WEB2_AWS_005_001(login_page: LoginPage):
    pass
