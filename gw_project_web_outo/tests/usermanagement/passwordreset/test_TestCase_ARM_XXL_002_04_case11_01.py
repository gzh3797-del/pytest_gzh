import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_ARM_XXL_002_04_case11_01
# 用例标题：验证默认密码功能
# 跳过原因：需要设备处于出厂状态（烧录升级后首次使用），自动化无法安全复现
# 澄清汇总 #13：不实现自动化
@pytest.mark.skip(reason="需要设备出厂状态，自动化无法安全复现；澄清汇总 #13 不实现自动化")
def test_TestCase_ARM_XXL_002_04_case11_01(login_page: LoginPage):
    pass
