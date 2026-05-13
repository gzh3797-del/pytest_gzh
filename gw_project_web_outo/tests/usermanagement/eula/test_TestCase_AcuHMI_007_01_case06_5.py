import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_01_case06_5
# 用例标题：验证默认用户首次登录，是否需要接受EULA
# 跳过原因：需要设备出厂状态（烧录升级后首次使用），无法在已投入使用的设备上自动化复现
@pytest.mark.skip(reason="需要设备处于出厂状态（首次烧录后），自动化无法安全复现此前置条件")
def test_TestCase_AcuHMI_007_01_case06_5(login_page: LoginPage):
    pass
