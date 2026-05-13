import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_01_case06_8
# 用例标题：验证更新EULA信息，所有用户登录是否需要重新接受用户许可协议
# 跳过原因：需要升级到含有不同版本EULA的固件，依赖固件升级环境，不纳入自动化
@pytest.mark.skip(reason="需要升级固件至含新版本EULA的版本，依赖固件升级环境，不纳入自动化")
def test_TestCase_AcuHMI_007_01_case06_8(login_page: LoginPage):
    pass
