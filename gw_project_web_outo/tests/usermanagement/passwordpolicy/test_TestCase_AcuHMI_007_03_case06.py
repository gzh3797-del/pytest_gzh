import pytest
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_03_case06
# 用例标题：设置宽限期为0，用户密码到期后，立马被锁定，无法登录
# 说明：需要密码过期（等待N天或修改系统时间）才能验证，属于时间依赖用例，不纳入自动化。
@pytest.mark.skip(reason="时间依赖用例：需密码到期后才能验证宽限期行为，不纳入自动化（见 usermanagement_struct.md 澄清第16条）")
def test_TestCase_AcuHMI_007_03_case06(login_page: LoginPage):
    pass
