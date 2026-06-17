import pytest
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_007_03_case03_4
# 用例标题：设置最短密码期限为90，密码在90天之内不允许修改
# 说明：需要等待90天验证效果，属于时间依赖用例，不纳入自动化。
@pytest.mark.skip(reason="时间依赖用例：需等待90天才能验证，不纳入自动化（见 usermanagement_struct.md 澄清第16条）")
def test_TestCase_AcuHMI_007_03_case03_4(login_page: LoginPage):
    pass
