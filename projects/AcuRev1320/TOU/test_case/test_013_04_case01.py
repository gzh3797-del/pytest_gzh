import logging

from projects.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

"""
    Function_AcuRev1320_013_04_case1
    十年节假日配置使能开关，enable和disable保存配置成功
"""

tou = TestAcuviewTou()


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()


def test_013_03_case1():
    # 上位机TOU Setting页面设置holiday_setting_enable打开
    tou.open_holiday_setting_enable()
    # 上位机上点击update
    update_success = tou.click_update()
    assert update_success
    # 上位机TOU Setting页面设置holiday_setting_enable关闭
    tou.close_holiday_setting_enable()
    # 上位机上点击update
    update_success = tou.click_update()
    assert update_success


def teardown_function():
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
