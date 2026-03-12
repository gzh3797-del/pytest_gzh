import logging

from test_case.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

"""
    Function_AcuRev1320_013_04_case7
    Enable场景，第1年手动添加&删除2个节日，配置TOU Schedules 中的任两个Schedule ID，update保存配置成功。
"""

tou = TestAcuviewTou()

start_date = ['06-01', '07-02']
schedule_id = [1, 2]


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()


def test_013_03_case1():
    # 上位机TOU Setting页面设置holiday_setting_enabled打开
    tou.open_holiday_setting_enable()
    tou.year1_holidays_add(holidays_id=2, start_date=start_date, schedule_id=schedule_id)
    # 上位机上点击update
    update_success = tou.click_update()
    assert update_success
    # 删除第2个节日
    tou.tou_year1_holidays_remove(number=2)
    # 上位机上点击update
    update_success = tou.click_update()
    assert update_success


def teardown_function():
    # 恢复TOU设置，上位机上点击reset to default
    tou.tou_reset()
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
