import logging

from projects.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

"""
    Function_AcuRev1320_013_04_case2
    Disable场景，添加&删除2个节日，配置TOU Schedules中的任两个Schedule ID，update保存配置成功。
"""

tou = TestAcuviewTou()


# start_date = ['05-10', '08-10']
# schedule_id = [1, 2]


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()


def test_013_03_case1():
    # 上位机TOU Setting页面设置holiday_setting_enable关闭
    tou.close_holiday_setting_enable()
    # 上位机上点击update
    update_success = tou.click_update()
    # 进入上位机tou_setting页面
    tou.enter_tou_setting()
    # TOU_enable使能打开
    tou.open_TOU_enable()
    # 上位机TOU Setting页面设置billing_and_tariff
    tou.billing_and_tariff_Setting(billing_mode=0, at_time='01 00:30:00', tariff=3)
    # 配置1个Schedule Id，Schedule ID对应配置1个segment ID
    tou.tou_schedules_add(schedule_id=2, segment_id=1)
    # 配置1个session,对应Schedule Id 1
    tou.tou_seasons_add(session_id=1, seasons_schedule_id=[1])
    # 添加2个节日
    tou.tou_holidays_add(holidays_id=30)

    result = tou.tou_holidays_add_31()
    assert result
    # 上位机上点击update
    update_success = tou.click_update()
    assert update_success
    # 删除2个节日
    tou.tou_holidays_remove(number=30)
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
