import logging

from projects.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

"""
    Function_AcuRev1320_013_04_case10
"""

tou = TestAcuviewTou()

segment_time = ['00:00', '01:00', '02:00', '03:00', '04:00', '05:00', '06:00',
                '07:00', '08:00', '09:00', '12:00', '13:00', '14:00', '15:00']
segment_tariff = [0, 1, 2, 3, 0, 1, 2, 3, 0, 1, 2, 3, 0, 1]
start_date = ['01-01', '02-01', '03-01', '04-01', '05-01', '06-01',
              '07-01', '08-01', '09-01', '10-01', '11-01', '12-01']
seasons_schedule_id = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]


# holidays_start_date = ['06-01', '07-02']
# schedule_id = [1, 2]


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()
    # 进入上位机tou_setting页面
    tou.enter_tou_setting()
    # TOU_enable使能打开
    tou.open_TOU_enable()
    # 上位机TOU Setting页面设置billing_and_tariff
    tou.billing_and_tariff_Setting(billing_mode=0, at_time='01 00:30:00', tariff=3)
    # 配置1个Schedule Id，Schedule ID对应配置1个segment ID
    tou.tou_schedules_add(schedule_id=1, segment_id=1, segment_tariff=segment_tariff, segment_time=segment_time)
    # 配置1个session,对应Schedule Id 1
    tou.tou_seasons_add(session_id=1, start_date=start_date, seasons_schedule_id=seasons_schedule_id)
    # 上位机上点击update
    update_success = tou.click_update()
    assert update_success


def test_013_03_case1():
    # Enable场景，第1和第2年，设置相同年份
    tou.open_holiday_setting_enable(holiday_start_date='2025', holiday_end_date='2024')
    # update保存设置成功
    update_success = tou.click_update()
    assert update_success

def teardown_function():
    # 恢复TOU设置，上位机上点击reset to default
    tou.tou_reset()
    # 上位机TOU ten_holidays 页面上点击ten_holidays_clear，恢复ten_holidays设置
    tou.ten_holidays_clear()
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
