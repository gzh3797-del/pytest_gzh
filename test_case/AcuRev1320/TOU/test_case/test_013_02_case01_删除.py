import logging

from test_case.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

tou = TestAcuviewTou()
segment_tariff = [0, 1, 2, 3]


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()
    # 进入上位机tou_setting页面
    tou.enter_tou_setting()
    # TOU_enable使能打开
    tou.open_TOU_enable()


def test_013_02_case1():
    result = []
    for tariff in segment_tariff:
        # 上位机TOU Setting页面设置billing_and_tariff
        tou.billing_and_tariff_Setting(billing_mode=0, at_time='01 00:30:00', tariff=3)
        # 配置1个Schedule Id，Schedule ID对应配置1个segment ID
        tou.tou_schedules_add(schedule_id=1, segment_id=1, segment_tariff=[tariff])
        # 配置1个session,对应Schedule Id 1
        tou.tou_seasons_add(session_id=1, seasons_schedule_id=[1])
        # 上位机上点击update
        tou.click_update()
        # 检查配置的segment_tariff是否正确
        res = tou.check_segment_1_setting_tariffs(tariff=tariff)
        # 上位机上点击reset to default
        result.append(res)
        # 恢复TOU设置，上位机上点击reset to default
        tou.tou_reset()
    assert result == [True, True, True, True]


def teardown_function():
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
