import logging

from test_case.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

"""
    Special Weekday Schedule栏，已存在season ID1配置Schedule ID1(勾选Mon、Tue、Wed、Thu、Fri、Sat Sun，Sat）可配置成功，
    特殊周休配置开关开启和关闭正常
    
    1、连接AcuRev1320电表到Acuview2上，连接状态为connected。
    2、Special Weekday Schedule栏，已存在season1配置Schedule1(勾选Mon、Tue、Wed、Thu、Fri、Sat Sun，Sat）
    3、打开“Enable Special Weekday Schedule”
    4、上位机上点击"Update"。
    5、检查步骤2-3中的配置是否保存成功。
    6、关闭“Enable Special Weekday Schedule”
    7、上位机上点击"Update"。
    8、检查步骤2-3中的配置是否保存成功。
"""

tou = TestAcuviewTou()


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()
    # 进入上位机tou_setting页面
    tou.enter_tou_setting()
    # TOU_enable使能打开
    tou.open_TOU_enable()


def test_013_03_case1():
    # 上位机TOU Setting页面设置billing_and_tariff
    tou.billing_and_tariff_Setting(billing_mode=0, at_time='01 00:30:00', tariff=3)
    # 配置1个Schedule Id，Schedule ID对应配置1个segment ID
    tou.tou_schedules_add(schedule_id=1, segment_id=1)
    # 配置1个session,对应Schedule Id 1
    tou.tou_seasons_add(session_id=1, seasons_schedule_id=[1])
    # Special Weekday Schedule栏，已存在season1配置Schedule1(勾选Mon、Tue、Wed、Thu、Fri、Sat Sun，Sat）
    tou.tou_special_weekday_schedule_add(session_id=1,
                                         weekend_selection=[[0, 1, 2, 3, 4, 5, 6]],
                                         weekend_schedule=[[1, 1, 1, 1, 1, 1, 1]])
    # enable_special_weekday_schedule使能打开
    tou.open_enable_special_weekday_schedule()
    # 上位机上点击update
    update_success = tou.click_update()
    assert update_success
    # enable_special_weekday_schedule使能关闭
    tou.close_enable_special_weekday_schedule()
    update_success = tou.click_update()
    assert update_success


def teardown_function():
    # 恢复TOU设置，上位机上点击reset to default
    tou.tou_reset()
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
