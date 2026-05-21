import logging

from test_case.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

"""
    Special Weekday Schedule栏，配置添加1个不存在的season ，添加失败或update更新保存失败。
    
    1、连接AcuRev1320电表到Acuview2上，连接状态为connected。
    2、Special Weekday Schedule栏，配置添加1个不存在的season
    3、上位机上点击"Update"。
    4、检查步骤2-3中的配置是否保存成功。
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
    tou.billing_and_tariff_Setting(billing_mode=0, at_time='01 00:30:00', tariff=1)
    # 配置1个Schedule Id，Schedule ID对应配置1个segment ID
    tou.tou_schedules_add(schedule_id=1, segment_id=1)
    # 配置1个session,对应Schedule Id 1
    tou.tou_seasons_add(session_id=1, seasons_schedule_id=[1])
    # Special Weekday Schedule栏，已存在season1配置Schedule2(勾选Mon、Tue、Wed、Thu、Fri、Sat Sun，Sat）
    tou.tou_special_weekday_schedule_add(session_id=2,
                                         weekend_selection=[[0], [0]],
                                         weekend_schedule=[[1], [1]])
    # enable_special_weekday_schedule使能打开
    tou.open_enable_special_weekday_schedule()
    # Special Weekday Schedule栏，配置添加1个不存在的season，检查是否出现season ID undefined弹窗
    result = tou.check_season_ID_undefined()
    assert result


def teardown_function():
    # 恢复TOU设置，上位机上点击reset to default
    tou.tou_reset()
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
