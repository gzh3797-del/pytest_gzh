import logging

from test_case.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

"""
    Function_AcuRev1320_013_02_case1
    
    配置1个session, 对应配置1个Schedule Id，对应配置14个segment ID(配置15个segment ID 提示失败)，
    tariff覆盖(尖sharp、峰peak、谷valley、平normal)，配置功能正常
    
    1、连接AcuRev1320电表到Acuview2上，连接状态为connected
    2、上位机上配置一个Schedule ID，对应配置14个segment ID((覆盖 尖sharp、峰peak、谷valley、平normal))
    3、添加一个Season， 并关联步骤二配置Schedule ID.
    4、上位机上点击"Update"。
    5、检查步骤2-3中的配置是否保存成功。 
    6、在步骤2的Schedule ID上，对应配置第15个segment ID
    7、检查上位机是否报错， 提示segment ID最大值14.

"""

tou = TestAcuviewTou()
segment_time = ['00:00', '01:00', '02:00', '03:00', '04:00', '05:00', '06:00',
                '07:00', '08:00', '09:00', '12:00', '13:00', '14:00', '15:00']
segment_tariff = [0, 1, 2, 3, 0, 1, 2, 3, 0, 1, 2, 3, 0, 1]


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()
    # 进入上位机tou_setting页面
    tou.enter_tou_setting()
    # TOU_enable使能打开
    tou.open_TOU_enable()


def test_013_02_case2():
    result = []
    # 上位机TOU Setting页面设置billing_and_tariff
    tou.billing_and_tariff_Setting(billing_mode=0, at_time='01 00:30:00', tariff=3)
    # 配置1个Schedule Id，Schedule ID对应配置14个segment ID
    tou.tou_schedules_add(schedule_id=1, segment_id=14, segment_tariff=segment_tariff, segment_time=segment_time)
    # 配置1个session,对应Schedule Id 1
    tou.tou_seasons_add(session_id=1, seasons_schedule_id=[1])
    # 上位机上点击update
    update_success = tou.click_update()
    result.append(update_success)
    # 检查配置的segment_tariff是否正确
    res = tou.check_number_of_segments(number=14)
    result.append(res)
    re = tou.tou_schedules_add_15()
    result.append(re)
    # 上位机上点击reset to default
    tou.tou_reset()
    assert result == [True, True, True]


def teardown_function():
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
