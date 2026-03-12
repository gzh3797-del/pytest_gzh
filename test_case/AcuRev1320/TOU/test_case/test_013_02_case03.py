import logging

from test_case.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])
"""
    配置12个session(配置13个session ID提示失败),对应配置1个Schedule Id，对应配置1个segment ID，配置功能正常 
    
    1、连接AcuRev1320电表到Acuview2上，连接状态为connected
    2、上位机上配置14个Schedule Id，每个各配置14个segment ID((覆盖 尖sharp、峰peak、谷valley、平normal))
    3、添加12个Season， Season1~12依次按顺序关联步骤二配置Schedule ID1~12.
    4、上位机上点击"Update"。
    5、检查步骤2-3中的配置是否保存成功。 
    6、在步骤3的Season ID基础上，对应配置第13个Season ID
    7、检查上位机是否报错， 提示Season ID最大值12.
"""

tou = TestAcuviewTou()
segment_time = ['00:00', '01:00', '02:00', '03:00', '04:00', '05:00', '06:00',
                '07:00', '08:00', '09:00', '12:00', '13:00', '14:00', '15:00']
segment_tariff = [0, 1, 2, 3, 0, 1, 2, 3, 0, 1, 2, 3, 0, 1]
start_date = ['01-01', '02-01', '03-01', '04-01', '05-01', '06-01',
              '07-01', '08-01', '09-01', '10-01', '11-01', '12-01']
seasons_schedule_id = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()
    # 进入上位机tou_setting页面
    tou.enter_tou_setting()
    # TOU_enable使能打开
    tou.open_TOU_enable()


def test_013_02_case4():
    result = []
    # 上位机TOU Setting页面设置billing_and_tariff
    tou.billing_and_tariff_Setting(billing_mode=0, at_time='01 00:30:00', tariff=3)
    # 配置1个Schedule Id，Schedule ID对应配置1个segment ID
    tou.tou_schedules_add(schedule_id=1, segment_id=1, segment_tariff=segment_tariff, segment_time=segment_time)
    # 配置12个session,对应Schedule Id 都为1
    tou.tou_seasons_add(session_id=12, start_date=start_date, seasons_schedule_id=seasons_schedule_id)
    # 上位机上点击update
    update_success = tou.click_update()
    assert update_success
    # # 检查配置的segment_tariff是否正确
    res = tou.check_number_of_segments(number=1)
    assert res
    # 检查电表seasons个数是否为12
    seasons = tou.check_number_of_seasons(number=12)
    assert seasons
    check_seasons_add_13 = tou.tou_seasons_add_13()
    assert check_seasons_add_13


def teardown_function():
    # 恢复TOU设置，上位机上点击reset to default
    tou.tou_reset()
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
