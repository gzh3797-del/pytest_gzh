import logging

from test_case.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log
from test_case.AcuRev1320.TOU.page_elements_config import PageAddr

Log(str(__file__).split("\\")[-1])

"""
Function_AcuRev1320_013_02_case8
    创建segment ID時选择sharp、peak、valley、normal并同步成功，tariff栏去勾选任何一个选项，再次update均提示失败。
"""

tou = TestAcuviewTou()
segment_time = ['00:00', '01:00', '02:00', '03:00', '04:00', '05:00', '06:00',
                '07:00', '08:00', '09:00', '12:00', '13:00', '14:00', '15:00']
segment_tariff = [0, 1, 2, 3, 0, 1, 2, 3, 0, 1, 2, 3, 0, 1]
start_date = ['01-01', '02-01', '03-01', '04-01', '05-01', '06-01',
              '07-01', '08-01', '09-01', '10-01', '11-01', '12-01']
seasons_schedule_id = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]


def setup_function():
    # # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()
    # 进入上位机tou_setting页面
    tou.enter_tou_setting()
    # TOU_enable使能打开
    tou.open_TOU_enable()


def test_013_02_case7():
    # 上位机TOU Setting页面设置billing_and_tariff
    tou.billing_and_tariff_Setting(billing_mode=0, at_time='01 00:30:00', tariff=3)
    # 配置1个Schedule Id，Schedule ID对应配置1个segment ID
    tou.tou_schedules_add(schedule_id=1, segment_id=4, segment_tariff=segment_tariff, segment_time=segment_time)
    # 配置1个session,对应Schedule Id 1
    tou.tou_seasons_add(session_id=1, start_date=start_date, seasons_schedule_id=seasons_schedule_id)
    # 上位机上点击update
    update_success = tou.click_update()
    assert update_success
    # 上位机点击右侧进度条，返回TOU页面顶部
    tou.helper.click_pos((1710, 325))
    # 上位机上tariff去勾选“normal”，(1353, 468)为上位机TOU页面“normal”位置
    tou.helper.click_pos((1353, 468))
    # 检查上位机是否弹窗提示the tariff has been used
    the_tariff_has_been_used = tou.helper.check_image_exists(PageAddr.the_tariff_has_been_used)
    # 关闭弹窗
    tou.helper.click_image(PageAddr.yes)
    assert the_tariff_has_been_used


def teardown_function():
    # 恢复TOU设置，上位机上点击reset to default
    tou.tou_reset()
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
