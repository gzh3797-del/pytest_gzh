import logging
import time

from projects.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

"""
    Function_AcuRev1320_013_01_case1
    按照格式一设置DST时间(1.1 00:00 ~ 12.31 00:00)，开启DTS和关闭DST开关，DTS功能正常
"""

dst_format = 0
start_time = "01-01 00:00"
end_time = "12-31 00:00"
time_year, time_month, time_day, time_hour, time_minute, time_second = 2025, 3, 1, 0, 0, 0

tou = TestAcuviewTou()


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()
    # 进入上位机DST_setting页面
    tou.enter_DST_setting()


def test_013_01_case1():
    # DST_enable使能打开
    tou.open_DST_enable()
    logging.info(f'上位机设置DST_enable:打开')
    # 上位机设置DST_format；0: format 1(fixed date) 1: format 2 (non fixed date)
    tou.DST_format(dst_format=dst_format)
    # 上位机设置DTS开始时间 1月1日 00:00, DTS结束时间 12月31 00:00
    tou.setting_format1_DST_time(start_time=start_time, end_time=end_time)
    logging.info(f'上位机设置DTS开始时间{start_time}, DTS结束时间{end_time}')
    # 设置电表时间time_year, time_month, time_day, time_hour, time_minute, time_second
    tou.set_device_time(time_year, time_month, time_day, time_hour, time_minute, time_second)
    logging.info(f'设置电表时间:{time_year, time_month, time_day, time_hour, time_minute, time_second}')
    # 上位机点击update，检查是否更新成功
    update_success = tou.click_update()
    assert update_success
    # 读取电表时间year, month, day, hour, minute, second
    (year, month, day, hour, minute, second) = tou.check_device_time()
    logging.info(f'读取电表时间:{year, month, day, hour, minute, second}')
    # 检查电表时间是否增加1小时
    if hour == time_hour:
        logging.info(f'电表时间未增加1小时')
        assert True
    else:
        assert False
    # DST_enable使能关闭
    tou.close_DST_enable()
    # 上位机点击update，检查是否更新成功
    update_success = tou.click_update()
    assert update_success
    (year, month, day, hour, minute, second) = tou.check_device_time()
    # 检查电表时间是否增加1小时
    if hour == time_hour:
        assert True
    else:
        assert False


def teardown_function():
    # DST_enable使能关闭
    res = tou.close_DST_enable()
    if res:
        tou.click_update()
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
