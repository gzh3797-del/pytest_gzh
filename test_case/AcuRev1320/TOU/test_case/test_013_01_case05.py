import logging
import time

from test_case.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])

"""
Function_AcuRev1320_013_01_case5
"""

dst_format = 0
start_time = "03-01 02:00"
end_time = "03-01 02:00"
start_year, start_month, start_day, start_hour, start_minute, start_second = 2024, 3, 1, 1, 59, 55
end_year, end_month, end_day, end_hour, end_minute, end_second = 2024, 10, 31, 2, 59, 55
tou = TestAcuviewTou()


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()
    # 进入上位机DST_setting页面
    tou.enter_DST_setting()


def test_013_01_case5():
    # DST_enable使能打开
    tou.open_DST_enable()
    logging.info(f'上位机设置DST_enable:打开')
    # 上位机设置DST_format；0: format 1(fixed date) 1: format 2 (non fixed date)
    tou.DST_format(dst_format=dst_format)
    # 上位机设置DTS开始时间 1月1日 00:00, DTS结束时间 12月31 00:00
    tou.setting_format1_DST_time(start_time=start_time, end_time=end_time)
    logging.info(f'上位机设置DTS开始时间{start_time}, DTS结束时间{end_time}')
    # 上位机检查是否弹窗format_adjust_time_invalid
    res = tou.check_format_adjust_time_invalid()
    assert res


def teardown_function():
    # DST_enable使能关闭
    res = tou.close_DST_enable()
    if res:
        tou.click_update()
    # 关闭所有Acuview相关进程
    tou.helper.kill_acuview_apps()
    # 关闭modbus_client
    tou.handle_memory.modbus_client.close()
