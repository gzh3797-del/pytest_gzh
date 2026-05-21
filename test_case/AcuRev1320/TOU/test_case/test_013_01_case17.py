import logging
import time

from test_case.AcuRev1320.TOU.TOU_setting import TestAcuviewTou
from tools.log import Log

Log(str(__file__).split("\\")[-1])
"""
Function_AcuRev1320_013_01_case17
"""

dst_format = 2
(start_month, start_day, start_week, start_time, start_adjust_time) = (3, 1, 6, '02:00', 60)
(end_month, end_day, end_week, end_time, end_adjust_time) = (12, 5, 2, '23:59', 60)
time_year, time_month, time_day, time_hour, time_minute, time_second = 2025, 3, 7, 1, 59, 55
end_time_year, end_time_month, end_time_day, end_time_hour, end_time_minute, end_time_second = 2025, 12, 29, 23, 58, 55
tou = TestAcuviewTou()


def setup_function():
    # 打开上位机，连接登录设备
    tou.login()
    # 进入上位机设置general页面，确认设置密码
    tou.confirm_password()
    # 进入上位机DST_setting页面
    tou.enter_DST_setting()


def test_013_01_case17():
    # DST_enable使能打开
    tou.open_DST_enable()
    logging.info(f'上位机设置DST_enable:打开')
    # 上位机设置DST_format；0: format 1(fixed date) 1: format 2 (non fixed date)
    tou.DST_format(dst_format=dst_format)
    # 上位机设置format2_DST_time和adjust_time
    tou.setting_format2_DST_time(start_month=start_month, start_day=start_day, start_week=start_week,
                                 start_time=start_time, start_adjust_time=start_adjust_time,
                                 end_month=end_month, end_day=end_day, end_week=end_week,
                                 end_time=end_time, end_adjust_time=end_adjust_time)
    logging.info(f'上位机设置DTS format2 开始时间:{start_month}月、{start_day}周、星期{start_week - 1}, '
                 f'DTS开始时间{start_time}、DTS开始调整时间{start_adjust_time}'
                 f'上位机设置DTS format2 结束时间:{end_month}月、{end_day}周、星期{end_week - 1}, '
                 f'DTS结束时间{end_time}、DTS结束调整时间{end_adjust_time}'
                 )
    # 上位机点击update，检查是否更新成功
    update_success = tou.click_update()
    assert update_success

    # 设置电表系统时间
    tou.set_device_time(time_year, time_month, time_day, time_hour, time_minute, time_second)
    logging.info(f'设置电表时间:{time_year, time_month, time_day, time_hour, time_minute, time_second}')

    s = 6
    time.sleep(s)
    logging.info(f'电表等待时间{s}S')
    # 读取电表时间year, month, day, hour, minute, second
    (year, month, day, hour, minute, second) = tou.check_device_time()
    logging.info(f'读取电表时间:{year, month, day, hour, minute, second}')
    # 检查电表时间是否增加1小时
    if hour == time_hour + 2:
        logging.info(f'电表时间增加1小时')
        assert True
    else:
        assert False

    # 设置电表系统时间
    tou.set_device_time(end_time_year, end_time_month, end_time_day, end_time_hour, end_time_minute, end_time_second)
    logging.info(
        f'设置电表时间:{end_time_year, end_time_month, end_time_day, end_time_hour, end_time_minute, end_time_second}')
    s = 66
    time.sleep(s)
    logging.info(f'电表等待时间{s}S')
    (year, month, day, hour, minute, second) = tou.check_device_time()
    logging.info(f'读取电表时间:{year, month, day, hour, minute, second}')
    # 检查电表时间是否增加1小时
    if hour == end_time_hour:
        logging.info(f'电表时间减小1小时')
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
