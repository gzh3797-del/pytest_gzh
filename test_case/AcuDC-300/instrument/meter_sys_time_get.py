"""
test_name:
test_num:
author:
modify:
"""
import datetime
import time
import xlwt
from comm.modbus_get_attr import get_real_time_clock
from comm.multi_threads import multi_threads_request
from tools.log import Log
import logging
from modbus_config import modbus_config

Log(str(__file__).split("\\")[-1])

stime = 0.005  # 单位小时


def get_acudc300_time(time_dict):
    dc300_time = get_real_time_clock(conn_mode=modbus_config["conn_mode"])
    logging.info('AcuDC time is:{}'.format(dc300_time))
    time_dict['acudc300_time'] = dc300_time


def get_sys_time(sys_time):
    cur_time = datetime.datetime.now()
    now = cur_time.strftime("%Y-%m-%d %H:%M:%S")
    logging.info('local time is:{}'.format(now))
    sys_time['local_time'] = now


def setup_function():
    pass


def test_meter_sys_time_get():
    test_data = {}
    test_data_l = []
    threads_dict = [[get_acudc300_time, (test_data,)],
                    [get_sys_time, (test_data,)]]
    i = 1

    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('AcuDC300 SYS TIME')
    sheet.write(0, 0, 'AcuDC300 time')
    sheet.write(0, 1, 'local time')
    sheet.write(0, 2, 'compare')
    cur_time = time.time()
    print(cur_time)
    while time.time() <= cur_time + stime * 3600:
        multi_threads_request(threads_dict)
        acudc300 = test_data['acudc300_time'].split(':')[-1]
        test_data_l.append(float(acudc300))
        sys = test_data['local_time'].split(':')[-1]
        test_data_l.append(float(sys))
        logging.info('AcuDC300 time is:{},localtime is:{}'.format(test_data['acudc300_time'], test_data['local_time']))
        sheet.write(i, 0, test_data['acudc300_time'])
        sheet.write(i, 1, test_data['local_time'])
        sheet.write(i, 2, abs(test_data_l[0] - test_data_l[1]) <= 2)
        i += 1
    my_workbook.save('{}.xlsx'.format(str(__file__).split("\\")[-1].split('.')[0]))


def teardown_function():
    pass
