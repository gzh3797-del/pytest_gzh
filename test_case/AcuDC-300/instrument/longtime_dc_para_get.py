import datetime
import time
from comm.modbus_set_attr import *
import xlwt
from tools.log import Log
import numpy as np
from PIL import ImageGrab, Image
import cv2
import time

Log(str(__file__).split("\\")[-1])

voltage = 60  # 输入电压可修改
current = -0.52  # 输入电流可修改
stime = 0.1# 测试时间可修改,单位为小时
enable_cable_loss_compensation = 0
cable_resistance = 0

my_workbook = xlwt.Workbook()
sheet = my_workbook.add_sheet('precision para')


def setup_function():
    sheet.write(0, 0, 'voltage precision')
    sheet.write(0, 1, 'current precision')
    sheet.write(0, 2, 'power precision')
    sheet.write(0, 3, 'import energy precision')
    sheet.write(0, 4, 'export energy precision')
    sheet.write(0, 5, 'net energy precision')
    sheet.write(0, 6, 'total energy precision')
    sheet.write(0, 7, 'import charge precision')
    sheet.write(0, 8, 'export charge precision')
    sheet.write(0, 9, 'net charge precision')
    sheet.write(0, 10, 'total_charge precision')
    # time.sleep(2)
    # assert set_cable_loss_compensation_enable(conn_mode=modbus_config["conn_mode"],
    #                                           value=enable_cable_loss_compensation) is True
    # # time.sleep(2)
    # assert set_cable_resistance(conn_mode=modbus_config["conn_mode"], value=cable_resistance) is True


def test_longtime_dc_para_get():
    result = {}
    conn_mode = 'tcp'
    i = 1
    # time.sleep(2)
    assert clear_energy(conn_mode=modbus_config["conn_mode"], value=1) is True
    # time.sleep(2)
    assert clear_charge(conn_mode=modbus_config["conn_mode"], value=1) is True
    # time.sleep(2)
    # assert clear_demand(conn_mode=modbus_config["conn_mode"], value=1) is True
    # time.sleep(2)
    cur_time = time.time()
    while time.time() <= cur_time + stime * 3600:
        # print(datetime)
        # while time.time() <= cur_time + 5:
        time.sleep(1)
        voltage_measu = read_voltage_measurement()
        current_measu = read_current_measurement()
        power_measu = read_power_measurement(conn_mode=conn_mode)
        # logging.info('voltage ret is:{}'.format(voltage_measu))
        # logging.info('current ret is:{}'.format(current_measu))
        # logging.info('power ret is:{}'.format(power_measu))
        voltage_precision = abs(abs(voltage_measu) - abs(voltage)) / abs(voltage)
        # print(voltage_measu, voltage)
        current_precision = abs(abs(current_measu) - abs(current)) / abs(current)
        power_precision = abs(abs(power_measu) - abs(voltage) * abs(current) / 1000) / (abs(voltage * current) / 1000)
        # print(power_measu, abs(voltage) * current / 1000)
        sheet.write(i, 0, voltage_precision)
        # # sheet.write(i, 1, current_measu)
        sheet.write(i, 1, current_precision)
        sheet.write(i, 2, power_precision)
        i += 1
    ret = read_energy_charge(conn_mode=modbus_config["conn_mode"])
    # time.sleep(2)
    # img = ImageGrab.grab(bbox=(0, 0, 1920, 1080))  # bbox 定义左、上、右和下像素的4元组
    # print(img.size[1], img.size[0])
    # img = np.array(img.getdata(), np.uint8).reshape(img.size[1], img.size[0], 3)
    # print(img)
    # img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)  # 看评论区有C友说颜色相反，于是加了这一条
    # cv2.imwrite('screenshot1.jpg', img)
    # end_time = time.time()
    # print(end_time - cur_time)

    # result['import_energy'] = ret[0]
    # result['export_energy'] = ret[1]
    # result['net_energy'] = ret[2]
    # result['total_energy'] = ret[3]
    # result['import_charge'] = ret[4]
    # result['export_charge'] = ret[5]
    # result['net_charge'] = ret[6]
    # result['total_charge'] = ret[7]
    logging.info('energy charge para is:{}'.format(ret))
    if current >= 0:
        # import_energy_precision = abs(ret[0] - voltage * current * stime / 1000) / (
        #         voltage * current * stime / 1000) * 100
        # export_energy_precision = 0
        # import_charge_precision = abs(ret[4] - current * stime) / (current * stime) * 100
        # export_charge_precision = 0
        # net_energy_precision = abs(ret[2] - (voltage * current * stime - 0) / 1000) / (
        #         (voltage * current * stime - 0) / 1000) * 100
        # total_energy_precision = abs(ret[3] - (voltage * current * stime + 0) / 1000) / (
        #         (voltage * current * stime + 0) / 1000) * 100
        # net_charge_precision = abs(ret[6] - (current * stime - 0)) / (current * stime - 0) * 100
        # total_charge_precision = abs(ret[7] - (current * stime + 0)) / (
        #         current * stime + 0) * 100
        # 测量值
        import_energy_precision = ret[0]
        export_energy_precision = 0
        import_charge_precision = ret[4]
        export_charge_precision = 0
        net_energy_precision = ret[2]
        total_energy_precision = ret[3]
        net_charge_precision = ret[6]
        total_charge_precision = ret[7]
    else:
        # import_energy_precision = 0
        # export_energy_precision = abs(ret[1] - voltage * abs(current) * stime / 1000) / (
        #         voltage * abs(current) * stime / 1000) * 100
        # import_charge_precision = 0
        # export_charge_precision = abs(ret[5] - abs(current) * stime) / (abs(current) * stime) * 100
        # net_energy_precision = abs(abs(ret[2]) - abs((0 - voltage * abs(current) * stime)) / 1000) / (
        #         abs((0 - voltage * abs(current) * stime)) / 1000) * 100
        # total_energy_precision = abs(ret[3] - (voltage * abs(current) * stime + 0) / 1000) / (
        #         (voltage * abs(current) * stime + 0) / 1000) * 100
        # net_charge_precision = abs(abs(ret[6]) - abs((0 - abs(current) * stime))) / abs(
        #     (0 - abs(current) * stime)) * 100
        # total_charge_precision = abs(ret[7] - (abs(current) * stime + 0)) / (
        #         abs(current) * stime + 0) * 100
        # 测量值
        import_energy_precision = 0
        export_energy_precision = ret[1]
        import_charge_precision = 0
        export_charge_precision = ret[5]
        net_energy_precision = ret[2]
        total_energy_precision = ret[3]
        net_charge_precision = ret[6]
        total_charge_precision = ret[7]
    sheet.write(1, 3, import_energy_precision)
    sheet.write(1, 4, export_energy_precision)
    sheet.write(1, 5, net_energy_precision)
    sheet.write(1, 6, total_energy_precision)
    sheet.write(1, 7, import_charge_precision)
    sheet.write(1, 8, export_charge_precision)
    sheet.write(1, 9, net_charge_precision)
    sheet.write(1, 10, total_charge_precision)
    my_workbook.save('{}.xlsx'.format(str(__file__).split("\\")[-1].split('.')[0]))


def teardown_function():
    pass
