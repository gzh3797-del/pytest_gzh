from comm.source_control import *
from tools.excel_operate import dcpara_addr_get
import numpy as np
from tools.log import Log
import xlwt
from comm.modbus_rtu_tcp import ModbusRtuOrTcp
import logging
import struct

dc_para_addr = dcpara_addr_get(r'/comm/test_data/AcuDC300.xlsx', 'Readings')
Log(str(__file__).split("\\")[-1])
my_workbook = xlwt.Workbook()
sheet = my_workbook.add_sheet('precision para')
mes = ModbusRtuOrTcp(conn_mode='rtu')
vole = 1000
curr = 600
stime = 1/3


def read_vol(standard_value):
    vol_list = []
    for v in range(10):
        time.sleep(1)
        voltages = mes.read_measurement(address=dc_para_addr['V(Measured) float32']['Start(Dec)'],
                                        count=dc_para_addr['V(Measured) float32']['Reg'], slave=1)
        logging.info('voltage ret is:{}'.format(voltages))
        reg = hex(voltages[0]).replace('0x', '').zfill(4) + hex(voltages[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        voltage_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        vol_list.append(voltage_measu)
    vol_list.sort()
    # print('vol:{}'.format(vol_list))
    mean = np.mean(vol_list)
    var = np.var(vol_list)
    if abs(vol_list[-1] - standard_value) > abs(vol_list[0] - standard_value):
        return vol_list[-1], mean, var
    return vol_list[0], mean, var


def read_cur(standard_value):
    vol_list = []
    for v in range(10):
        time.sleep(1)
        voltages = mes.read_measurement(address=dc_para_addr['Current float32']['Start(Dec)'],
                                        count=dc_para_addr['Current float32']['Reg'], slave=1)
        logging.info('voltage ret is:{}'.format(voltages))
        reg = hex(voltages[0]).replace('0x', '').zfill(4) + hex(voltages[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        voltage_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        vol_list.append(voltage_measu)
    vol_list.sort()
    # print('cur:{}'.format(vol_list))
    mean = np.mean(vol_list)
    var = np.var(vol_list)
    if abs(vol_list[-1] - standard_value) > abs(vol_list[0] - standard_value):
        return vol_list[-1], mean, var
    return vol_list[0], mean, var


def read_pow(standard_value):
    vol_list = []
    for v in range(10):
        time.sleep(1)
        voltages = mes.read_measurement(address=dc_para_addr['Power float32']['Start(Dec)'],
                                        count=dc_para_addr['Power float32']['Reg'], slave=1)
        logging.info('voltage ret is:{}'.format(voltages))
        reg = hex(voltages[0]).replace('0x', '').zfill(4) + hex(voltages[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        voltage_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        vol_list.append(voltage_measu)
    vol_list.sort()
    # print('pow:{}'.format(vol_list))
    mean = np.mean(vol_list)
    var = np.var(vol_list)
    if abs(vol_list[-1] - standard_value) > abs(vol_list[0] - standard_value):
        return vol_list[-1], mean, var
    return vol_list[0], mean, var


sheet.write(0, 0, 'voltage')
sheet.write(0, 1, 'voltage measu')
sheet.write(0, 2, 'voltage precision')
sheet.write(0, 3, 'voltage average')
sheet.write(0, 4, 'voltage variance')
sheet.write(0, 5, 'current')
sheet.write(0, 6, 'current measu')
sheet.write(0, 7, 'current precision')
sheet.write(0, 8, 'current average')
sheet.write(0, 9, 'current variance')
sheet.write(0, 10, 'power')
sheet.write(0, 11, 'power measu')
sheet.write(0, 12, 'power precision')
sheet.write(0, 13, 'power average')
sheet.write(0, 14, 'power variance')

j = 1
for i in range(1):
    cur_time = time.time()
    sour_output(voltage=vole, current=curr)
    while time.time() <= cur_time + stime * 3600:
        time.sleep(25)
        vol = read_vol(vole)
        cur = read_cur(curr)
        power = vole * curr / 1000
        ret = read_pow(power)
        sheet.write(j, 0, vole)
        sheet.write(j, 1, abs(vol[0]))
        sheet.write(j, 2, abs(abs(vole - abs(vol[0])) / abs(abs(vole)) * 100))
        sheet.write(j, 3, vol[1])
        sheet.write(j, 4, vol[2])
        sheet.write(j, 5, curr)
        sheet.write(j, 6, abs(cur[0]))
        sheet.write(j, 7, abs(abs(curr) - abs(cur[0])) / abs(abs(curr)) * 100)
        sheet.write(j, 8, cur[1])
        sheet.write(j, 9, cur[2])
        sheet.write(j, 10, power)
        sheet.write(j, 11, abs(ret[0]))
        sheet.write(j, 12, abs(abs(power) - abs(ret[0])) / abs(power) * 100)
        sheet.write(j, 13, ret[1])
        sheet.write(j, 14, ret[2])
        print((time.time() - cur_time) / 60)
        j += 1
    time.sleep(2)

sour_stop()
mes.close()
my_workbook.save(
    '{}_{}.xls'.format(str(__file__).split("\\")[-1].split('.')[0], time.strftime('%Y%m%d%H%M%S')))
