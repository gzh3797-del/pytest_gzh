from comm.modbus_get_attr import *
from comm.source_control import *
from tools.log import Log
import xlwt
from tools.excel_operate import data_read, dcpara_addr_get
import numpy as np

dc_para_addr = dcpara_addr_get(r'/comm/test_data/AcuDC300.xlsx', 'Readings')
volt_cur_list = data_read(r'/comm/test_data/dc_data.xlsx', 'Sheet1')
Log(str(__file__).split("\\")[-1])
my_workbook = xlwt.Workbook()
sheet = my_workbook.add_sheet('precision para')
mes = ModbusRtuOrTcp(conn_mode='rtu')


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

test_num = len(volt_cur_list)
for i in range(test_num):
    mv_sour_output(voltage=volt_cur_list[i][0], current=0)
    ret = read_vol(volt_cur_list[i][0])

    sheet.write(i + 1, 0, volt_cur_list[i][0])
    sheet.write(i + 1, 1, abs(ret[0]))
    sheet.write(i + 1, 2, abs(abs(volt_cur_list[i][0]) - abs(ret[0])) / abs(abs(volt_cur_list[i][0])) * 100)
    sheet.write(i + 1, 3, ret[1])
    sheet.write(i + 1, 4, ret[2])
    sheet.write(i + 1, 5, 'null')
    sheet.write(i + 1, 6, 'null')
    sheet.write(i + 1, 7, 'null')
    sheet.write(i + 1, 8, 'null')
    sheet.write(i + 1, 9, 'null')
    sheet.write(i + 1, 10, 'null')
    sheet.write(i + 1, 11, 'null')
    sheet.write(i + 1, 12, 'null')
    sheet.write(i + 1, 13, 'null')
    sheet.write(i + 1, 14, 'null')
for i in range(test_num):
    mv_sour_output(voltage=0, current=volt_cur_list[i][1])
    ret = read_cur(volt_cur_list[i][1])
    sheet.write(test_num + i + 1, 0, 'null')
    sheet.write(test_num + i + 1, 1, 'null')
    sheet.write(test_num + i + 1, 2, 'null')
    sheet.write(test_num + i + 1, 3, 'null')
    sheet.write(test_num + i + 1, 4, 'null')
    sheet.write(test_num + i + 1, 5, volt_cur_list[i][1])
    sheet.write(test_num + i + 1, 6, abs(ret[0]))
    sheet.write(test_num + i + 1, 7, abs(abs(volt_cur_list[i][1]) - abs(ret[0])) / abs(abs(volt_cur_list[i][1])) * 100)
    sheet.write(test_num + i + 1, 8, ret[1])
    sheet.write(test_num + i + 1, 9, ret[2])
    sheet.write(test_num + i + 1, 10, 'null')
    sheet.write(test_num + i + 1, 11, 'null')
    sheet.write(test_num + i + 1, 12, 'null')
    sheet.write(test_num + i + 1, 13, 'null')
    sheet.write(test_num + i + 1, 14, 'null')
for i in range(test_num):
    sour_output(voltage=volt_cur_list[i][0], current=volt_cur_list[i][1])
    vol = read_vol(volt_cur_list[i][0])
    cur = read_cur(volt_cur_list[i][1])
    power = volt_cur_list[i][0] * volt_cur_list[i][1] / 1000
    ret = read_pow(power)
    sheet.write(test_num * 2 + i + 1, 0, volt_cur_list[i][0])
    sheet.write(test_num * 2 + i + 1, 1, abs(vol[0]))
    sheet.write(test_num * 2 + i + 1, 2,
                abs(abs(volt_cur_list[i][0]) - abs(vol[0])) / abs(abs(volt_cur_list[i][0])) * 100)
    sheet.write(test_num * 2 + i + 1, 3, vol[1])
    sheet.write(test_num * 2 + i + 1, 4, vol[2])
    sheet.write(test_num * 2 + i + 1, 5, volt_cur_list[i][1])
    sheet.write(test_num * 2 + i + 1, 6, abs(cur[0]))
    sheet.write(test_num * 2 + i + 1, 7,
                abs(abs(volt_cur_list[i][1]) - abs(cur[0])) / abs(abs(volt_cur_list[i][1])) * 100)
    sheet.write(test_num * 2 + i + 1, 8, cur[1])
    sheet.write(test_num * 2 + i + 1, 9, cur[2])
    sheet.write(test_num * 2 + i + 1, 10, power)
    sheet.write(test_num * 2 + i + 1, 11, abs(ret[0]))
    sheet.write(test_num * 2 + i + 1, 12, abs(abs(power) - abs(ret[0])) / abs(power) * 100)
    sheet.write(test_num * 2 + i + 1, 13, ret[1])
    sheet.write(test_num * 2 + i + 1, 14, ret[2])
sour_stop()
mes.close()
my_workbook.save(
    '{}_{}.xls'.format(str(__file__).split("\\")[-1].split('.')[0], time.strftime('%Y%m%d%H%M%S')))
