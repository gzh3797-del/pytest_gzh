from comm.source_control import *
from comm.modbus_set_attr import set_energy_pulse_parameter, set_energy_pulse_constant
from tools.excel_operate import data_read
from tools.log import Log
import xlwt
import math


def get_pulse_accuracy(vol, cur, times=0.2, pluse_mode=0):
    power = vol * cur / 1000
    pluse_cons = math.floor(times * 3600 / power)
    sour_para_conf(input_method='直接', pluse_cons=pluse_cons)
    time.sleep(10)
    sour_output(voltage=vol, current=cur)
    set_energy_pulse_parameter(conn_mode='tcp', value=pluse_mode)
    set_energy_pulse_constant(conn_mode='tcp', value=pluse_cons)
    timeout = 2 / times
    timeout *= 60
    ret = get_verification_error(timeout=timeout)
    return ret, pluse_cons


def run():
    volt_cur_list = data_read(r'../../../comm/test_data/dc_data.xlsx', 'Sheet1')
    Log(str(__file__).split("\\")[-1])
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('energy pluse test', cell_overwrite_ok=True)
    sheet.write(0, 0, '电压输入值')
    sheet.write(0, 1, '电流输入值')
    sheet.write(0, 2, '脉冲常数设置值')
    sheet.write(0, 3, '脉冲精度')
    for i in range(len(volt_cur_list)):
        sheet.write(i + 1, 0, float(volt_cur_list[i][0]))
        sheet.write(i + 1, 1, float(volt_cur_list[i][1]))
        ret = get_pulse_accuracy(vol=float(volt_cur_list[i][0]), cur=float(volt_cur_list[i][1]), pluse_mode=1)
        sheet.write(i + 1, 2, ret[1])
        sheet.write(i + 1, 3, ret[0])
    sour_stop()
    my_workbook.save('energy_pluse_test_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


run()
