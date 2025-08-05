#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# @File     :1.py
# @Author   :lcs
# @Time     :2025/8/5
# @Desc     :

import time
from AcuRev4100_modbus_get import *
import threading
import tkinter as tk
from tkinter import messagebox

# volt_cur_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'test_data')
Log(str(__file__).split("\\")[-1])


# ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])


def frequency_precision_measure():
    frequency_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'frequency')
    print(frequency_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('frequency_precision_measure', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, '频率输入值')
    sheet.write(0, 2, '电表实际频率')
    sheet.write(0, 3, '频率精度')
    sheet.write(0, 4, '测试结果')
    for i in range(len(frequency_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(frequency_list[i]))
            print('测试进度:{}'.format(frequency_list[i]))
        else:
            logging.info('测试进度:{},执行时间:{}'.format(frequency_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(frequency_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(frequency_list) - 1:
            break
        sheet.write(i + 1, 0, frequency_list[i + 1][0])
        sheet.write(i + 1, 1, frequency_list[i + 1][1])
        if frequency_list[i + 1][1] != 'null' and frequency_list[i + 1][2] != 'null':
            ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 1, 1, 1, frequency_list[i + 1][1])
            frequency = read_frequency(frequency_list[i + 1][1], times=40)
            scale = abs(frequency - frequency_list[i + 1][1]) / frequency_list[i + 1][1]
            sheet.write(i + 1, 2, frequency)
            sheet.write(i + 1, 3, f'{scale:.3%}')
            if scale * 100 <= frequency_list[i + 1][2]:
                sheet.write(i + 1, 4, 'Passed')
            else:
                sheet.write(i + 1, 4, 'Failed')
        if frequency_list[i + 1][1] != 'null' and frequency_list[i + 1][2] == 'null':
            ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 1, 1, 1, frequency_list[i + 1][1])
            frequency = read_frequency(frequency_list[i + 1][1], times=40)
            sheet.write(i + 1, 2, frequency)
            if frequency != 0:
                sheet.write(i + 1, 4, 'Passed')
            else:
                sheet.write(i + 1, 4, 'Failed')
    # mes.close()
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def line_to_neutral_voltage_precision_measure():
    voltage_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Line_to_Neutral_Voltage')
    print(voltage_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('line_to_neutral_voltage', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Van输入值')
    sheet.write(0, 2, 'Vbn输入值')
    sheet.write(0, 3, 'Vcn输入值')
    sheet.write(0, 4, 'Va电压角度输入值')
    sheet.write(0, 5, 'Vb电压角度输入值')
    sheet.write(0, 6, 'Vc电压角度输入值')
    sheet.write(0, 7, '电表实际Van值')
    sheet.write(0, 8, '电表实际Vbn值')
    sheet.write(0, 9, '电表实际Vcn值')
    sheet.write(0, 10, '电表实际Vlnavg值')
    sheet.write(0, 11, 'Van精度')
    sheet.write(0, 12, 'Vbn精度')
    sheet.write(0, 13, 'Vcn精度')
    sheet.write(0, 14, 'Vlnavg精度')
    sheet.write(0, 15, '测试结果')
    for i in range(len(voltage_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(voltage_list[i]))
            print('测试进度:{}'.format(voltage_list[i]))
        else:
            logging.info('测试进度:{},执行时间:{}'.format(voltage_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(voltage_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(voltage_list) - 1:
            break
        sheet.write(i + 1, 0, voltage_list[i + 1][0])
        sheet.write(i + 1, 1, voltage_list[i + 1][1])
        sheet.write(i + 1, 2, voltage_list[i + 1][2])
        sheet.write(i + 1, 3, voltage_list[i + 1][3])
        sheet.write(i + 1, 4, voltage_list[i + 1][4])
        sheet.write(i + 1, 5, voltage_list[i + 1][5])
        sheet.write(i + 1, 6, voltage_list[i + 1][6])
        if voltage_list[i + 1][7] != 'null' and voltage_list[i + 1][1] != 'null' and voltage_list[i + 1][
            2] != 'null' and voltage_list[i + 1][3] != 'null' and voltage_list[i + 1][8] == 'null':
            ret = set_ac(120, 240, 0, 120, 240, 0, voltage_list[i + 1][3], voltage_list[i + 1][2],
                         voltage_list[i + 1][1], 1, 1, 1, 50)
            Average_Vol = (voltage_list[i + 1][1] + voltage_list[i + 1][2] + voltage_list[i + 1][3]) / 3
            Phase_A = read_phase_a_voltage(voltage_list[i + 1][1], times=40)
            Phase_B = read_phase_b_voltage(voltage_list[i + 1][2], times=40)
            Phase_C = read_phase_c_voltage(voltage_list[i + 1][3], times=40)
            Read_Average_Vol = read_average_ln_voltage(Average_Vol, times=40)
            scale_A = abs(Phase_A - voltage_list[i + 1][1]) / voltage_list[i + 1][1]
            scale_B = abs(Phase_B - voltage_list[i + 1][2]) / voltage_list[i + 1][2]
            scale_C = abs(Phase_C - voltage_list[i + 1][3]) / voltage_list[i + 1][3]
            scale_Average_Vol = abs(Read_Average_Vol - Average_Vol) / Average_Vol
            sheet.write(i + 1, 7, Phase_A)
            sheet.write(i + 1, 8, Phase_B)
            sheet.write(i + 1, 9, Phase_C)
            sheet.write(i + 1, 10, Read_Average_Vol)
            sheet.write(i + 1, 11, f'{scale_A:.2%}')
            sheet.write(i + 1, 12, f'{scale_B:.2%}')
            sheet.write(i + 1, 13, f'{scale_C:.2%}')
            sheet.write(i + 1, 14, f'{scale_Average_Vol:.2%}')
            if scale_A * 100 <= voltage_list[i + 1][7] and scale_B * 100 <= voltage_list[i + 1][7] and scale_C * 100 <= \
                    voltage_list[i + 1][7]:
                sheet.write(i + 1, 15, 'Passed')
            else:
                sheet.write(i + 1, 15, 'Failed')
        if voltage_list[i + 1][7] == 'null' and voltage_list[i + 1][1] != 'null' and voltage_list[i + 1][
            2] != 'null' and voltage_list[i + 1][3] != 'null' and voltage_list[i + 1][8] == 'null':
            ret = set_ac(120, 240, 0, 120, 240, 0, voltage_list[i + 1][3], voltage_list[i + 1][2],
                         voltage_list[i + 1][1], 1, 1, 1, 50)
            Average_Vol = (voltage_list[i + 1][1] + voltage_list[i + 1][2] + voltage_list[i + 1][3]) / 3
            Phase_A = read_phase_a_voltage(voltage_list[i + 1][1], times=40)
            Phase_B = read_phase_b_voltage(voltage_list[i + 1][2], times=40)
            Phase_C = read_phase_c_voltage(voltage_list[i + 1][3], times=40)
            Read_Average_Vol = read_average_ln_voltage(Average_Vol, times=40)
            sheet.write(i + 1, 7, Phase_A)
            sheet.write(i + 1, 8, Phase_B)
            sheet.write(i + 1, 9, Phase_C)
            sheet.write(i + 1, 10, Read_Average_Vol)
            sheet.write(i + 1, 11, 'null')
            sheet.write(i + 1, 12, 'null')
            sheet.write(i + 1, 13, 'null')
            sheet.write(i + 1, 14, 'null')
            if voltage_list[i + 1][1] < 10 and voltage_list[i + 1][2] < 10 and voltage_list[i + 1][3] < 10:
                if Phase_A == Phase_B == Phase_C == Read_Average_Vol == 0:
                    sheet.write(i + 1, 15, 'Passed')
                    print('Passed')
                else:
                    sheet.write(i + 1, 15, 'Failed')
            else:
                if Phase_A != 0 and Phase_B != 0 and Phase_C != 0 and Read_Average_Vol != 0:
                    sheet.write(i + 1, 15, 'Passed')
                    print('Passed')
                else:
                    sheet.write(i + 1, 15, 'Failed')

        if voltage_list[i + 1][7] != 'null' and voltage_list[i + 1][1] != 'null' and voltage_list[i + 1][
            2] != 'null' and voltage_list[i + 1][3] != 'null' and voltage_list[i + 1][8] != 'null':
            if voltage_list[i + 1][8] == 'ABC':
                set_phase_order(0)
                sheet.write(i + 1, 16, '相序设置：ABC')
            if voltage_list[i + 1][8] == 'ACB':
                set_phase_order(1)
                sheet.write(i + 1, 16, '相序设置：ACB')
            ret = set_ac(voltage_list[i + 1][6], voltage_list[i + 1][5], voltage_list[i + 1][4], 120, 240, 0,
                         voltage_list[i + 1][3], voltage_list[i + 1][2], voltage_list[i + 1][1], 1, 1, 1, 50)
            Average_Vol = (voltage_list[i + 1][1] + voltage_list[i + 1][2] + voltage_list[i + 1][3]) / 3
            Phase_A = read_phase_a_voltage(voltage_list[i + 1][1], times=40)
            Phase_B = read_phase_b_voltage(voltage_list[i + 1][2], times=40)
            Phase_C = read_phase_c_voltage(voltage_list[i + 1][3], times=40)
            Read_Average_Vol = read_average_ln_voltage(Average_Vol, times=40)
            scale_A = abs(Phase_A - voltage_list[i + 1][1]) / voltage_list[i + 1][1]
            scale_B = abs(Phase_B - voltage_list[i + 1][2]) / voltage_list[i + 1][2]
            scale_C = abs(Phase_C - voltage_list[i + 1][3]) / voltage_list[i + 1][3]
            scale_Average_Vol = abs(Read_Average_Vol - Average_Vol) / Average_Vol
            sheet.write(i + 1, 7, Phase_A)
            sheet.write(i + 1, 8, Phase_B)
            sheet.write(i + 1, 9, Phase_C)
            sheet.write(i + 1, 10, Read_Average_Vol)
            sheet.write(i + 1, 11, f'{scale_A:.2%}')
            sheet.write(i + 1, 12, f'{scale_B:.2%}')
            sheet.write(i + 1, 13, f'{scale_C:.2%}')
            sheet.write(i + 1, 14, f'{scale_Average_Vol:.2%}')
            if scale_A * 100 <= voltage_list[i + 1][7] and scale_B * 100 <= voltage_list[i + 1][7] and scale_C * 100 <= \
                    voltage_list[i + 1][7]:
                sheet.write(i + 1, 15, 'Passed')
            else:
                sheet.write(i + 1, 15, 'Failed')
            time.sleep(0.2)
            set_phase_order(0)
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def line_to_line_voltage_precision_measure():
    voltage_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Line_to_Line_Voltage')
    print(voltage_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('line_to_line_voltage', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Van输入值')
    sheet.write(0, 2, 'Vbn输入值')
    sheet.write(0, 3, 'Vcn输入值')
    sheet.write(0, 4, 'Va电压角度输入值')
    sheet.write(0, 5, 'Vb电压角度输入值')
    sheet.write(0, 6, 'Vc电压角度输入值')
    sheet.write(0, 7, '电表实际Vab值')
    sheet.write(0, 8, '电表实际Vbc值')
    sheet.write(0, 9, '电表实际Vca值')
    sheet.write(0, 10, '电表实际Vllavg值')
    sheet.write(0, 11, 'Vab精度')
    sheet.write(0, 12, 'Vbc精度')
    sheet.write(0, 13, 'Vca精度')
    sheet.write(0, 14, 'Vllavg精度')
    sheet.write(0, 15, '测试结果')
    for i in range(len(voltage_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(voltage_list[i]))
            print('测试进度:{}'.format(voltage_list[i]))
        else:
            logging.info('测试进度:{},执行时间:{}'.format(voltage_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(voltage_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(voltage_list) - 1:
            break
        sheet.write(i + 1, 0, voltage_list[i + 1][0])
        sheet.write(i + 1, 1, voltage_list[i + 1][1])
        sheet.write(i + 1, 2, voltage_list[i + 1][2])
        sheet.write(i + 1, 3, voltage_list[i + 1][3])
        sheet.write(i + 1, 4, voltage_list[i + 1][4])
        sheet.write(i + 1, 5, voltage_list[i + 1][5])
        sheet.write(i + 1, 6, voltage_list[i + 1][6])
        if voltage_list[i + 1][7] != 'null' and voltage_list[i + 1][8] == 'null':
            ret = set_ac(voltage_list[i + 1][6], voltage_list[i + 1][5], voltage_list[i + 1][4], 120, 240, 0,
                         voltage_list[i + 1][3], voltage_list[i + 1][2], voltage_list[i + 1][1], 1, 1, 1, 50)
            line_voltage_list = line_to_line_voltage_calculate(voltage_list[i + 1][1], voltage_list[i + 1][2],
                                                               voltage_list[i + 1][3],
                                                               voltage_list[i + 1][4], voltage_list[i + 1][5],
                                                               voltage_list[i + 1][6])
            Average_line_voltage = (line_voltage_list[0] + line_voltage_list[1] + line_voltage_list[2]) / 3
            Phase_AB_Voltage = read_phase_ab_voltage(line_voltage_list[0], times=40)
            Phase_BC_Voltage = read_phase_bc_voltage(line_voltage_list[1], times=40)
            Phase_CA_Voltage = read_phase_ca_voltage(line_voltage_list[2], times=40)
            Average_ll_Voltage = read_average_ll_voltage(Average_line_voltage, times=40)
            sheet.write(i + 1, 7, Phase_AB_Voltage)
            sheet.write(i + 1, 8, Phase_BC_Voltage)
            sheet.write(i + 1, 9, Phase_CA_Voltage)
            sheet.write(i + 1, 10, Average_ll_Voltage)
            scale_AB = abs(Phase_AB_Voltage - line_voltage_list[0]) / line_voltage_list[0]
            scale_BC = abs(Phase_BC_Voltage - line_voltage_list[1]) / line_voltage_list[1]
            scale_CA = abs(Phase_CA_Voltage - line_voltage_list[2]) / line_voltage_list[2]
            scale_Average_ll_Voltage = abs(Average_ll_Voltage - Average_line_voltage) / Average_line_voltage
            sheet.write(i + 1, 11, f'{scale_AB:.2%}')
            sheet.write(i + 1, 12, f'{scale_BC:.2%}')
            sheet.write(i + 1, 13, f'{scale_CA:.2%}')
            sheet.write(i + 1, 14, f'{scale_Average_ll_Voltage:.2%}')
            if scale_AB * 100 <= voltage_list[i + 1][7] and scale_BC * 100 <= voltage_list[i + 1][
                7] and scale_CA * 100 <= voltage_list[i + 1][7]:
                sheet.write(i + 1, 15, 'Passed')
            else:
                sheet.write(i + 1, 15, 'Failed')
        if voltage_list[i + 1][7] == 'null' and voltage_list[i + 1][8] == 'null':
            ret = set_ac(voltage_list[i + 1][6], voltage_list[i + 1][5], voltage_list[i + 1][4], 120, 240, 0,
                         voltage_list[i + 1][3], voltage_list[i + 1][2], voltage_list[i + 1][1], 1, 1, 1, 50)
            line_voltage_list = line_to_line_voltage_calculate(voltage_list[i + 1][1], voltage_list[i + 1][2],
                                                               voltage_list[i + 1][3],
                                                               voltage_list[i + 1][4], voltage_list[i + 1][5],
                                                               voltage_list[i + 1][6])
            Average_line_voltage = (line_voltage_list[0] + line_voltage_list[1] + line_voltage_list[2]) / 3
            Phase_AB_Voltage = read_phase_ab_voltage(line_voltage_list[0], times=40)
            Phase_BC_Voltage = read_phase_bc_voltage(line_voltage_list[1], times=40)
            Phase_CA_Voltage = read_phase_ca_voltage(line_voltage_list[2], times=40)
            Average_ll_Voltage = read_average_ll_voltage(Average_line_voltage, times=40)
            sheet.write(i + 1, 7, Phase_AB_Voltage)
            sheet.write(i + 1, 8, Phase_BC_Voltage)
            sheet.write(i + 1, 9, Phase_CA_Voltage)
            sheet.write(i + 1, 10, Average_ll_Voltage)
            sheet.write(i + 1, 11, 'null')
            sheet.write(i + 1, 12, 'null')
            sheet.write(i + 1, 13, 'null')
            sheet.write(i + 1, 14, 'null')
            if voltage_list[i + 1][1] < 10 and voltage_list[i + 1][2] < 10 and voltage_list[i + 1][3] < 10:
                if Phase_AB_Voltage == Phase_BC_Voltage == Phase_CA_Voltage == Average_ll_Voltage == 0:
                    sheet.write(i + 1, 15, 'Passed')
                else:
                    sheet.write(i + 1, 15, 'Failed')
            else:
                if Phase_AB_Voltage != 0 and Phase_BC_Voltage != 0 and Phase_CA_Voltage != 0 and Average_ll_Voltage != 0:
                    sheet.write(i + 1, 15, 'Passed')
                else:
                    sheet.write(i + 1, 15, 'Failed')

        if voltage_list[i + 1][7] != 'null' and voltage_list[i + 1][8] != 'null':
            if voltage_list[i + 1][8] == 'ABC':
                set_phase_order(0)
                sheet.write(i + 1, 16, '相序设置：ABC')
            if voltage_list[i + 1][8] == 'ACB':
                set_phase_order(1)
                sheet.write(i + 1, 16, '相序设置：ACB')
            ret = set_ac(voltage_list[i + 1][6], voltage_list[i + 1][5], voltage_list[i + 1][4], 120, 240, 0,
                         voltage_list[i + 1][3], voltage_list[i + 1][2], voltage_list[i + 1][1], 1, 1, 1, 50)
            line_voltage_list = line_to_line_voltage_calculate(voltage_list[i + 1][1], voltage_list[i + 1][2],
                                                               voltage_list[i + 1][3],
                                                               voltage_list[i + 1][4], voltage_list[i + 1][5],
                                                               voltage_list[i + 1][6])
            Average_line_voltage = (line_voltage_list[0] + line_voltage_list[1] + line_voltage_list[2]) / 3
            Phase_AB_Voltage = read_phase_ab_voltage(line_voltage_list[0], times=40)
            Phase_BC_Voltage = read_phase_bc_voltage(line_voltage_list[1], times=40)
            Phase_CA_Voltage = read_phase_ca_voltage(line_voltage_list[2], times=40)
            Average_ll_Voltage = read_average_ll_voltage(Average_line_voltage, times=40)
            sheet.write(i + 1, 7, Phase_AB_Voltage)
            sheet.write(i + 1, 8, Phase_BC_Voltage)
            sheet.write(i + 1, 9, Phase_CA_Voltage)
            sheet.write(i + 1, 10, Average_ll_Voltage)
            scale_AB = abs(Phase_AB_Voltage - line_voltage_list[0]) / line_voltage_list[0]
            scale_BC = abs(Phase_BC_Voltage - line_voltage_list[1]) / line_voltage_list[1]
            scale_CA = abs(Phase_CA_Voltage - line_voltage_list[2]) / line_voltage_list[2]
            scale_Average_ll_Voltage = abs(Average_ll_Voltage - Average_line_voltage) / Average_line_voltage
            sheet.write(i + 1, 11, f'{scale_AB:.2%}')
            sheet.write(i + 1, 12, f'{scale_BC:.2%}')
            sheet.write(i + 1, 13, f'{scale_CA:.2%}')
            sheet.write(i + 1, 14, f'{scale_Average_ll_Voltage:.2%}')
            if scale_AB * 100 <= voltage_list[i + 1][7] and scale_BC * 100 <= voltage_list[i + 1][
                7] and scale_CA * 100 <= voltage_list[i + 1][7]:
                sheet.write(i + 1, 15, 'Passed')
            else:
                sheet.write(i + 1, 15, 'Failed')
            time.sleep(0.2)
            set_phase_order(0)
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def current_5a_333mv_ct_precision_measure():
    Current_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Current_5A_333mV_CT')
    print(Current_list)
    j = 30
    for i, Current in enumerate(Current_list):
        if 'Input and user' in Current:
            j = i
            break
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Current_5A_333mV_CT', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Phase A Current输入值')
    sheet.write(0, 2, 'Phase B Current输入值')
    sheet.write(0, 3, 'Phase C Current输入值')
    sheet.write(0, 4, '电表实际Phase A Current值')
    sheet.write(0, 5, '电表实际Phase B Current值')
    sheet.write(0, 6, '电表实际Phase C Current值')
    sheet.write(0, 7, '电表实际Iavg值')
    sheet.write(0, 8, 'Ia精度')
    sheet.write(0, 9, 'Ib精度')
    sheet.write(0, 10, 'Ic精度')
    sheet.write(0, 11, 'Iavg精度')
    sheet.write(0, 12, '测试结果')
    for i in range(len(Current_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Current_list[i]))
            print('测试进度:{}'.format(Current_list[i]))
        else:
            logging.info('测试进度:{},执行时间:{}'.format(Current_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Current_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Current_list) - 1:
            break
        if Current_list[i + 1][5] == '项电流':
            sheet.write(i + 1, 0, Current_list[i + 1][0])
            sheet.write(i + 1, 1, Current_list[i + 1][1])
            sheet.write(i + 1, 2, Current_list[i + 1][2])
            sheet.write(i + 1, 3, Current_list[i + 1][3])
            if Current_list[i + 1][1] != 'null' and Current_list[i + 1][2] != 'null' and Current_list[i + 1][
                3] != 'null' and Current_list[i + 1][4] != 'null' and Current_list[i + 1][6] == 'null':
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = read_phase_a_current(Current_list[i + 1][1], times=40)
                Phase_B_Current = read_phase_b_current(Current_list[i + 1][2], times=40)
                Phase_C_Current = read_phase_c_current(Current_list[i + 1][3], times=40)
                Iavg = read_system_average_current(Average_Current, times=40)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                scale_A_Current = abs(Phase_A_Current - Current_list[i + 1][1]) / Current_list[i + 1][1]
                scale_B_Current = abs(Phase_B_Current - Current_list[i + 1][2]) / Current_list[i + 1][2]
                scale_C_Current = abs(Phase_C_Current - Current_list[i + 1][3]) / Current_list[i + 1][3]
                scale_Iavg = abs(Iavg - Average_Current) / Average_Current
                sheet.write(i + 1, 8, f'{scale_A_Current:.2%}')
                sheet.write(i + 1, 9, f'{scale_B_Current:.2%}')
                sheet.write(i + 1, 10, f'{scale_C_Current:.2%}')
                sheet.write(i + 1, 11, f'{scale_Iavg:.2%}')
                if scale_A_Current * 100 <= Current_list[i + 1][4] and scale_B_Current * 100 <= Current_list[i + 1][
                    4] and scale_C_Current * 100 <= Current_list[i + 1][4] and scale_Iavg * 100 <= Current_list[i + 1][
                    4]:
                    sheet.write(i + 1, 12, 'Passed')
                else:
                    sheet.write(i + 1, 12, 'Failed')
            if Current_list[i + 1][1] == 'null' or Current_list[i + 1][2] == 'null' or Current_list[i + 1][
                3] == 'null' and Current_list[i + 1][4] != 'null' and Current_list[i + 1][6] == 'null':
                if Current_list[i + 1][1] == 'null':
                    Current_list[i + 1][1] = 0
                    sheet.write(i + 1, 8, 'null')
                if Current_list[i + 1][2] == 'null':
                    Current_list[i + 1][2] = 0
                    sheet.write(i + 1, 9, 'null')
                if Current_list[i + 1][3] == 'null':
                    Current_list[i + 1][3] = 0
                    sheet.write(i + 1, 10, 'null')
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = read_phase_a_current(Current_list[i + 1][1], times=40)
                Phase_B_Current = read_phase_b_current(Current_list[i + 1][2], times=40)
                Phase_C_Current = read_phase_c_current(Current_list[i + 1][3], times=40)
                Iavg = read_system_average_current(Average_Current, times=40)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                scale_A_Current = 0
                scale_B_Current = 0
                scale_C_Current = 0
                if Current_list[i + 1][1] != 'null':
                    if Current_list[i + 1][1] != 0:
                        scale_A_Current = abs(Phase_A_Current - Current_list[i + 1][1]) / Current_list[i + 1][1]
                        sheet.write(i + 1, 8, f'{scale_A_Current:.2%}')
                    else:
                        sheet.write(i + 1, 8, 'null')
                if Current_list[i + 1][2] != 'null':
                    if Current_list[i + 1][2] != 0:
                        scale_B_Current = abs(Phase_B_Current - Current_list[i + 1][2]) / Current_list[i + 1][2]
                        sheet.write(i + 1, 9, f'{scale_B_Current:.2%}')
                    else:
                        sheet.write(i + 1, 9, 'null')
                if Current_list[i + 1][3] != 'null':
                    if Current_list[i + 1][3] != 0:
                        scale_C_Current = abs(Phase_C_Current - Current_list[i + 1][3]) / Current_list[i + 1][3]
                        sheet.write(i + 1, 10, f'{scale_C_Current:.2%}')
                    else:
                        sheet.write(i + 1, 10, 'null')
                scale_Iavg = abs(Iavg - Average_Current) / Average_Current
                sheet.write(i + 1, 11, f'{scale_Iavg:.2%}')
                if scale_A_Current * 100 <= Current_list[i + 1][4] and scale_B_Current * 100 <= Current_list[i + 1][
                    4] and scale_C_Current * 100 <= Current_list[i + 1][4] and scale_Iavg * 100 <= Current_list[i + 1][
                    4]:
                    sheet.write(i + 1, 12, 'Passed')
                else:
                    sheet.write(i + 1, 12, 'Failed')
            if Current_list[i + 1][4] == 'null' and Current_list[i + 1][6] == 'null':
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = read_phase_a_current(Current_list[i + 1][1], times=40)
                Phase_B_Current = read_phase_b_current(Current_list[i + 1][2], times=40)
                Phase_C_Current = read_phase_c_current(Current_list[i + 1][3], times=40)
                Iavg = read_system_average_current(Average_Current, times=40)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                sheet.write(i + 1, 8, 'null')
                sheet.write(i + 1, 9, 'null')
                sheet.write(i + 1, 10, 'null')
                sheet.write(i + 1, 11, 'null')
                if Current_list[i + 1][1] < 0.005 and Current_list[i + 1][2] < 0.005 and Current_list[i + 1][3] < 0.005:
                    if Phase_A_Current == Phase_B_Current == Phase_C_Current == Iavg == 0:
                        sheet.write(i + 1, 12, 'Passed')
                    else:
                        sheet.write(i + 1, 12, 'Failed')
                else:
                    if Phase_A_Current != 0 and Phase_B_Current != 0 and Phase_C_Current != 0 and Iavg != 0:
                        sheet.write(i + 1, 12, 'Passed')
                    else:
                        sheet.write(i + 1, 12, 'Failed')

            if Current_list[i + 1][1] != 'null' and Current_list[i + 1][2] != 'null' and Current_list[i + 1][
                3] != 'null' and Current_list[i + 1][4] != 'null' and Current_list[i + 1][6] == 'ABC':
                set_phase_order(0)
                sheet.write(i + 1, 13, '相序设置：ABC')
                sheet.write(i + 1, 14,
                            f'输入相角：{Current_list[i + 1][7]}、{Current_list[i + 1][8]}、{Current_list[i + 1][9]}')
                ret = set_ac(120, 240, 0, Current_list[i + 1][9], Current_list[i + 1][8], Current_list[i + 1][7], 50,
                             50, 50, Current_list[i + 1][3], Current_list[i + 1][2], Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = read_phase_a_current(Current_list[i + 1][1], times=40)
                Phase_B_Current = read_phase_b_current(Current_list[i + 1][2], times=40)
                Phase_C_Current = read_phase_c_current(Current_list[i + 1][3], times=40)
                Iavg = read_system_average_current(Average_Current, times=40)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                scale_A_Current = abs(Phase_A_Current - Current_list[i + 1][1]) / Current_list[i + 1][1]
                scale_B_Current = abs(Phase_B_Current - Current_list[i + 1][2]) / Current_list[i + 1][2]
                scale_C_Current = abs(Phase_C_Current - Current_list[i + 1][3]) / Current_list[i + 1][3]
                scale_Iavg = abs(Iavg - Average_Current) / Average_Current
                sheet.write(i + 1, 8, f'{scale_A_Current:.2%}')
                sheet.write(i + 1, 9, f'{scale_B_Current:.2%}')
                sheet.write(i + 1, 10, f'{scale_C_Current:.2%}')
                sheet.write(i + 1, 11, f'{scale_Iavg:.2%}')
                if scale_A_Current * 100 <= Current_list[i + 1][4] and scale_B_Current * 100 <= Current_list[i + 1][
                    4] and scale_C_Current * 100 <= Current_list[i + 1][4] and scale_Iavg * 100 <= Current_list[i + 1][
                    4]:
                    sheet.write(i + 1, 12, 'Passed')
                else:
                    sheet.write(i + 1, 12, 'Failed')

        elif Current_list[i + 1][5] == 'Input and user':
            sheet.write(j, 0, '测试用例')
            sheet.write(j, 1, 'Phase A Current输入值')
            sheet.write(j, 2, 'Phase B Current输入值')
            sheet.write(j, 3, 'Phase C Current输入值')
            sheet.write(j, 4, '电表实际Input1 Current值')
            sheet.write(j, 5, '电表实际Input2 Current值')
            sheet.write(j, 6, '电表实际Input3 Current值')
            sheet.write(j, 7, '电表实际User1 Current值')
            sheet.write(j, 8, '电表实际User2 Current值')
            sheet.write(j, 9, '电表实际User3 Current值')
            sheet.write(j, 10, '电表实际Phase A Current值值')
            sheet.write(j, 11, '电表实际Phase B Current值值')
            sheet.write(j, 12, '电表实际Phase C Current值值')
            sheet.write(j, 13, 'Input1 精度')
            sheet.write(j, 14, 'Input2 精度')
            sheet.write(j, 15, 'Input3 精度')
            sheet.write(j, 16, 'User1 精度')
            sheet.write(j, 17, 'User2 精度')
            sheet.write(j, 18, 'User3 精度')
            sheet.write(j, 19, 'Phase A 精度')
            sheet.write(j, 20, 'Phase B 精度')
            sheet.write(j, 21, 'Phase C 精度')
            sheet.write(j, 22, '测试结果')
            sheet.write(i + 2, 0, Current_list[i + 1][0])
            sheet.write(i + 2, 1, Current_list[i + 1][1])
            sheet.write(i + 2, 2, Current_list[i + 1][2])
            sheet.write(i + 2, 3, Current_list[i + 1][3])
            if Current_list[i + 1][4] != 'null' and Current_list[i + 1][6] == '1E1p2w':
                set_service_configuration(0)
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Input_Channel_1_Current = read_input_channel_1_current(Current_list[i + 1][1], times=40)
                Input_Channel_2_Current = read_input_channel_2_current(Current_list[i + 1][2], times=40)
                Input_Channel_3_Current = read_input_channel_3_current(Current_list[i + 1][3], times=40)
                User_Channel_1_Current = read_user_channel_1_current(Current_list[i + 1][1], times=40)
                User_Channel_2_Current = read_user_channel_2_current(Current_list[i + 1][2], times=40)
                User_Channel_3_Current = read_user_channel_3_current(Current_list[i + 1][3], times=40)
                Phase_A_Current_standard = Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]
                Phase_A_Current = read_phase_a_current(Phase_A_Current_standard, times=40)
                Phase_B_Current = read_phase_b_current(0, times=40)
                Phase_C_Current = read_phase_c_current(0, times=40)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                sheet.write(i + 2, 10, Phase_A_Current)
                sheet.write(i + 2, 11, Phase_B_Current)
                sheet.write(i + 2, 12, Phase_C_Current)
                scale_Input_Channel_1_Current = abs(Input_Channel_1_Current - Current_list[i + 1][1]) / \
                                                Current_list[i + 1][1]
                scale_Input_Channel_2_Current = abs(Input_Channel_2_Current - Current_list[i + 1][2]) / \
                                                Current_list[i + 1][2]
                scale_Input_Channel_3_Current = abs(Input_Channel_3_Current - Current_list[i + 1][3]) / \
                                                Current_list[i + 1][3]
                scale_User_Channel_1_Current = abs(User_Channel_1_Current - Current_list[i + 1][1]) / \
                                               Current_list[i + 1][1]
                scale_User_Channel_2_Current = abs(User_Channel_2_Current - Current_list[i + 1][2]) / \
                                               Current_list[i + 1][2]
                scale_User_Channel_3_Current = abs(User_Channel_3_Current - Current_list[i + 1][3]) / \
                                               Current_list[i + 1][3]
                scale_Phase_A_Current = abs(Phase_A_Current - Phase_A_Current_standard) / Phase_A_Current_standard
                sheet.write(i + 2, 13, f'{scale_Input_Channel_1_Current:.2%}')
                sheet.write(i + 2, 14, f'{scale_Input_Channel_2_Current:.2%}')
                sheet.write(i + 2, 15, f'{scale_Input_Channel_3_Current:.2%}')
                sheet.write(i + 2, 16, f'{scale_User_Channel_1_Current:.2%}')
                sheet.write(i + 2, 17, f'{scale_User_Channel_2_Current:.2%}')
                sheet.write(i + 2, 18, f'{scale_User_Channel_3_Current:.2%}')
                sheet.write(i + 2, 19, f'{scale_Phase_A_Current:.2%}')
                sheet.write(i + 2, 20, f'null')
                sheet.write(i + 2, 21, f'null')
                if scale_Input_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_2_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_3_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_2_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_3_Current * 100 <= Current_list[i + 1][
                    4] and scale_Phase_A_Current * 100 <= Current_list[i + 1][
                    4] and Phase_B_Current == Phase_C_Current == 0:
                    sheet.write(i + 2, 22, 'Passed')
                else:
                    sheet.write(i + 2, 22, 'Failed')
            if Current_list[i + 1][4] != 'null' and Current_list[i + 1][6] == '2E3p3w':
                set_service_configuration(1)
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 0, Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Input_Channel_1_Current = read_input_channel_1_current(Current_list[i + 1][1], times=40)
                Input_Channel_2_Current = read_input_channel_2_current(Current_list[i + 1][2], times=40)
                Input_Channel_3_Current = read_input_channel_3_current(0, times=40)
                User_Channel_1_standard_Current = (Current_list[i + 1][1] + Current_list[i + 1][2]) / 2
                User_Channel_1_Current = read_user_channel_1_current(User_Channel_1_standard_Current, times=40)
                User_Channel_2_Current = read_user_channel_2_current(0, times=40)
                User_Channel_3_Current = read_user_channel_3_current(0, times=40)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                scale_Input_Channel_1_Current = abs(Input_Channel_1_Current - Current_list[i + 1][1]) / \
                                                Current_list[i + 1][1]
                scale_Input_Channel_2_Current = abs(Input_Channel_2_Current - Current_list[i + 1][2]) / \
                                                Current_list[i + 1][2]
                scale_User_Channel_1_Current = abs(User_Channel_1_Current - User_Channel_1_standard_Current) / \
                                               User_Channel_1_standard_Current
                sheet.write(i + 2, 13, f'{scale_Input_Channel_1_Current:.2%}')
                sheet.write(i + 2, 14, f'{scale_Input_Channel_2_Current:.2%}')
                sheet.write(i + 2, 15, f'null')
                sheet.write(i + 2, 16, f'{scale_User_Channel_1_Current:.2%}')
                sheet.write(i + 2, 17, f'null')
                sheet.write(i + 2, 18, f'null')
                if scale_Input_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_2_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and Input_Channel_3_Current == User_Channel_2_Current == User_Channel_3_Current == 0:
                    sheet.write(i + 2, 22, 'Passed')
                else:
                    sheet.write(i + 2, 22, 'Failed')

            if Current_list[i + 1][4] != 'null' and Current_list[i + 1][6] == '3E3p4w':
                set_service_configuration(4)
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Input_Channel_1_Current = read_input_channel_1_current(Current_list[i + 1][1], times=40)
                Input_Channel_2_Current = read_input_channel_2_current(Current_list[i + 1][2], times=40)
                Input_Channel_3_Current = read_input_channel_3_current(Current_list[i + 1][3], times=40)
                User_Channel_1_Standard_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] +
                                                   Current_list[i + 1][
                                                       3]) / 3
                User_Channel_1_Current = read_user_channel_1_current(User_Channel_1_Standard_Current, times=40)
                User_Channel_2_Current = read_user_channel_2_current(0, times=40)
                User_Channel_3_Current = read_user_channel_3_current(0, times=40)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                scale_Input_Channel_1_Current = abs(Input_Channel_1_Current - Current_list[i + 1][1]) / \
                                                Current_list[i + 1][1]
                scale_Input_Channel_2_Current = abs(Input_Channel_2_Current - Current_list[i + 1][2]) / \
                                                Current_list[i + 1][2]
                scale_Input_Channel_3_Current = abs(Input_Channel_3_Current - Current_list[i + 1][3]) / \
                                                Current_list[i + 1][3]
                scale_User_Channel_1_Current = abs(User_Channel_1_Current - User_Channel_1_Standard_Current) / \
                                               User_Channel_1_Standard_Current
                sheet.write(i + 2, 13, f'{scale_Input_Channel_1_Current:.2%}')
                sheet.write(i + 2, 14, f'{scale_Input_Channel_2_Current:.2%}')
                sheet.write(i + 2, 15, f'{scale_Input_Channel_3_Current:.2%}')
                sheet.write(i + 2, 16, f'{scale_User_Channel_1_Current:.2%}')
                sheet.write(i + 2, 17, f'null')
                sheet.write(i + 2, 18, f'null')
                if scale_Input_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_2_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_3_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and User_Channel_2_Current == User_Channel_3_Current == 0:
                    sheet.write(i + 2, 22, 'Passed')
                else:
                    sheet.write(i + 2, 22, 'Failed')

            if Current_list[i + 1][4] == 'null' and Current_list[i + 1][6] == '1E1p2w':
                set_service_configuration(0)
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Input_Channel_1_Current = read_input_channel_1_current(Current_list[i + 1][1], times=40)
                Input_Channel_2_Current = read_input_channel_2_current(Current_list[i + 1][2], times=40)
                Input_Channel_3_Current = read_input_channel_3_current(Current_list[i + 1][3], times=40)
                User_Channel_1_Current = read_user_channel_1_current(Current_list[i + 1][1], times=40)
                User_Channel_2_Current = read_user_channel_2_current(Current_list[i + 1][2], times=40)
                User_Channel_3_Current = read_user_channel_3_current(Current_list[i + 1][3], times=40)
                Phase_A_Current_standard = Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]
                Phase_A_Current = read_phase_a_current(Phase_A_Current_standard, times=40)
                Phase_B_Current = read_phase_b_current(0, times=40)
                Phase_C_Current = read_phase_c_current(0, times=40)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                sheet.write(i + 2, 10, Phase_A_Current)
                sheet.write(i + 2, 11, Phase_B_Current)
                sheet.write(i + 2, 12, Phase_C_Current)
                sheet.write(i + 2, 13, f'null')
                sheet.write(i + 2, 14, f'null')
                sheet.write(i + 2, 15, f'null')
                sheet.write(i + 2, 16, f'null')
                sheet.write(i + 2, 17, f'null')
                sheet.write(i + 2, 18, f'null')
                if Current_list[i + 1][1] >= 0.005 and Current_list[i + 1][2] >= 0.005 and Current_list[i + 1][
                    3] >= 0.005:
                    if Phase_A_Current != 0 and Input_Channel_1_Current != 0 and Input_Channel_2_Current != 0 and Input_Channel_3_Current != 0 and User_Channel_1_Current != 0 and User_Channel_2_Current != 0 and User_Channel_3_Current != 0:
                        sheet.write(i + 2, 22, 'Passed')
                    else:
                        sheet.write(i + 2, 22, 'Failed')
                if Current_list[i + 1][1] == 0 and Current_list[i + 1][2] == 0 and Current_list[i + 1][
                    3] == 0:
                    if Input_Channel_1_Current == Input_Channel_2_Current == Input_Channel_3_Current == User_Channel_1_Current == User_Channel_2_Current == User_Channel_3_Current == 0:
                        sheet.write(i + 2, 22, 'Passed')
                    else:
                        sheet.write(i + 2, 22, 'Failed')
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def current_20a_100ma_ct_precision_measure():
    Current_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Current_20A_100mA_CT')
    print(Current_list)
    j = 30
    for i, Current in enumerate(Current_list):
        if 'Input and user' in Current:
            j = i
            break
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Current_20A_100mA_CT', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Phase A Current输入值')
    sheet.write(0, 2, 'Phase B Current输入值')
    sheet.write(0, 3, 'Phase C Current输入值')
    sheet.write(0, 4, '电表实际Phase A Current值')
    sheet.write(0, 5, '电表实际Phase B Current值')
    sheet.write(0, 6, '电表实际Phase C Current值')
    sheet.write(0, 7, '电表实际Iavg值')
    sheet.write(0, 8, 'Ia精度')
    sheet.write(0, 9, 'Ib精度')
    sheet.write(0, 10, 'Ic精度')
    sheet.write(0, 11, 'Iavg精度')
    sheet.write(0, 12, '测试结果')
    for i in range(len(Current_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Current_list[i]))
            print('测试进度:{}'.format(Current_list[i]))
        else:
            logging.info('测试进度:{},执行时间:{}'.format(Current_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Current_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Current_list) - 1:
            break
        if Current_list[i + 1][5] == '项电流':
            sheet.write(i + 1, 0, Current_list[i + 1][0])
            sheet.write(i + 1, 1, Current_list[i + 1][1])
            sheet.write(i + 1, 2, Current_list[i + 1][2])
            sheet.write(i + 1, 3, Current_list[i + 1][3])
            if Current_list[i + 1][1] != 'null' and Current_list[i + 1][2] != 'null' and Current_list[i + 1][
                3] != 'null' and Current_list[i + 1][4] != 'null':
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = read_phase_a_current(Current_list[i + 1][1], times=40)
                Phase_B_Current = read_phase_b_current(Current_list[i + 1][2], times=40)
                Phase_C_Current = read_phase_c_current(Current_list[i + 1][3], times=40)
                Iavg = read_system_average_current(Average_Current, times=40)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                scale_A_Current = abs(Phase_A_Current - Current_list[i + 1][1]) / Current_list[i + 1][1]
                scale_B_Current = abs(Phase_B_Current - Current_list[i + 1][2]) / Current_list[i + 1][2]
                scale_C_Current = abs(Phase_C_Current - Current_list[i + 1][3]) / Current_list[i + 1][3]
                scale_Iavg = abs(Iavg - Average_Current) / Average_Current
                sheet.write(i + 1, 8, f'{scale_A_Current:.2%}')
                sheet.write(i + 1, 9, f'{scale_B_Current:.2%}')
                sheet.write(i + 1, 10, f'{scale_C_Current:.2%}')
                sheet.write(i + 1, 11, f'{scale_Iavg:.2%}')
                if scale_A_Current * 100 <= Current_list[i + 1][4] and scale_B_Current * 100 <= Current_list[i + 1][
                    4] and scale_C_Current * 100 <= Current_list[i + 1][4] and scale_Iavg * 100 <= Current_list[i + 1][
                    4]:
                    sheet.write(i + 1, 12, 'Passed')
                else:
                    sheet.write(i + 1, 12, 'Failed')
            if Current_list[i + 1][1] == 'null' or Current_list[i + 1][2] == 'null' or Current_list[i + 1][
                3] == 'null' and Current_list[i + 1][4] != 'null':
                if Current_list[i + 1][1] == 'null':
                    Current_list[i + 1][1] = 0
                    sheet.write(i + 1, 8, 'null')
                if Current_list[i + 1][2] == 'null':
                    Current_list[i + 1][2] = 0
                    sheet.write(i + 1, 9, 'null')
                if Current_list[i + 1][3] == 'null':
                    Current_list[i + 1][3] = 0
                    sheet.write(i + 1, 10, 'null')
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = read_phase_a_current(Current_list[i + 1][1], times=40)
                Phase_B_Current = read_phase_b_current(Current_list[i + 1][2], times=40)
                Phase_C_Current = read_phase_c_current(Current_list[i + 1][3], times=40)
                Iavg = read_system_average_current(Average_Current, times=40)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                scale_A_Current = 0
                scale_B_Current = 0
                scale_C_Current = 0
                if Current_list[i + 1][1] != 'null':
                    if Current_list[i + 1][1] != 0:
                        scale_A_Current = abs(Phase_A_Current - Current_list[i + 1][1]) / Current_list[i + 1][1]
                        sheet.write(i + 1, 8, f'{scale_A_Current:.2%}')
                    else:
                        sheet.write(i + 1, 8, 'null')
                if Current_list[i + 1][2] != 'null':
                    if Current_list[i + 1][2] != 0:
                        scale_B_Current = abs(Phase_B_Current - Current_list[i + 1][2]) / Current_list[i + 1][2]
                        sheet.write(i + 1, 9, f'{scale_B_Current:.2%}')
                    else:
                        sheet.write(i + 1, 9, 'null')
                if Current_list[i + 1][3] != 'null':
                    if Current_list[i + 1][3] != 0:
                        scale_C_Current = abs(Phase_C_Current - Current_list[i + 1][3]) / Current_list[i + 1][3]
                        sheet.write(i + 1, 10, f'{scale_C_Current:.2%}')
                    else:
                        sheet.write(i + 1, 10, 'null')
                scale_Iavg = abs(Iavg - Average_Current) / Average_Current
                sheet.write(i + 1, 11, f'{scale_Iavg:.2%}')
                if scale_A_Current * 100 <= Current_list[i + 1][4] and scale_B_Current * 100 <= Current_list[i + 1][
                    4] and scale_C_Current * 100 <= Current_list[i + 1][4] and scale_Iavg * 100 <= Current_list[i + 1][
                    4]:
                    sheet.write(i + 1, 12, 'Passed')
                else:
                    sheet.write(i + 1, 12, 'Failed')
            if Current_list[i + 1][4] == 'null':
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = read_phase_a_current(Current_list[i + 1][1], times=40)
                Phase_B_Current = read_phase_b_current(Current_list[i + 1][2], times=40)
                Phase_C_Current = read_phase_c_current(Current_list[i + 1][3], times=40)
                Iavg = read_system_average_current(Average_Current, times=40)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                sheet.write(i + 1, 8, 'null')
                sheet.write(i + 1, 9, 'null')
                sheet.write(i + 1, 10, 'null')
                sheet.write(i + 1, 11, 'null')
                if Current_list[i + 1][1] <= 0.02 and Current_list[i + 1][2] <= 0.02 and Current_list[i + 1][3] <= 0.02:
                    if Phase_A_Current == Phase_B_Current == Phase_C_Current == Iavg == 0:
                        sheet.write(i + 1, 12, 'Passed')
                    else:
                        sheet.write(i + 1, 12, 'Failed')
                else:
                    if Phase_A_Current != 0 and Phase_B_Current != 0 and Phase_C_Current != 0 and Iavg != 0:
                        sheet.write(i + 1, 12, 'Passed')
                    else:
                        sheet.write(i + 1, 12, 'Failed')
        elif Current_list[i + 1][5] == 'Input and user':
            sheet.write(j, 0, '测试用例')
            sheet.write(j, 1, 'Phase A Current输入值')
            sheet.write(j, 2, 'Phase B Current输入值')
            sheet.write(j, 3, 'Phase C Current输入值')
            sheet.write(j, 4, '电表实际Input1 Current值')
            sheet.write(j, 5, '电表实际Input2 Current值')
            sheet.write(j, 6, '电表实际Input3 Current值')
            sheet.write(j, 7, '电表实际User1 Current值')
            sheet.write(j, 8, '电表实际User2 Current值')
            sheet.write(j, 9, '电表实际User3 Current值')
            sheet.write(j, 10, '电表实际Phase A Current值值')
            sheet.write(j, 11, '电表实际Phase B Current值值')
            sheet.write(j, 12, '电表实际Phase C Current值值')
            sheet.write(j, 13, 'Input1 精度')
            sheet.write(j, 14, 'Input2 精度')
            sheet.write(j, 15, 'Input3 精度')
            sheet.write(j, 16, 'User1 精度')
            sheet.write(j, 17, 'User2 精度')
            sheet.write(j, 18, 'User3 精度')
            sheet.write(j, 19, 'Phase A 精度')
            sheet.write(j, 20, 'Phase B 精度')
            sheet.write(j, 21, 'Phase C 精度')
            sheet.write(j, 22, '测试结果')
            sheet.write(i + 2, 0, Current_list[i + 1][0])
            sheet.write(i + 2, 1, Current_list[i + 1][1])
            sheet.write(i + 2, 2, Current_list[i + 1][2])
            sheet.write(i + 2, 3, Current_list[i + 1][3])
            if Current_list[i + 1][4] != 'null' and Current_list[i + 1][6] == '1E1p2w':
                set_service_configuration(0)
                sheet.write(i + 2, 23, '1E1p2w')
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Input_Channel_1_Current = read_input_channel_1_current(Current_list[i + 1][1], times=40)
                Input_Channel_2_Current = read_input_channel_2_current(Current_list[i + 1][2], times=40)
                Input_Channel_3_Current = read_input_channel_3_current(Current_list[i + 1][3], times=40)
                User_Channel_1_Current = read_user_channel_1_current(Current_list[i + 1][1], times=40)
                User_Channel_2_Current = read_user_channel_2_current(Current_list[i + 1][2], times=40)
                User_Channel_3_Current = read_user_channel_3_current(Current_list[i + 1][3], times=40)
                Phase_A_Current_standard = Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]
                Phase_A_Current = read_phase_a_current(Phase_A_Current_standard, times=40)
                Phase_B_Current = read_phase_b_current(0, times=40)
                Phase_C_Current = read_phase_c_current(0, times=40)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                sheet.write(i + 2, 10, Phase_A_Current)
                sheet.write(i + 2, 11, Phase_B_Current)
                sheet.write(i + 2, 12, Phase_C_Current)
                scale_Input_Channel_1_Current = abs(Input_Channel_1_Current - Current_list[i + 1][1]) / \
                                                Current_list[i + 1][1]
                scale_Input_Channel_2_Current = abs(Input_Channel_2_Current - Current_list[i + 1][2]) / \
                                                Current_list[i + 1][2]
                scale_Input_Channel_3_Current = abs(Input_Channel_3_Current - Current_list[i + 1][3]) / \
                                                Current_list[i + 1][3]
                scale_User_Channel_1_Current = abs(User_Channel_1_Current - Current_list[i + 1][1]) / \
                                               Current_list[i + 1][1]
                scale_User_Channel_2_Current = abs(User_Channel_2_Current - Current_list[i + 1][2]) / \
                                               Current_list[i + 1][2]
                scale_User_Channel_3_Current = abs(User_Channel_3_Current - Current_list[i + 1][3]) / \
                                               Current_list[i + 1][3]
                scale_Phase_A_Current = abs(Phase_A_Current - Phase_A_Current_standard) / Phase_A_Current_standard
                sheet.write(i + 2, 13, f'{scale_Input_Channel_1_Current:.2%}')
                sheet.write(i + 2, 14, f'{scale_Input_Channel_2_Current:.2%}')
                sheet.write(i + 2, 15, f'{scale_Input_Channel_3_Current:.2%}')
                sheet.write(i + 2, 16, f'{scale_User_Channel_1_Current:.2%}')
                sheet.write(i + 2, 17, f'{scale_User_Channel_2_Current:.2%}')
                sheet.write(i + 2, 18, f'{scale_User_Channel_3_Current:.2%}')
                sheet.write(i + 2, 19, f'{scale_Phase_A_Current:.2%}')
                sheet.write(i + 2, 20, f'null')
                sheet.write(i + 2, 21, f'null')
                if scale_Input_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_2_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_3_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_2_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_3_Current * 100 <= Current_list[i + 1][
                    4] and scale_Phase_A_Current * 100 <= Current_list[i + 1][
                    4] and Phase_B_Current == Phase_C_Current == 0:
                    sheet.write(i + 2, 22, 'Passed')
                else:
                    sheet.write(i + 2, 22, 'Failed')
            if Current_list[i + 1][4] != 'null' and Current_list[i + 1][6] == '2E3p3w':
                set_service_configuration(1)
                sheet.write(i + 2, 23, '2E3p3w')
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 0, Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Input_Channel_1_Current = read_input_channel_1_current(Current_list[i + 1][1], times=40)
                Input_Channel_2_Current = read_input_channel_2_current(Current_list[i + 1][2], times=40)
                Input_Channel_3_Current = read_input_channel_3_current(0, times=40)
                User_Channel_1_standard_Current = (Current_list[i + 1][1] + Current_list[i + 1][2]) / 2
                User_Channel_1_Current = read_user_channel_1_current(User_Channel_1_standard_Current, times=40)
                User_Channel_2_Current = read_user_channel_2_current(0, times=40)
                User_Channel_3_Current = read_user_channel_3_current(0, times=40)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                scale_Input_Channel_1_Current = abs(Input_Channel_1_Current - Current_list[i + 1][1]) / \
                                                Current_list[i + 1][1]
                scale_Input_Channel_2_Current = abs(Input_Channel_2_Current - Current_list[i + 1][2]) / \
                                                Current_list[i + 1][2]
                scale_User_Channel_1_Current = abs(User_Channel_1_Current - User_Channel_1_standard_Current) / \
                                               User_Channel_1_standard_Current
                sheet.write(i + 2, 13, f'{scale_Input_Channel_1_Current:.2%}')
                sheet.write(i + 2, 14, f'{scale_Input_Channel_2_Current:.2%}')
                sheet.write(i + 2, 15, f'null')
                sheet.write(i + 2, 16, f'{scale_User_Channel_1_Current:.2%}')
                sheet.write(i + 2, 17, f'null')
                sheet.write(i + 2, 18, f'null')
                if scale_Input_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_2_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and Input_Channel_3_Current == User_Channel_2_Current == User_Channel_3_Current == 0:
                    sheet.write(i + 2, 22, 'Passed')
                else:
                    sheet.write(i + 2, 22, 'Failed')

            if Current_list[i + 1][4] != 'null' and Current_list[i + 1][6] == '3E3p4w':
                set_service_configuration(4)
                sheet.write(i + 2, 23, '3E3p4w')
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Input_Channel_1_Current = read_input_channel_1_current(Current_list[i + 1][1], times=40)
                Input_Channel_2_Current = read_input_channel_2_current(Current_list[i + 1][2], times=40)
                Input_Channel_3_Current = read_input_channel_3_current(Current_list[i + 1][3], times=40)
                User_Channel_1_Standard_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] +
                                                   Current_list[i + 1][
                                                       3]) / 3
                User_Channel_1_Current = read_user_channel_1_current(User_Channel_1_Standard_Current, times=40)
                User_Channel_2_Current = read_user_channel_2_current(0, times=40)
                User_Channel_3_Current = read_user_channel_3_current(0, times=40)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                scale_Input_Channel_1_Current = abs(Input_Channel_1_Current - Current_list[i + 1][1]) / \
                                                Current_list[i + 1][1]
                scale_Input_Channel_2_Current = abs(Input_Channel_2_Current - Current_list[i + 1][2]) / \
                                                Current_list[i + 1][2]
                scale_Input_Channel_3_Current = abs(Input_Channel_3_Current - Current_list[i + 1][3]) / \
                                                Current_list[i + 1][3]
                scale_User_Channel_1_Current = abs(User_Channel_1_Current - User_Channel_1_Standard_Current) / \
                                               User_Channel_1_Standard_Current
                sheet.write(i + 2, 13, f'{scale_Input_Channel_1_Current:.2%}')
                sheet.write(i + 2, 14, f'{scale_Input_Channel_2_Current:.2%}')
                sheet.write(i + 2, 15, f'{scale_Input_Channel_3_Current:.2%}')
                sheet.write(i + 2, 16, f'{scale_User_Channel_1_Current:.2%}')
                sheet.write(i + 2, 17, f'null')
                sheet.write(i + 2, 18, f'null')
                if scale_Input_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_2_Current * 100 <= Current_list[i + 1][
                    4] and scale_Input_Channel_3_Current * 100 <= Current_list[i + 1][
                    4] and scale_User_Channel_1_Current * 100 <= Current_list[i + 1][
                    4] and User_Channel_2_Current == User_Channel_3_Current == 0:
                    sheet.write(i + 2, 22, 'Passed')
                else:
                    sheet.write(i + 2, 22, 'Failed')
            if Current_list[i + 1][4] == 'null' and Current_list[i + 1][6] == '1E1p2w':
                ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                             Current_list[i + 1][1], 50)
                Input_Channel_1_Current = read_input_channel_1_current(Current_list[i + 1][1], times=40)
                Input_Channel_2_Current = read_input_channel_2_current(Current_list[i + 1][2], times=40)
                Input_Channel_3_Current = read_input_channel_3_current(Current_list[i + 1][3], times=40)
                User_Channel_1_Current = read_user_channel_1_current(Current_list[i + 1][1], times=40)
                User_Channel_2_Current = read_user_channel_2_current(Current_list[i + 1][2], times=40)
                User_Channel_3_Current = read_user_channel_3_current(Current_list[i + 1][3], times=40)
                Phase_A_Current_standard = Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]
                Phase_A_Current = read_phase_a_current(Phase_A_Current_standard, times=40)
                Phase_B_Current = read_phase_b_current(0, times=40)
                Phase_C_Current = read_phase_c_current(0, times=40)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                sheet.write(i + 2, 10, Phase_A_Current)
                sheet.write(i + 2, 11, Phase_B_Current)
                sheet.write(i + 2, 12, Phase_C_Current)
                sheet.write(i + 2, 13, f'null')
                sheet.write(i + 2, 14, f'null')
                sheet.write(i + 2, 15, f'null')
                sheet.write(i + 2, 16, f'null')
                sheet.write(i + 2, 17, f'null')
                sheet.write(i + 2, 18, f'null')
                if Current_list[i + 1][1] >= 0.02 and Current_list[i + 1][2] >= 0.02 and Current_list[i + 1][
                    3] >= 0.02:
                    if Phase_A_Current != 0 and Input_Channel_1_Current != 0 and Input_Channel_2_Current != 0 and Input_Channel_3_Current != 0 and User_Channel_1_Current != 0 and User_Channel_2_Current != 0 and User_Channel_3_Current != 0:
                        sheet.write(i + 2, 22, 'Passed')
                    else:
                        sheet.write(i + 2, 22, 'Failed')
                if Current_list[i + 1][1] == 0 and Current_list[i + 1][2] == 0 and Current_list[i + 1][
                    3] == 0:
                    if Input_Channel_1_Current == Input_Channel_2_Current == Input_Channel_3_Current == User_Channel_1_Current == User_Channel_2_Current == User_Channel_3_Current == 0:
                        sheet.write(i + 2, 22, 'Passed')
                    else:
                        sheet.write(i + 2, 22, 'Failed')
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def phase_voltage_angle_precision_measure():
    Voltage_Angle_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Phase_Voltage_Angle')
    print(Voltage_Angle_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Phase_Voltage_Angle', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Phase A Voltage Angle输入值')
    sheet.write(0, 2, 'Phase B Voltage Angle输入值')
    sheet.write(0, 3, 'Phase C Voltage Angle输入值')
    sheet.write(0, 4, '电表实际Phase A Voltage Angle值')
    sheet.write(0, 5, '电表实际Phase B Voltage Angle值')
    sheet.write(0, 6, '电表实际Phase C Voltage Angle值')
    sheet.write(0, 7, 'Va Angle精度')
    sheet.write(0, 8, 'Vb Angle精度')
    sheet.write(0, 9, 'Vc Angle精度')
    sheet.write(0, 10, '测试结果')
    for i in range(len(Voltage_Angle_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Voltage_Angle_list[i]))
            print('测试进度:{}'.format(Voltage_Angle_list[i]))
        else:
            logging.info('测试进度:{},执行时间:{}'.format(Voltage_Angle_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Voltage_Angle_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Voltage_Angle_list) - 1:
            break
        sheet.write(i + 1, 0, Voltage_Angle_list[i + 1][0])
        sheet.write(i + 1, 1, Voltage_Angle_list[i + 1][1])
        sheet.write(i + 1, 2, Voltage_Angle_list[i + 1][2])
        sheet.write(i + 1, 3, Voltage_Angle_list[i + 1][3])
        if Voltage_Angle_list[i + 1][3] != 'null' and Voltage_Angle_list[i + 1][1] != 'null' and \
                Voltage_Angle_list[i + 1][2] != 'null' and Voltage_Angle_list[i + 1][3] != 'null':
            ret = set_ac(Voltage_Angle_list[i + 1][3], Voltage_Angle_list[i + 1][2], Voltage_Angle_list[i + 1][1], 120,
                         240, 0, 100, 100, 100, 1, 1, 1, 50)
            Phase_A_Voltage_Angle = read_phase_a_voltage_angle(Voltage_Angle_list[i + 1][1], times=40)
            Phase_B_Voltage_Angle = read_phase_b_voltage_angle(Voltage_Angle_list[i + 1][2], times=40)
            Phase_C_Voltage_Angle = read_phase_c_voltage_angle(Voltage_Angle_list[i + 1][3], times=40)
            sheet.write(i + 1, 4, Phase_A_Voltage_Angle)
            sheet.write(i + 1, 5, Phase_B_Voltage_Angle)
            sheet.write(i + 1, 6, Phase_C_Voltage_Angle)
            scale_B_Voltage_Angle = Phase_B_Voltage_Angle - Voltage_Angle_list[i + 1][2]
            scale_C_Voltage_Angle = Phase_C_Voltage_Angle - Voltage_Angle_list[i + 1][3]
            if Phase_A_Voltage_Angle == 0:
                sheet.write(i + 1, 7, 0)
            else:
                sheet.write(i + 1, 7, 'null')
            sheet.write(i + 1, 8, f'{scale_B_Voltage_Angle}')
            sheet.write(i + 1, 9, f'{scale_C_Voltage_Angle}')
            if Phase_A_Voltage_Angle == 0 and abs(scale_B_Voltage_Angle) <= Voltage_Angle_list[i + 1][4] and abs(
                    scale_C_Voltage_Angle) <= Voltage_Angle_list[i + 1][4]:
                sheet.write(i + 1, 10, 'Passed')
            else:
                sheet.write(i + 1, 10, 'Failed')
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def input1_current_angle_precision_measure():
    Input_Current_Angle = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Input_Current_Angle')
    print(Input_Current_Angle)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Input_Current_Angle', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Phase A Current Angle输入值')
    sheet.write(0, 2, '电表实际Input1 Current Angle值')
    sheet.write(0, 3, 'Input1 Current Angle精度')
    sheet.write(0, 4, '测试结果')
    for i in range(len(Input_Current_Angle)):
        if i == 0:
            logging.info('测试进度:{}'.format(Input_Current_Angle[i]))
            print('测试进度:{}'.format(Input_Current_Angle[i]))
        else:
            logging.info('测试进度:{},执行时间:{}'.format(Input_Current_Angle[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Input_Current_Angle[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Input_Current_Angle) - 1:
            break
        sheet.write(i + 1, 0, Input_Current_Angle[i + 1][0])
        sheet.write(i + 1, 1, Input_Current_Angle[i + 1][1])
        if Input_Current_Angle[i + 1][4] != 'null' and Input_Current_Angle[i + 1][1] != 'null' and \
                Input_Current_Angle[i + 1][2] == 'null':
            ret = set_ac(120, 240, 0, 120, 240, Input_Current_Angle[i + 1][1], 100, 100, 100, 1, 1, 1, 50)
            Input_Channel_1_Current_Phase_Angle = read_input_channel_1_current_phase_angle(
                Input_Current_Angle[i + 1][1], times=40)
            sheet.write(i + 1, 2, Input_Channel_1_Current_Phase_Angle)
            if Input_Current_Angle[i + 1][1] == 0:
                sheet.write(i + 1, 3, 'null')
                if Input_Channel_1_Current_Phase_Angle == 0:
                    sheet.write(i + 1, 4, 'Passed')
                else:
                    sheet.write(i + 1, 4, 'Failed')
            else:
                scale_Input1_Current_Angle = Input_Channel_1_Current_Phase_Angle - Input_Current_Angle[i + 1][1]
                sheet.write(i + 1, 3, f'{scale_Input1_Current_Angle}')
                if abs(scale_Input1_Current_Angle) <= Input_Current_Angle[i + 1][4]:
                    sheet.write(i + 1, 4, 'Passed')
                else:
                    sheet.write(i + 1, 4, 'Failed')
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def power_5a_333mv_ct_precision_measure():
    Power_5A_333mV_CT_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Power_5A_333mV_CT')
    print(Power_5A_333mV_CT_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Power_5A_333mV_CT', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Va输入值')
    sheet.write(0, 2, 'Vb输入值')
    sheet.write(0, 3, 'Vc输入值')

    sheet.write(0, 4, 'Ia输入值')
    sheet.write(0, 5, 'Ib输入值')
    sheet.write(0, 6, 'Ic输入值')

    sheet.write(0, 7, 'Va_ang输入值')
    sheet.write(0, 8, 'Vb_ang输入值')
    sheet.write(0, 9, 'Vc_ang输入值')

    sheet.write(0, 10, 'Ia_ang输入值')
    sheet.write(0, 11, 'Ib_ang输入值')
    sheet.write(0, 12, 'Ic_ang输入值')

    sheet.write(0, 13, '电表Phase_A_P，精度')
    sheet.write(0, 14, '电表Phase_A_Q，精度')
    sheet.write(0, 15, '电表Phase_A_S，精度')
    sheet.write(0, 16, '电表Phase_A_PF，精度')

    sheet.write(0, 17, '电表Phase_B_P，精度')
    sheet.write(0, 18, '电表Phase_B_Q，精度')
    sheet.write(0, 19, '电表Phase_B_S，精度')
    sheet.write(0, 20, '电表Phase_B_PF，精度')

    sheet.write(0, 21, '电表Phase_C_P，精度')
    sheet.write(0, 22, '电表Phase_C_Q，精度')
    sheet.write(0, 23, '电表Phase_C_S，精度')
    sheet.write(0, 24, '电表Phase_C_PF，精度')

    sheet.write(0, 25, '电表Sys_P，精度')
    sheet.write(0, 26, '电表Sys_Q，精度')
    sheet.write(0, 27, '电表Sys_S，精度')
    sheet.write(0, 28, '电表Sys_PF，精度')

    sheet.write(0, 29, '电表Input1_P，精度')
    sheet.write(0, 30, '电表Input1_Q，精度')
    sheet.write(0, 31, '电表Input1_S，精度')
    sheet.write(0, 32, '电表Input1_PF，精度')

    sheet.write(0, 33, '电表Input2_P，精度')
    sheet.write(0, 34, '电表Input2_Q，精度')
    sheet.write(0, 35, '电表Input2_S，精度')
    sheet.write(0, 36, '电表Input2_PF，精度')

    sheet.write(0, 37, '电表Input3_P，精度')
    sheet.write(0, 38, '电表Input3_Q，精度')
    sheet.write(0, 39, '电表Input3_S，精度')
    sheet.write(0, 40, '电表Input3_PF，精度')

    sheet.write(0, 41, '电表User1_P，精度')
    sheet.write(0, 42, '电表User1_Q，精度')
    sheet.write(0, 43, '电表User1_S，精度')
    sheet.write(0, 44, '电表User1_PF，精度')

    sheet.write(0, 45, '电表User2_P，精度')
    sheet.write(0, 46, '电表User2_Q，精度')
    sheet.write(0, 47, '电表User2_S，精度')
    sheet.write(0, 48, '电表User2_PF，精度')

    sheet.write(0, 49, '电表User3_P，精度')
    sheet.write(0, 50, '电表User3_Q，精度')
    sheet.write(0, 51, '电表User3_S，精度')
    sheet.write(0, 52, '电表User3_PF，精度')

    sheet.write(0, 53, '测试结果')

    for i in range(len(Power_5A_333mV_CT_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Power_5A_333mV_CT_list[i]))
            print('测试进度:{}'.format(Power_5A_333mV_CT_list[i]))
        else:
            logging.info(
                '测试进度:{},执行时间:{}'.format(Power_5A_333mV_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Power_5A_333mV_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Power_5A_333mV_CT_list) - 1:
            break
        sheet.write(i + 1, 0, Power_5A_333mV_CT_list[i + 1][0])
        sheet.write(i + 1, 1, Power_5A_333mV_CT_list[i + 1][1])
        sheet.write(i + 1, 2, Power_5A_333mV_CT_list[i + 1][2])
        sheet.write(i + 1, 3, Power_5A_333mV_CT_list[i + 1][3])
        sheet.write(i + 1, 4, Power_5A_333mV_CT_list[i + 1][4])
        sheet.write(i + 1, 5, Power_5A_333mV_CT_list[i + 1][5])
        sheet.write(i + 1, 6, Power_5A_333mV_CT_list[i + 1][6])
        sheet.write(i + 1, 7, Power_5A_333mV_CT_list[i + 1][7])
        sheet.write(i + 1, 8, Power_5A_333mV_CT_list[i + 1][8])
        sheet.write(i + 1, 9, Power_5A_333mV_CT_list[i + 1][9])
        sheet.write(i + 1, 10, Power_5A_333mV_CT_list[i + 1][10])
        sheet.write(i + 1, 11, Power_5A_333mV_CT_list[i + 1][11])
        sheet.write(i + 1, 12, Power_5A_333mV_CT_list[i + 1][12])
        if Power_5A_333mV_CT_list[i + 1][14] == '3E3p4w' and Power_5A_333mV_CT_list[i + 1][13] != 'null' or \
                Power_5A_333mV_CT_list[i + 1][15] == 'True Reactive Power':
            set_service_configuration(4)
            set_reactive_power_calculation_methodme(0)
            set_ac(Power_5A_333mV_CT_list[i + 1][9], Power_5A_333mV_CT_list[i + 1][8], Power_5A_333mV_CT_list[i + 1][7],
                   Power_5A_333mV_CT_list[i + 1][12], Power_5A_333mV_CT_list[i + 1][11],
                   Power_5A_333mV_CT_list[i + 1][10], Power_5A_333mV_CT_list[i + 1][3],
                   Power_5A_333mV_CT_list[i + 1][2], Power_5A_333mV_CT_list[i + 1][1], Power_5A_333mV_CT_list[i + 1][6],
                   Power_5A_333mV_CT_list[i + 1][5], Power_5A_333mV_CT_list[i + 1][4], 50)
            AcuRev4100_Power = read_acurev4100_power_new(Power_5A_333mV_CT_list[i + 1][1],
                                                         Power_5A_333mV_CT_list[i + 1][2],
                                                         Power_5A_333mV_CT_list[i + 1][3],
                                                         Power_5A_333mV_CT_list[i + 1][4],
                                                         Power_5A_333mV_CT_list[i + 1][5],
                                                         Power_5A_333mV_CT_list[i + 1][6],
                                                         Power_5A_333mV_CT_list[i + 1][7],
                                                         Power_5A_333mV_CT_list[i + 1][8],
                                                         Power_5A_333mV_CT_list[i + 1][9],
                                                         Power_5A_333mV_CT_list[i + 1][10],
                                                         Power_5A_333mV_CT_list[i + 1][11],
                                                         Power_5A_333mV_CT_list[i + 1][12],
                                                         Power_5A_333mV_CT_list[i + 1][14])
            scale_list = []
            for j in range(len(AcuRev4100_Power)):
                if AcuRev4100_Power[j][1] != 'null':
                    sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]:.2%}')
                    scale_list.append(AcuRev4100_Power[j][1])
                else:
                    sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]}')
                    scale_list.append(AcuRev4100_Power[j][1])
            if Power_5A_333mV_CT_list[i + 1][15] != 'True Reactive Power':
                for k in range(len(scale_list)):
                    if scale_list[k] != 'null' and scale_list[k] * 100 <= Power_5A_333mV_CT_list[i + 1][13]:
                        sheet.write(i + 1, 53, 'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 53, 'Failed')
                        sheet.write(i + 1, 54, f'请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                        break
            else:
                if Power_5A_333mV_CT_list[i + 1][7] - Power_5A_333mV_CT_list[i + 1][10] == 0:
                    if AcuRev4100_Power[1][0] == AcuRev4100_Power[5][0] == AcuRev4100_Power[9][0] == \
                            AcuRev4100_Power[13][0] == AcuRev4100_Power[17][0] == AcuRev4100_Power[21][0] == \
                            AcuRev4100_Power[25][0] == 0:
                        sheet.write(i + 1, 53, 'Passed')
                else:
                    for k in range(len(scale_list)):
                        if scale_list[k] != 'null' and scale_list[k] * 100 <= Power_5A_333mV_CT_list[i + 1][13]:
                            sheet.write(i + 1, 53, 'Passed')
                            continue
                        else:
                            sheet.write(i + 1, 53, 'Failed')
                            sheet.write(i + 1, 54, f'请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                            break

        if Power_5A_333mV_CT_list[i + 1][14] == '3E3p4w' and Power_5A_333mV_CT_list[i + 1][13] == 'null':
            set_service_configuration(4)
            set_reactive_power_calculation_methodme(0)
            set_ac(Power_5A_333mV_CT_list[i + 1][9], Power_5A_333mV_CT_list[i + 1][8], Power_5A_333mV_CT_list[i + 1][7],
                   Power_5A_333mV_CT_list[i + 1][12], Power_5A_333mV_CT_list[i + 1][11],
                   Power_5A_333mV_CT_list[i + 1][10], Power_5A_333mV_CT_list[i + 1][3],
                   Power_5A_333mV_CT_list[i + 1][2], Power_5A_333mV_CT_list[i + 1][1], Power_5A_333mV_CT_list[i + 1][6],
                   Power_5A_333mV_CT_list[i + 1][5], Power_5A_333mV_CT_list[i + 1][4], 50)
            AcuRev4100_Power = read_acurev4100_power_new(Power_5A_333mV_CT_list[i + 1][1],
                                                         Power_5A_333mV_CT_list[i + 1][2],
                                                         Power_5A_333mV_CT_list[i + 1][3],
                                                         Power_5A_333mV_CT_list[i + 1][4],
                                                         Power_5A_333mV_CT_list[i + 1][5],
                                                         Power_5A_333mV_CT_list[i + 1][6],
                                                         Power_5A_333mV_CT_list[i + 1][7],
                                                         Power_5A_333mV_CT_list[i + 1][8],
                                                         Power_5A_333mV_CT_list[i + 1][9],
                                                         Power_5A_333mV_CT_list[i + 1][10],
                                                         Power_5A_333mV_CT_list[i + 1][11],
                                                         Power_5A_333mV_CT_list[i + 1][12],
                                                         Power_5A_333mV_CT_list[i + 1][14])
            Power_list = []
            for j in range(len(AcuRev4100_Power)):
                sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]}')
                Power_list.append(AcuRev4100_Power[j][0])
            if Power_5A_333mV_CT_list[i + 1][1] >= 10 and Power_5A_333mV_CT_list[i + 1][2] >= 10 and \
                    Power_5A_333mV_CT_list[i + 1][3] >= 10 and Power_5A_333mV_CT_list[i + 1][4] >= 0.005 and \
                    Power_5A_333mV_CT_list[i + 1][5] >= 0.005 and Power_5A_333mV_CT_list[i + 1][6] >= 0.005:
                for k in range(len(Power_list)):
                    if Power_list[k] != 0:
                        sheet.write(i + 1, 53, 'Passed')
                        break
                    else:
                        sheet.write(i + 1, 53, 'Failed')
                        sheet.write(i + 1, 54, f'请检查{k + 14}列数据')
                        break
            else:
                for k in range(len(Power_list)):
                    if Power_list[k] == 0:
                        sheet.write(i + 1, 53, 'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 53, 'Failed')
                        sheet.write(i + 1, 54, f'请检查{k + 14}列数据')
                        break
        if Power_5A_333mV_CT_list[i + 1][14] == '1E1p2w' and Power_5A_333mV_CT_list[i + 1][13] != 'null':
            set_service_configuration(0)
            set_ac(Power_5A_333mV_CT_list[i + 1][9], Power_5A_333mV_CT_list[i + 1][8], Power_5A_333mV_CT_list[i + 1][7],
                   Power_5A_333mV_CT_list[i + 1][12], Power_5A_333mV_CT_list[i + 1][11],
                   Power_5A_333mV_CT_list[i + 1][10], Power_5A_333mV_CT_list[i + 1][3],
                   Power_5A_333mV_CT_list[i + 1][2], Power_5A_333mV_CT_list[i + 1][1], Power_5A_333mV_CT_list[i + 1][6],
                   Power_5A_333mV_CT_list[i + 1][5], Power_5A_333mV_CT_list[i + 1][4], 50)
            AcuRev4100_Power = read_acurev4100_power_new(Power_5A_333mV_CT_list[i + 1][1],
                                                         Power_5A_333mV_CT_list[i + 1][2],
                                                         Power_5A_333mV_CT_list[i + 1][3],
                                                         Power_5A_333mV_CT_list[i + 1][4],
                                                         Power_5A_333mV_CT_list[i + 1][5],
                                                         Power_5A_333mV_CT_list[i + 1][6],
                                                         Power_5A_333mV_CT_list[i + 1][7],
                                                         Power_5A_333mV_CT_list[i + 1][8],
                                                         Power_5A_333mV_CT_list[i + 1][9],
                                                         Power_5A_333mV_CT_list[i + 1][10],
                                                         Power_5A_333mV_CT_list[i + 1][11],
                                                         Power_5A_333mV_CT_list[i + 1][12],
                                                         Power_5A_333mV_CT_list[i + 1][14])
            scale_list = []
            for j in range(len(AcuRev4100_Power)):
                if AcuRev4100_Power[j][1] != 'null':
                    sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]:.2%}')
                    scale_list.append(AcuRev4100_Power[j][1])
                else:
                    sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]}')
                    scale_list.append(AcuRev4100_Power[j][1])

            for k in range(len(scale_list)):
                if scale_list[k] != 'null' and scale_list[k] * 100 <= Power_5A_333mV_CT_list[i + 1][13]:
                    sheet.write(i + 1, 53, 'Passed')
                    continue
                else:
                    sheet.write(i + 1, 53, 'Failed')
                    sheet.write(i + 1, 54, f'请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                    break
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def power_20a_100ma_ct_precision_measure():
    Power_20A_100mA_CT_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Power_20A_100mA_CT1')
    print(Power_20A_100mA_CT_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Power_20A_100mA_CT', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Va输入值')
    sheet.write(0, 2, 'Vb输入值')
    sheet.write(0, 3, 'Vc输入值')

    sheet.write(0, 4, 'Ia输入值')
    sheet.write(0, 5, 'Ib输入值')
    sheet.write(0, 6, 'Ic输入值')

    sheet.write(0, 7, 'Va_ang输入值')
    sheet.write(0, 8, 'Vb_ang输入值')
    sheet.write(0, 9, 'Vc_ang输入值')

    sheet.write(0, 10, 'Ia_ang输入值')
    sheet.write(0, 11, 'Ib_ang输入值')
    sheet.write(0, 12, 'Ic_ang输入值')

    sheet.write(0, 13, '电表Phase_A_P，精度')
    sheet.write(0, 14, '电表Phase_A_Q，精度')
    sheet.write(0, 15, '电表Phase_A_S，精度')
    sheet.write(0, 16, '电表Phase_A_PF，精度')

    sheet.write(0, 17, '电表Phase_B_P，精度')
    sheet.write(0, 18, '电表Phase_B_Q，精度')
    sheet.write(0, 19, '电表Phase_B_S，精度')
    sheet.write(0, 20, '电表Phase_B_PF，精度')

    sheet.write(0, 21, '电表Phase_C_P，精度')
    sheet.write(0, 22, '电表Phase_C_Q，精度')
    sheet.write(0, 23, '电表Phase_C_S，精度')
    sheet.write(0, 24, '电表Phase_C_PF，精度')

    sheet.write(0, 25, '电表Sys_P，精度')
    sheet.write(0, 26, '电表Sys_Q，精度')
    sheet.write(0, 27, '电表Sys_S，精度')
    sheet.write(0, 28, '电表Sys_PF，精度')

    sheet.write(0, 29, '电表Input1_P，精度')
    sheet.write(0, 30, '电表Input1_Q，精度')
    sheet.write(0, 31, '电表Input1_S，精度')
    sheet.write(0, 32, '电表Input1_PF，精度')

    sheet.write(0, 33, '电表Input2_P，精度')
    sheet.write(0, 34, '电表Input2_Q，精度')
    sheet.write(0, 35, '电表Input2_S，精度')
    sheet.write(0, 36, '电表Input2_PF，精度')

    sheet.write(0, 37, '电表Input3_P，精度')
    sheet.write(0, 38, '电表Input3_Q，精度')
    sheet.write(0, 39, '电表Input3_S，精度')
    sheet.write(0, 40, '电表Input3_PF，精度')

    sheet.write(0, 41, '电表User1_P，精度')
    sheet.write(0, 42, '电表User1_Q，精度')
    sheet.write(0, 43, '电表User1_S，精度')
    sheet.write(0, 44, '电表User1_PF，精度')

    sheet.write(0, 45, '电表User2_P，精度')
    sheet.write(0, 46, '电表User2_Q，精度')
    sheet.write(0, 47, '电表User2_S，精度')
    sheet.write(0, 48, '电表User2_PF，精度')

    sheet.write(0, 49, '电表User3_P，精度')
    sheet.write(0, 50, '电表User3_Q，精度')
    sheet.write(0, 51, '电表User3_S，精度')
    sheet.write(0, 52, '电表User3_PF，精度')

    sheet.write(0, 53, '测试结果')

    for i in range(len(Power_20A_100mA_CT_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Power_20A_100mA_CT_list[i]))
            print('测试进度:{}'.format(Power_20A_100mA_CT_list[i]))
        else:
            logging.info(
                '测试进度:{},执行时间:{}'.format(Power_20A_100mA_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Power_20A_100mA_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Power_20A_100mA_CT_list) - 1:
            break
        sheet.write(i + 1, 0, Power_20A_100mA_CT_list[i + 1][0])
        sheet.write(i + 1, 1, Power_20A_100mA_CT_list[i + 1][1])
        sheet.write(i + 1, 2, Power_20A_100mA_CT_list[i + 1][2])
        sheet.write(i + 1, 3, Power_20A_100mA_CT_list[i + 1][3])
        sheet.write(i + 1, 4, Power_20A_100mA_CT_list[i + 1][4])
        sheet.write(i + 1, 5, Power_20A_100mA_CT_list[i + 1][5])
        sheet.write(i + 1, 6, Power_20A_100mA_CT_list[i + 1][6])
        sheet.write(i + 1, 7, Power_20A_100mA_CT_list[i + 1][7])
        sheet.write(i + 1, 8, Power_20A_100mA_CT_list[i + 1][8])
        sheet.write(i + 1, 9, Power_20A_100mA_CT_list[i + 1][9])
        sheet.write(i + 1, 10, Power_20A_100mA_CT_list[i + 1][10])
        sheet.write(i + 1, 11, Power_20A_100mA_CT_list[i + 1][11])
        sheet.write(i + 1, 12, Power_20A_100mA_CT_list[i + 1][12])
        if Power_20A_100mA_CT_list[i + 1][14] == '3E3p4w' and Power_20A_100mA_CT_list[i + 1][13] != 'null' or \
                Power_20A_100mA_CT_list[i + 1][15] == 'True Reactive Power':
            set_service_configuration(4)
            set_reactive_power_calculation_methodme(0)
            set_ac(Power_20A_100mA_CT_list[i + 1][9], Power_20A_100mA_CT_list[i + 1][8],
                   Power_20A_100mA_CT_list[i + 1][7],
                   Power_20A_100mA_CT_list[i + 1][12], Power_20A_100mA_CT_list[i + 1][11],
                   Power_20A_100mA_CT_list[i + 1][10], Power_20A_100mA_CT_list[i + 1][3],
                   Power_20A_100mA_CT_list[i + 1][2], Power_20A_100mA_CT_list[i + 1][1],
                   Power_20A_100mA_CT_list[i + 1][6],
                   Power_20A_100mA_CT_list[i + 1][5], Power_20A_100mA_CT_list[i + 1][4], 50)
            AcuRev4100_Power = read_acurev4100_power_new(Power_20A_100mA_CT_list[i + 1][1],
                                                         Power_20A_100mA_CT_list[i + 1][2],
                                                         Power_20A_100mA_CT_list[i + 1][3],
                                                         Power_20A_100mA_CT_list[i + 1][4],
                                                         Power_20A_100mA_CT_list[i + 1][5],
                                                         Power_20A_100mA_CT_list[i + 1][6],
                                                         Power_20A_100mA_CT_list[i + 1][7],
                                                         Power_20A_100mA_CT_list[i + 1][8],
                                                         Power_20A_100mA_CT_list[i + 1][9],
                                                         Power_20A_100mA_CT_list[i + 1][10],
                                                         Power_20A_100mA_CT_list[i + 1][11],
                                                         Power_20A_100mA_CT_list[i + 1][12],
                                                         Power_20A_100mA_CT_list[i + 1][14])
            scale_list = []
            for j in range(len(AcuRev4100_Power)):
                if AcuRev4100_Power[j][1] != 'null':
                    sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]:.2%}')
                    scale_list.append(AcuRev4100_Power[j][1])
                else:
                    sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]}')
                    scale_list.append(AcuRev4100_Power[j][1])
            if Power_20A_100mA_CT_list[i + 1][15] != 'True Reactive Power':
                for k in range(len(scale_list)):
                    if scale_list[k] != 'null' and scale_list[k] * 100 <= Power_20A_100mA_CT_list[i + 1][13] or \
                            scale_list[k] == 1:
                        sheet.write(i + 1, 53, 'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 53, 'Failed')
                        sheet.write(i + 1, 54, f'请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                        print(f'{AcuRev4100_Power}请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                        messagebox.showerror("错误",
                                             f'{AcuRev4100_Power}请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                        time.sleep(100000)
                        break
            else:
                if Power_20A_100mA_CT_list[i + 1][7] - Power_20A_100mA_CT_list[i + 1][10] == 0:
                    if AcuRev4100_Power[1][0] == AcuRev4100_Power[5][0] == AcuRev4100_Power[9][0] == \
                            AcuRev4100_Power[13][0] == AcuRev4100_Power[17][0] == AcuRev4100_Power[21][0] == \
                            AcuRev4100_Power[25][0] == 0:
                        sheet.write(i + 1, 53, 'Passed')
                else:
                    for k in range(len(scale_list)):
                        if scale_list[k] != 'null' and scale_list[k] * 100 <= Power_20A_100mA_CT_list[i + 1][13] or \
                                scale_list[k] == 1:
                            sheet.write(i + 1, 53, 'Passed')
                            continue
                        else:
                            sheet.write(i + 1, 53, 'Failed')
                            sheet.write(i + 1, 54, f'请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                            print(f'{AcuRev4100_Power}请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                            messagebox.showerror("错误",
                                                 f'{AcuRev4100_Power}请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                            time.sleep(100000)
                            break

        if Power_20A_100mA_CT_list[i + 1][14] == '3E3p4w' and Power_20A_100mA_CT_list[i + 1][13] == 'null':
            set_service_configuration(4)
            set_reactive_power_calculation_methodme(0)
            set_ac(Power_20A_100mA_CT_list[i + 1][9], Power_20A_100mA_CT_list[i + 1][8],
                   Power_20A_100mA_CT_list[i + 1][7],
                   Power_20A_100mA_CT_list[i + 1][12], Power_20A_100mA_CT_list[i + 1][11],
                   Power_20A_100mA_CT_list[i + 1][10], Power_20A_100mA_CT_list[i + 1][3],
                   Power_20A_100mA_CT_list[i + 1][2], Power_20A_100mA_CT_list[i + 1][1],
                   Power_20A_100mA_CT_list[i + 1][6],
                   Power_20A_100mA_CT_list[i + 1][5], Power_20A_100mA_CT_list[i + 1][4], 50)
            AcuRev4100_Power = read_acurev4100_power_new(Power_20A_100mA_CT_list[i + 1][1],
                                                         Power_20A_100mA_CT_list[i + 1][2],
                                                         Power_20A_100mA_CT_list[i + 1][3],
                                                         Power_20A_100mA_CT_list[i + 1][4],
                                                         Power_20A_100mA_CT_list[i + 1][5],
                                                         Power_20A_100mA_CT_list[i + 1][6],
                                                         Power_20A_100mA_CT_list[i + 1][7],
                                                         Power_20A_100mA_CT_list[i + 1][8],
                                                         Power_20A_100mA_CT_list[i + 1][9],
                                                         Power_20A_100mA_CT_list[i + 1][10],
                                                         Power_20A_100mA_CT_list[i + 1][11],
                                                         Power_20A_100mA_CT_list[i + 1][12],
                                                         Power_20A_100mA_CT_list[i + 1][14])
            Power_list = []
            for j in range(len(AcuRev4100_Power)):
                sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]}')
                Power_list.append(AcuRev4100_Power[j][0])
            if Power_20A_100mA_CT_list[i + 1][1] >= 10 and Power_20A_100mA_CT_list[i + 1][2] >= 10 and \
                    Power_20A_100mA_CT_list[i + 1][3] >= 10 and Power_20A_100mA_CT_list[i + 1][4] >= 0.02 and \
                    Power_20A_100mA_CT_list[i + 1][5] >= 0.02 and Power_20A_100mA_CT_list[i + 1][6] >= 0.02:
                for k in range(len(Power_list)):
                    if Power_list[k] != 0:
                        sheet.write(i + 1, 53, 'Passed')
                        break
                    else:
                        sheet.write(i + 1, 53, 'Failed')
                        sheet.write(i + 1, 54, f'请检查{k + 14}列数据')
                        print(f'{AcuRev4100_Power}请检查{k + 14}列数据')
                        messagebox.showerror("错误",
                                             f'{AcuRev4100_Power}请检查{k + 14}列数据')
                        time.sleep(100000)
                        break
            else:
                for k in range(len(Power_list)):
                    if Power_list[k] == 0:
                        sheet.write(i + 1, 53, 'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 53, 'Failed')
                        sheet.write(i + 1, 54, f'请检查{k + 14}列数据')
                        print(f'{AcuRev4100_Power}请检查{k + 14}列数据')
                        messagebox.showerror("错误",
                                             f'{AcuRev4100_Power}请检查{k + 14}列数据')
                        time.sleep(100000)
                        break
        if Power_20A_100mA_CT_list[i + 1][14] == '1E1p2w' and Power_20A_100mA_CT_list[i + 1][13] != 'null':
            set_service_configuration(0)
            set_reactive_power_calculation_methodme(0)
            set_ac(Power_20A_100mA_CT_list[i + 1][9], Power_20A_100mA_CT_list[i + 1][8],
                   Power_20A_100mA_CT_list[i + 1][7],
                   Power_20A_100mA_CT_list[i + 1][12], Power_20A_100mA_CT_list[i + 1][11],
                   Power_20A_100mA_CT_list[i + 1][10], Power_20A_100mA_CT_list[i + 1][3],
                   Power_20A_100mA_CT_list[i + 1][2], Power_20A_100mA_CT_list[i + 1][1],
                   Power_20A_100mA_CT_list[i + 1][6],
                   Power_20A_100mA_CT_list[i + 1][5], Power_20A_100mA_CT_list[i + 1][4], 50)
            AcuRev4100_Power = read_acurev4100_power_new(Power_20A_100mA_CT_list[i + 1][1],
                                                         Power_20A_100mA_CT_list[i + 1][2],
                                                         Power_20A_100mA_CT_list[i + 1][3],
                                                         Power_20A_100mA_CT_list[i + 1][4],
                                                         Power_20A_100mA_CT_list[i + 1][5],
                                                         Power_20A_100mA_CT_list[i + 1][6],
                                                         Power_20A_100mA_CT_list[i + 1][7],
                                                         Power_20A_100mA_CT_list[i + 1][8],
                                                         Power_20A_100mA_CT_list[i + 1][9],
                                                         Power_20A_100mA_CT_list[i + 1][10],
                                                         Power_20A_100mA_CT_list[i + 1][11],
                                                         Power_20A_100mA_CT_list[i + 1][12],
                                                         Power_20A_100mA_CT_list[i + 1][14])
            scale_list = []
            for j in range(len(AcuRev4100_Power)):
                if AcuRev4100_Power[j][1] != 'null':
                    sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]:.2%}')
                    scale_list.append(AcuRev4100_Power[j][1])
                else:
                    sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]}')
                    scale_list.append(AcuRev4100_Power[j][1])
            for k in range(len(scale_list)):
                if scale_list[k] != 'null' and scale_list[k] * 100 <= Power_20A_100mA_CT_list[i + 1][13] or \
                        scale_list[k] == 1:
                    sheet.write(i + 1, 53, 'Passed')
                    continue
                else:
                    sheet.write(i + 1, 53, 'Failed')
                    sheet.write(i + 1, 54, f'请检查{k + 14}列数据,精度{scale_list[k] * 100}%')
                    print(f'{AcuRev4100_Power}请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                    messagebox.showerror("错误",
                                         f'{AcuRev4100_Power}请检查{k + 14}列数据，精度{scale_list[k] * 100}%')
                    time.sleep(100000)
                    break
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def load_nature_measure():
    Load_Nature_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Load_Nature')
    print(Load_Nature_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Load_Nature', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Va输入值')
    sheet.write(0, 2, 'Vb输入值')
    sheet.write(0, 3, 'Vc输入值')

    sheet.write(0, 4, 'Ia输入值')
    sheet.write(0, 5, 'Ib输入值')
    sheet.write(0, 6, 'Ic输入值')

    sheet.write(0, 7, 'Va_ang输入值')
    sheet.write(0, 8, 'Vb_ang输入值')
    sheet.write(0, 9, 'Vc_ang输入值')

    sheet.write(0, 10, 'Ia_ang输入值')
    sheet.write(0, 11, 'Ib_ang输入值')
    sheet.write(0, 12, 'Ic_ang输入值')

    sheet.write(0, 13, '电表User1_Load_Nature')
    sheet.write(0, 14, '电表Input1_Load_Nature')
    sheet.write(0, 15, '电表Input2_Load_Nature')
    sheet.write(0, 16, '电表Input3_Load_Nature')

    sheet.write(0, 17, '电表PhaseA_Load_Nature')
    sheet.write(0, 18, '电表PhaseB_Load_Nature')
    sheet.write(0, 19, '电表PhaseC_Load_Nature')
    sheet.write(0, 20, '电表System_Load_Nature')

    sheet.write(0, 21, '测试结果')
    for i in range(len(Load_Nature_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Load_Nature_list[i]))
            print('测试进度:{}'.format(Load_Nature_list[i]))
        else:
            logging.info(
                '测试进度:{},执行时间:{}'.format(Load_Nature_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Load_Nature_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Load_Nature_list) - 1:
            break
        sheet.write(i + 1, 0, Load_Nature_list[i + 1][0])
        sheet.write(i + 1, 1, Load_Nature_list[i + 1][1])
        sheet.write(i + 1, 2, Load_Nature_list[i + 1][2])
        sheet.write(i + 1, 3, Load_Nature_list[i + 1][3])
        sheet.write(i + 1, 4, Load_Nature_list[i + 1][4])
        sheet.write(i + 1, 5, Load_Nature_list[i + 1][5])
        sheet.write(i + 1, 6, Load_Nature_list[i + 1][6])
        sheet.write(i + 1, 7, Load_Nature_list[i + 1][7])
        sheet.write(i + 1, 8, Load_Nature_list[i + 1][8])
        sheet.write(i + 1, 9, Load_Nature_list[i + 1][9])
        sheet.write(i + 1, 10, Load_Nature_list[i + 1][10])
        sheet.write(i + 1, 11, Load_Nature_list[i + 1][11])
        sheet.write(i + 1, 12, Load_Nature_list[i + 1][12])
        if Load_Nature_list[i + 1][13] == '一种负载类型' or Load_Nature_list[i + 1][14] == '1E1p2w':
            if Load_Nature_list[i + 1][14] == '1E1p2w':
                set_service_configuration(0)
            else:
                set_service_configuration(4)
            set_ac(Load_Nature_list[i + 1][9], Load_Nature_list[i + 1][8],
                   Load_Nature_list[i + 1][7],
                   Load_Nature_list[i + 1][12], Load_Nature_list[i + 1][11],
                   Load_Nature_list[i + 1][10], Load_Nature_list[i + 1][3],
                   Load_Nature_list[i + 1][2], Load_Nature_list[i + 1][1],
                   Load_Nature_list[i + 1][6],
                   Load_Nature_list[i + 1][5], Load_Nature_list[i + 1][4], 50)
            standard_Load_Nature = []
            user1_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][7], Load_Nature_list[i + 1][10])
            Input1_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][7], Load_Nature_list[i + 1][10])
            Input2_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][8], Load_Nature_list[i + 1][11])
            Input3_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][9], Load_Nature_list[i + 1][12])
            PhaseA_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][7], Load_Nature_list[i + 1][10])
            PhaseB_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][8], Load_Nature_list[i + 1][11])
            PhaseC_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][9], Load_Nature_list[i + 1][12])
            System_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][9], Load_Nature_list[i + 1][12])
            standard_Load_Nature.extend(
                [user1_standard_Load_Nature,
                 Input1_standard_Load_Nature, Input2_standard_Load_Nature, Input3_standard_Load_Nature,
                 PhaseA_standard_Load_Nature, PhaseB_standard_Load_Nature, PhaseC_standard_Load_Nature,
                 System_standard_Load_Nature])
            print(standard_Load_Nature)
            Acu4100_Load_Nature = []
            user1_Load_Nature = read_user_channel_1_load_nature()
            Input1_Load_Nature = read_input_channel_1_load_nature()
            Input2_Load_Nature = read_input_channel_2_load_nature()
            Input3_Load_Nature = read_input_channel_3_load_nature()
            Phase_A_Load_Nature = read_phase_a_load_nature()
            Phase_B_Load_Nature = read_phase_b_load_nature()
            Phase_C_Load_Nature = read_phase_c_load_nature()
            System_Load_Nature = read_system_load_nature()
            sheet.write(i + 1, 13, user1_Load_Nature)
            sheet.write(i + 1, 14, Input1_Load_Nature)
            sheet.write(i + 1, 15, Input2_Load_Nature)
            sheet.write(i + 1, 16, Input3_Load_Nature)
            sheet.write(i + 1, 17, Phase_A_Load_Nature)
            sheet.write(i + 1, 18, Phase_B_Load_Nature)
            sheet.write(i + 1, 19, Phase_C_Load_Nature)
            sheet.write(i + 1, 20, System_Load_Nature)
            Acu4100_Load_Nature.extend(
                [user1_Load_Nature, Input1_Load_Nature, Input2_Load_Nature, Input3_Load_Nature, Phase_A_Load_Nature,
                 Phase_B_Load_Nature, Phase_C_Load_Nature, System_Load_Nature])
            print(Acu4100_Load_Nature)
            if standard_Load_Nature == Acu4100_Load_Nature:
                sheet.write(i + 1, 21, 'Passed')
            else:
                sheet.write(i + 1, 21, 'Failed')
                sheet.write(i + 1, 22, f'标准值：{standard_Load_Nature}')
        if Load_Nature_list[i + 1][13] == '多种负载类型' and Load_Nature_list[i + 1][14] != '1E1p2w':
            set_service_configuration(4)
            set_ac(Load_Nature_list[i + 1][9], Load_Nature_list[i + 1][8],
                   Load_Nature_list[i + 1][7],
                   Load_Nature_list[i + 1][12], Load_Nature_list[i + 1][11],
                   Load_Nature_list[i + 1][10], Load_Nature_list[i + 1][3],
                   Load_Nature_list[i + 1][2], Load_Nature_list[i + 1][1],
                   Load_Nature_list[i + 1][6],
                   Load_Nature_list[i + 1][5], Load_Nature_list[i + 1][4], 50)
            standard_Load_Nature = []
            Input1_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][7], Load_Nature_list[i + 1][10])
            Input2_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][8], Load_Nature_list[i + 1][11])
            Input3_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][9], Load_Nature_list[i + 1][12])
            PhaseA_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][7], Load_Nature_list[i + 1][10])
            PhaseB_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][8], Load_Nature_list[i + 1][11])
            PhaseC_standard_Load_Nature = load_nature_calculate(Load_Nature_list[i + 1][9], Load_Nature_list[i + 1][12])
            A_Active_Power = active_power_calculate(Load_Nature_list[i + 1][1], Load_Nature_list[i + 1][4],
                                                    Load_Nature_list[i + 1][7],
                                                    Load_Nature_list[i + 1][10])
            B_Active_Power = active_power_calculate(Load_Nature_list[i + 1][2], Load_Nature_list[i + 1][5],
                                                    Load_Nature_list[i + 1][8],
                                                    Load_Nature_list[i + 1][11])
            C_Active_Power = active_power_calculate(Load_Nature_list[i + 1][3], Load_Nature_list[i + 1][6],
                                                    Load_Nature_list[i + 1][9],
                                                    Load_Nature_list[i + 1][12])
            P_Sum = A_Active_Power + B_Active_Power + C_Active_Power
            A_Reactive_Power = reactive_power_calculate(Load_Nature_list[i + 1][1], Load_Nature_list[i + 1][4],
                                                        Load_Nature_list[i + 1][7],
                                                        Load_Nature_list[i + 1][10])
            B_Reactive_Power = reactive_power_calculate(Load_Nature_list[i + 1][2], Load_Nature_list[i + 1][5],
                                                        Load_Nature_list[i + 1][8],
                                                        Load_Nature_list[i + 1][11])
            C_Reactive_Power = reactive_power_calculate(Load_Nature_list[i + 1][3], Load_Nature_list[i + 1][6],
                                                        Load_Nature_list[i + 1][9],
                                                        Load_Nature_list[i + 1][12])
            Q_Sum = A_Reactive_Power + B_Reactive_Power + C_Reactive_Power

            print(P_Sum, Q_Sum)
            System_standard_Load_Nature = system_load_nature_calculate(P_Sum, Q_Sum)
            user1_standard_Load_Nature = System_standard_Load_Nature
            standard_Load_Nature.extend(
                [user1_standard_Load_Nature,
                 Input1_standard_Load_Nature, Input2_standard_Load_Nature, Input3_standard_Load_Nature,
                 PhaseA_standard_Load_Nature, PhaseB_standard_Load_Nature, PhaseC_standard_Load_Nature,
                 System_standard_Load_Nature])
            print(standard_Load_Nature)
            Acu4100_Load_Nature = []
            user1_Load_Nature = read_user_channel_1_load_nature()
            Input1_Load_Nature = read_input_channel_1_load_nature()
            Input2_Load_Nature = read_input_channel_2_load_nature()
            Input3_Load_Nature = read_input_channel_3_load_nature()
            Phase_A_Load_Nature = read_phase_a_load_nature()
            Phase_B_Load_Nature = read_phase_b_load_nature()
            Phase_C_Load_Nature = read_phase_c_load_nature()
            System_Load_Nature = read_system_load_nature()
            sheet.write(i + 1, 13, user1_Load_Nature)
            sheet.write(i + 1, 14, Input1_Load_Nature)
            sheet.write(i + 1, 15, Input2_Load_Nature)
            sheet.write(i + 1, 16, Input3_Load_Nature)
            sheet.write(i + 1, 17, Phase_A_Load_Nature)
            sheet.write(i + 1, 18, Phase_B_Load_Nature)
            sheet.write(i + 1, 19, Phase_C_Load_Nature)
            sheet.write(i + 1, 20, System_Load_Nature)
            Acu4100_Load_Nature.extend(
                [user1_Load_Nature, Input1_Load_Nature, Input2_Load_Nature, Input3_Load_Nature, Phase_A_Load_Nature,
                 Phase_B_Load_Nature, Phase_C_Load_Nature, System_Load_Nature])
            print(Acu4100_Load_Nature)
            if standard_Load_Nature == Acu4100_Load_Nature:
                sheet.write(i + 1, 21, 'Passed')
            else:
                sheet.write(i + 1, 21, 'Failed')
                sheet.write(i + 1, 22, f'标准值：{standard_Load_Nature}')
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def energy_5a_333mv_ct_measure():
    Energy_5A_333mV_CT_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Energy_5A_333mV_CT')
    print(Energy_5A_333mV_CT_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Energy_5A_333mV_CT', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Va输入值')
    sheet.write(0, 2, 'Vb输入值')
    sheet.write(0, 3, 'Vc输入值')

    sheet.write(0, 4, 'Ia输入值')
    sheet.write(0, 5, 'Ib输入值')
    sheet.write(0, 6, 'Ic输入值')

    sheet.write(0, 7, 'Va_ang输入值')
    sheet.write(0, 8, 'Vb_ang输入值')
    sheet.write(0, 9, 'Vc_ang输入值')

    sheet.write(0, 10, 'Ia_ang输入值')
    sheet.write(0, 11, 'Ib_ang输入值')
    sheet.write(0, 12, 'Ic_ang输入值')

    sheet.write(0, 13, '等待时间(min)')

    sheet.write(0, 14, 'Phase_A_P_E_import')
    sheet.write(0, 15, 'Phase_A_P_E_export')
    sheet.write(0, 16, 'Phase_A_P_E_net')
    sheet.write(0, 17, 'Phase_A_P_E_total')
    sheet.write(0, 18, 'Phase_A_Q_E_import')
    sheet.write(0, 19, 'Phase_A_Q_E_export')
    sheet.write(0, 20, 'Phase_A_Q_E_net')
    sheet.write(0, 21, 'Phase_A_Q_E_total')
    sheet.write(0, 22, 'Phase_A_S_E')

    sheet.write(0, 23, 'Phase_B_P_E_import')
    sheet.write(0, 24, 'Phase_B_P_E_export')
    sheet.write(0, 25, 'Phase_B_P_E_net')
    sheet.write(0, 26, 'Phase_B_P_E_total')
    sheet.write(0, 27, 'Phase_B_Q_E_import')
    sheet.write(0, 28, 'Phase_B_Q_E_export')
    sheet.write(0, 29, 'Phase_B_Q_E_net')
    sheet.write(0, 30, 'Phase_B_Q_E_total')
    sheet.write(0, 31, 'Phase_B_S_E')

    sheet.write(0, 32, 'Phase_C_P_E_import')
    sheet.write(0, 33, 'Phase_C_P_E_export')
    sheet.write(0, 34, 'Phase_C_P_E_net')
    sheet.write(0, 35, 'Phase_C_P_E_total')
    sheet.write(0, 36, 'Phase_C_Q_E_import')
    sheet.write(0, 37, 'Phase_C_Q_E_export')
    sheet.write(0, 38, 'Phase_C_Q_E_net')
    sheet.write(0, 39, 'Phase_C_Q_E_total')
    sheet.write(0, 40, 'Phase_C_S_E')

    sheet.write(0, 41, 'System_P_E_import')
    sheet.write(0, 42, 'System_P_E_export')
    sheet.write(0, 43, 'System_P_E_net')
    sheet.write(0, 44, 'System_P_E_total')
    sheet.write(0, 45, 'System_Q_E_import')
    sheet.write(0, 46, 'System_Q_E_export')
    sheet.write(0, 47, 'System_Q_E_net')
    sheet.write(0, 48, 'System_Q_E_total')
    sheet.write(0, 49, 'System_S_E')

    sheet.write(0, 50, 'Input_1_P_E_import')
    sheet.write(0, 51, 'Input_1_P_E_export')
    sheet.write(0, 52, 'Input_1_P_E_net')
    sheet.write(0, 53, 'Input_1_P_E_total')
    sheet.write(0, 54, 'Input_1_Q_E_import')
    sheet.write(0, 55, 'Input_1_Q_E_export')
    sheet.write(0, 56, 'Input_1_Q_E_net')
    sheet.write(0, 57, 'Input_1_Q_E_total')
    sheet.write(0, 58, 'Input_1_S_E')

    sheet.write(0, 59, 'Input_2_P_E_import')
    sheet.write(0, 60, 'Input_2_P_E_export')
    sheet.write(0, 61, 'Input_2_P_E_net')
    sheet.write(0, 62, 'Input_2_P_E_total')
    sheet.write(0, 63, 'Input_2_Q_E_import')
    sheet.write(0, 64, 'Input_2_Q_E_export')
    sheet.write(0, 65, 'Input_2_Q_E_net')
    sheet.write(0, 66, 'Input_2_Q_E_total')
    sheet.write(0, 67, 'Input_2_S_E')

    sheet.write(0, 68, 'Input_3_P_E_import')
    sheet.write(0, 69, 'Input_3_P_E_export')
    sheet.write(0, 70, 'Input_3_P_E_net')
    sheet.write(0, 71, 'Input_3_P_E_total')
    sheet.write(0, 72, 'Input_3_Q_E_import')
    sheet.write(0, 73, 'Input_3_Q_E_export')
    sheet.write(0, 74, 'Input_3_Q_E_net')
    sheet.write(0, 75, 'Input_3_Q_E_total')
    sheet.write(0, 76, 'Input_3_S_E')

    sheet.write(0, 77, 'User_1_P_E_import')
    sheet.write(0, 78, 'User_1_P_E_export')
    sheet.write(0, 79, 'User_1_P_E_net')
    sheet.write(0, 80, 'User_1_P_E_total')
    sheet.write(0, 81, 'User_1_Q_E_import')
    sheet.write(0, 82, 'User_1_Q_E_export')
    sheet.write(0, 83, 'User_1_Q_E_net')
    sheet.write(0, 84, 'User_1_Q_E_total')
    sheet.write(0, 85, 'User_1_S_E')

    sheet.write(0, 86, 'User_2_P_E_import')
    sheet.write(0, 87, 'User_2_P_E_export')
    sheet.write(0, 88, 'User_2_P_E_net')
    sheet.write(0, 89, 'User_2_P_E_total')
    sheet.write(0, 90, 'User_2_Q_E_import')
    sheet.write(0, 91, 'User_2_Q_E_export')
    sheet.write(0, 92, 'User_2_Q_E_net')
    sheet.write(0, 93, 'User_2_Q_E_total')
    sheet.write(0, 94, 'User_2_S_E')

    sheet.write(0, 95, 'User_3_P_E_import')
    sheet.write(0, 96, 'User_3_P_E_export')
    sheet.write(0, 97, 'User_3_P_E_net')
    sheet.write(0, 98, 'User_3_P_E_total')
    sheet.write(0, 99, 'User_3_Q_E_import')
    sheet.write(0, 100, 'User_3_Q_E_export')
    sheet.write(0, 101, 'User_3_Q_E_net')
    sheet.write(0, 102, 'User_3_Q_E_total')
    sheet.write(0, 103, 'User_3_S_E')

    sheet.write(0, 104, '测试结果')
    for i in range(len(Energy_5A_333mV_CT_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Energy_5A_333mV_CT_list[i]))
            print('测试进度:{}'.format(Energy_5A_333mV_CT_list[i]))
        else:
            logging.info(
                '测试进度:{},执行时间:{}'.format(Energy_5A_333mV_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Energy_5A_333mV_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Energy_5A_333mV_CT_list) - 1:
            break
        sheet.write(i + 1, 0, Energy_5A_333mV_CT_list[i + 1][0])
        sheet.write(i + 1, 1, Energy_5A_333mV_CT_list[i + 1][1])
        sheet.write(i + 1, 2, Energy_5A_333mV_CT_list[i + 1][2])
        sheet.write(i + 1, 3, Energy_5A_333mV_CT_list[i + 1][3])
        sheet.write(i + 1, 4, Energy_5A_333mV_CT_list[i + 1][4])
        sheet.write(i + 1, 5, Energy_5A_333mV_CT_list[i + 1][5])
        sheet.write(i + 1, 6, Energy_5A_333mV_CT_list[i + 1][6])
        sheet.write(i + 1, 7, Energy_5A_333mV_CT_list[i + 1][7])
        sheet.write(i + 1, 8, Energy_5A_333mV_CT_list[i + 1][8])
        sheet.write(i + 1, 9, Energy_5A_333mV_CT_list[i + 1][9])
        sheet.write(i + 1, 10, Energy_5A_333mV_CT_list[i + 1][10])
        sheet.write(i + 1, 11, Energy_5A_333mV_CT_list[i + 1][11])
        sheet.write(i + 1, 12, Energy_5A_333mV_CT_list[i + 1][12])
        sheet.write(i + 1, 13, Energy_5A_333mV_CT_list[i + 1][13])
        if Energy_5A_333mV_CT_list[i + 1][15] == '3E3p4w' and Energy_5A_333mV_CT_list[i + 1][14] != 'null':
            set_service_configuration(4)
            time.sleep(1)
            # set_device_reboot(1)
            # time.sleep(15)
            set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
                   Energy_5A_333mV_CT_list[i + 1][7],
                   Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
                   Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
                   Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
                   Energy_5A_333mV_CT_list[i + 1][6],
                   Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
            set_clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = read_energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
                                                           Energy_5A_333mV_CT_list[i + 1][2],
                                                           Energy_5A_333mV_CT_list[i + 1][3],
                                                           Energy_5A_333mV_CT_list[i + 1][4],
                                                           Energy_5A_333mV_CT_list[i + 1][5],
                                                           Energy_5A_333mV_CT_list[i + 1][6],
                                                           Energy_5A_333mV_CT_list[i + 1][7],
                                                           Energy_5A_333mV_CT_list[i + 1][8],
                                                           Energy_5A_333mV_CT_list[i + 1][9],
                                                           Energy_5A_333mV_CT_list[i + 1][10],
                                                           Energy_5A_333mV_CT_list[i + 1][11],
                                                           Energy_5A_333mV_CT_list[i + 1][12],
                                                           Energy_5A_333mV_CT_list[i + 1][13],
                                                           Energy_5A_333mV_CT_list[i + 1][15])
            for j in range(len(Read_Energy_scale_list[0])):
                if Read_Energy_scale_list[1][j] != 'null':
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},{Read_Energy_scale_list[1][j]:.3%}')
                else:
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},null')
            for k in range(len(Read_Energy_scale_list[1])):
                if Read_Energy_scale_list[1][k] != 'null' and Read_Energy_scale_list[1][k] * 100 <= \
                        Energy_5A_333mV_CT_list[i + 1][14]:
                    sheet.write(i + 1, 104, f'Passed')
                    continue
                else:
                    sheet.write(i + 1, 104, f'Failed')
                    sheet.write(i + 1, 105, f'{k + 15}列精度不达标或null')
                    break
        if Energy_5A_333mV_CT_list[i + 1][15] == '1E1p2w' and Energy_5A_333mV_CT_list[i + 1][14] != 'null':
            set_service_configuration(0)
            time.sleep(1)
            # Set_channle2_voltage_assignment(1)
            # Set_channle3_voltage_assignment(2)
            # set_device_reboot(1)
            # time.sleep(15)
            set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
                   Energy_5A_333mV_CT_list[i + 1][7],
                   Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
                   Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
                   Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
                   Energy_5A_333mV_CT_list[i + 1][6],
                   Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
            set_clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = read_energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
                                                           Energy_5A_333mV_CT_list[i + 1][2],
                                                           Energy_5A_333mV_CT_list[i + 1][3],
                                                           Energy_5A_333mV_CT_list[i + 1][4],
                                                           Energy_5A_333mV_CT_list[i + 1][5],
                                                           Energy_5A_333mV_CT_list[i + 1][6],
                                                           Energy_5A_333mV_CT_list[i + 1][7],
                                                           Energy_5A_333mV_CT_list[i + 1][8],
                                                           Energy_5A_333mV_CT_list[i + 1][9],
                                                           Energy_5A_333mV_CT_list[i + 1][10],
                                                           Energy_5A_333mV_CT_list[i + 1][11],
                                                           Energy_5A_333mV_CT_list[i + 1][12],
                                                           Energy_5A_333mV_CT_list[i + 1][13],
                                                           Energy_5A_333mV_CT_list[i + 1][15])
            for j in range(len(Read_Energy_scale_list[0])):
                if Read_Energy_scale_list[1][j] != 'null':
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},{Read_Energy_scale_list[1][j]:.3%}')
                else:
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},null')
            for k in range(len(Read_Energy_scale_list[1])):
                if Read_Energy_scale_list[1][k] != 'null' and Read_Energy_scale_list[1][k] * 100 <= \
                        Energy_5A_333mV_CT_list[i + 1][14]:
                    sheet.write(i + 1, 104, f'Passed')
                    continue
                else:
                    sheet.write(i + 1, 104, f'Failed')
                    sheet.write(i + 1, 105, f'{k + 15}列精度不达标或null')
                    break
        if Energy_5A_333mV_CT_list[i + 1][15] == '3E3p4w' and Energy_5A_333mV_CT_list[i + 1][14] == 'null':
            set_service_configuration(4)
            time.sleep(1)
            # set_device_reboot(1)
            # time.sleep(15)
            set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
                   Energy_5A_333mV_CT_list[i + 1][7],
                   Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
                   Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
                   Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
                   Energy_5A_333mV_CT_list[i + 1][6],
                   Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
            set_clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = read_energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
                                                           Energy_5A_333mV_CT_list[i + 1][2],
                                                           Energy_5A_333mV_CT_list[i + 1][3],
                                                           Energy_5A_333mV_CT_list[i + 1][4],
                                                           Energy_5A_333mV_CT_list[i + 1][5],
                                                           Energy_5A_333mV_CT_list[i + 1][6],
                                                           Energy_5A_333mV_CT_list[i + 1][7],
                                                           Energy_5A_333mV_CT_list[i + 1][8],
                                                           Energy_5A_333mV_CT_list[i + 1][9],
                                                           Energy_5A_333mV_CT_list[i + 1][10],
                                                           Energy_5A_333mV_CT_list[i + 1][11],
                                                           Energy_5A_333mV_CT_list[i + 1][12],
                                                           Energy_5A_333mV_CT_list[i + 1][13],
                                                           Energy_5A_333mV_CT_list[i + 1][15])
            for j in range(len(Read_Energy_scale_list[0])):
                sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]}')
            for j in range(len(Read_Energy_scale_list[0])):
                if Energy_5A_333mV_CT_list[i + 1][4] < 0.005 and Energy_5A_333mV_CT_list[i + 1][5] < 0.005 and \
                        Energy_5A_333mV_CT_list[i + 1][6] < 0.005:
                    if Read_Energy_scale_list[0][j] == 0:
                        sheet.write(i + 1, 104, f'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 104, f'Failed')
                        sheet.write(i + 1, 105, f'{j + 15}列能量数据预期为0')
                        break
                else:
                    if Read_Energy_scale_list[0][j] != 0:
                        sheet.write(i + 1, 104, f'Passed')
                        break
                    else:
                        sheet.write(i + 1, 104, f'Failed')
                        continue
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def energy_20a_100ma_ct_measure():
    Energy_20A_100mA_CT_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Energy_20A_100mA_CT')
    print(Energy_20A_100mA_CT_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Energy_20A_100mA_CT', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Va输入值')
    sheet.write(0, 2, 'Vb输入值')
    sheet.write(0, 3, 'Vc输入值')

    sheet.write(0, 4, 'Ia输入值')
    sheet.write(0, 5, 'Ib输入值')
    sheet.write(0, 6, 'Ic输入值')

    sheet.write(0, 7, 'Va_ang输入值')
    sheet.write(0, 8, 'Vb_ang输入值')
    sheet.write(0, 9, 'Vc_ang输入值')

    sheet.write(0, 10, 'Ia_ang输入值')
    sheet.write(0, 11, 'Ib_ang输入值')
    sheet.write(0, 12, 'Ic_ang输入值')

    sheet.write(0, 13, '等待时间(min)')

    sheet.write(0, 14, 'Phase_A_P_E_import')
    sheet.write(0, 15, 'Phase_A_P_E_export')
    sheet.write(0, 16, 'Phase_A_P_E_net')
    sheet.write(0, 17, 'Phase_A_P_E_total')
    sheet.write(0, 18, 'Phase_A_Q_E_import')
    sheet.write(0, 19, 'Phase_A_Q_E_export')
    sheet.write(0, 20, 'Phase_A_Q_E_net')
    sheet.write(0, 21, 'Phase_A_Q_E_total')
    sheet.write(0, 22, 'Phase_A_S_E')

    sheet.write(0, 23, 'Phase_B_P_E_import')
    sheet.write(0, 24, 'Phase_B_P_E_export')
    sheet.write(0, 25, 'Phase_B_P_E_net')
    sheet.write(0, 26, 'Phase_B_P_E_total')
    sheet.write(0, 27, 'Phase_B_Q_E_import')
    sheet.write(0, 28, 'Phase_B_Q_E_export')
    sheet.write(0, 29, 'Phase_B_Q_E_net')
    sheet.write(0, 30, 'Phase_B_Q_E_total')
    sheet.write(0, 31, 'Phase_B_S_E')

    sheet.write(0, 32, 'Phase_C_P_E_import')
    sheet.write(0, 33, 'Phase_C_P_E_export')
    sheet.write(0, 34, 'Phase_C_P_E_net')
    sheet.write(0, 35, 'Phase_C_P_E_total')
    sheet.write(0, 36, 'Phase_C_Q_E_import')
    sheet.write(0, 37, 'Phase_C_Q_E_export')
    sheet.write(0, 38, 'Phase_C_Q_E_net')
    sheet.write(0, 39, 'Phase_C_Q_E_total')
    sheet.write(0, 40, 'Phase_C_S_E')

    sheet.write(0, 41, 'System_P_E_import')
    sheet.write(0, 42, 'System_P_E_export')
    sheet.write(0, 43, 'System_P_E_net')
    sheet.write(0, 44, 'System_P_E_total')
    sheet.write(0, 45, 'System_Q_E_import')
    sheet.write(0, 46, 'System_Q_E_export')
    sheet.write(0, 47, 'System_Q_E_net')
    sheet.write(0, 48, 'System_Q_E_total')
    sheet.write(0, 49, 'System_S_E')

    sheet.write(0, 50, 'Input_1_P_E_import')
    sheet.write(0, 51, 'Input_1_P_E_export')
    sheet.write(0, 52, 'Input_1_P_E_net')
    sheet.write(0, 53, 'Input_1_P_E_total')
    sheet.write(0, 54, 'Input_1_Q_E_import')
    sheet.write(0, 55, 'Input_1_Q_E_export')
    sheet.write(0, 56, 'Input_1_Q_E_net')
    sheet.write(0, 57, 'Input_1_Q_E_total')
    sheet.write(0, 58, 'Input_1_S_E')

    sheet.write(0, 59, 'Input_2_P_E_import')
    sheet.write(0, 60, 'Input_2_P_E_export')
    sheet.write(0, 61, 'Input_2_P_E_net')
    sheet.write(0, 62, 'Input_2_P_E_total')
    sheet.write(0, 63, 'Input_2_Q_E_import')
    sheet.write(0, 64, 'Input_2_Q_E_export')
    sheet.write(0, 65, 'Input_2_Q_E_net')
    sheet.write(0, 66, 'Input_2_Q_E_total')
    sheet.write(0, 67, 'Input_2_S_E')

    sheet.write(0, 68, 'Input_3_P_E_import')
    sheet.write(0, 69, 'Input_3_P_E_export')
    sheet.write(0, 70, 'Input_3_P_E_net')
    sheet.write(0, 71, 'Input_3_P_E_total')
    sheet.write(0, 72, 'Input_3_Q_E_import')
    sheet.write(0, 73, 'Input_3_Q_E_export')
    sheet.write(0, 74, 'Input_3_Q_E_net')
    sheet.write(0, 75, 'Input_3_Q_E_total')
    sheet.write(0, 76, 'Input_3_S_E')

    sheet.write(0, 77, 'User_1_P_E_import')
    sheet.write(0, 78, 'User_1_P_E_export')
    sheet.write(0, 79, 'User_1_P_E_net')
    sheet.write(0, 80, 'User_1_P_E_total')
    sheet.write(0, 81, 'User_1_Q_E_import')
    sheet.write(0, 82, 'User_1_Q_E_export')
    sheet.write(0, 83, 'User_1_Q_E_net')
    sheet.write(0, 84, 'User_1_Q_E_total')
    sheet.write(0, 85, 'User_1_S_E')

    sheet.write(0, 86, 'User_2_P_E_import')
    sheet.write(0, 87, 'User_2_P_E_export')
    sheet.write(0, 88, 'User_2_P_E_net')
    sheet.write(0, 89, 'User_2_P_E_total')
    sheet.write(0, 90, 'User_2_Q_E_import')
    sheet.write(0, 91, 'User_2_Q_E_export')
    sheet.write(0, 92, 'User_2_Q_E_net')
    sheet.write(0, 93, 'User_2_Q_E_total')
    sheet.write(0, 94, 'User_2_S_E')

    sheet.write(0, 95, 'User_3_P_E_import')
    sheet.write(0, 96, 'User_3_P_E_export')
    sheet.write(0, 97, 'User_3_P_E_net')
    sheet.write(0, 98, 'User_3_P_E_total')
    sheet.write(0, 99, 'User_3_Q_E_import')
    sheet.write(0, 100, 'User_3_Q_E_export')
    sheet.write(0, 101, 'User_3_Q_E_net')
    sheet.write(0, 102, 'User_3_Q_E_total')
    sheet.write(0, 103, 'User_3_S_E')

    sheet.write(0, 104, '测试结果')
    for i in range(len(Energy_20A_100mA_CT_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Energy_20A_100mA_CT_list[i]))
            print('测试进度:{}'.format(Energy_20A_100mA_CT_list[i]))
        else:
            logging.info(
                '测试进度:{},执行时间:{}'.format(Energy_20A_100mA_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Energy_20A_100mA_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Energy_20A_100mA_CT_list) - 1:
            break
        sheet.write(i + 1, 0, Energy_20A_100mA_CT_list[i + 1][0])
        sheet.write(i + 1, 1, Energy_20A_100mA_CT_list[i + 1][1])
        sheet.write(i + 1, 2, Energy_20A_100mA_CT_list[i + 1][2])
        sheet.write(i + 1, 3, Energy_20A_100mA_CT_list[i + 1][3])
        sheet.write(i + 1, 4, Energy_20A_100mA_CT_list[i + 1][4])
        sheet.write(i + 1, 5, Energy_20A_100mA_CT_list[i + 1][5])
        sheet.write(i + 1, 6, Energy_20A_100mA_CT_list[i + 1][6])
        sheet.write(i + 1, 7, Energy_20A_100mA_CT_list[i + 1][7])
        sheet.write(i + 1, 8, Energy_20A_100mA_CT_list[i + 1][8])
        sheet.write(i + 1, 9, Energy_20A_100mA_CT_list[i + 1][9])
        sheet.write(i + 1, 10, Energy_20A_100mA_CT_list[i + 1][10])
        sheet.write(i + 1, 11, Energy_20A_100mA_CT_list[i + 1][11])
        sheet.write(i + 1, 12, Energy_20A_100mA_CT_list[i + 1][12])
        sheet.write(i + 1, 13, Energy_20A_100mA_CT_list[i + 1][13])
        if Energy_20A_100mA_CT_list[i + 1][15] == '3E3p4w' and Energy_20A_100mA_CT_list[i + 1][14] != 'null':
            set_service_configuration(4)
            time.sleep(1)
            # set_device_reboot(1)
            # time.sleep(15)
            set_ac(Energy_20A_100mA_CT_list[i + 1][9], Energy_20A_100mA_CT_list[i + 1][8],
                   Energy_20A_100mA_CT_list[i + 1][7],
                   Energy_20A_100mA_CT_list[i + 1][12], Energy_20A_100mA_CT_list[i + 1][11],
                   Energy_20A_100mA_CT_list[i + 1][10], Energy_20A_100mA_CT_list[i + 1][3],
                   Energy_20A_100mA_CT_list[i + 1][2], Energy_20A_100mA_CT_list[i + 1][1],
                   Energy_20A_100mA_CT_list[i + 1][6],
                   Energy_20A_100mA_CT_list[i + 1][5], Energy_20A_100mA_CT_list[i + 1][4], 50)
            set_clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = read_energy_scale_new(Energy_20A_100mA_CT_list[i + 1][1],
                                                           Energy_20A_100mA_CT_list[i + 1][2],
                                                           Energy_20A_100mA_CT_list[i + 1][3],
                                                           Energy_20A_100mA_CT_list[i + 1][4],
                                                           Energy_20A_100mA_CT_list[i + 1][5],
                                                           Energy_20A_100mA_CT_list[i + 1][6],
                                                           Energy_20A_100mA_CT_list[i + 1][7],
                                                           Energy_20A_100mA_CT_list[i + 1][8],
                                                           Energy_20A_100mA_CT_list[i + 1][9],
                                                           Energy_20A_100mA_CT_list[i + 1][10],
                                                           Energy_20A_100mA_CT_list[i + 1][11],
                                                           Energy_20A_100mA_CT_list[i + 1][12],
                                                           Energy_20A_100mA_CT_list[i + 1][13],
                                                           Energy_20A_100mA_CT_list[i + 1][15])
            for j in range(len(Read_Energy_scale_list[0])):
                if Read_Energy_scale_list[1][j] != 'null':
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},{Read_Energy_scale_list[1][j]:.3%}')
                else:
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},null')
            for k in range(len(Read_Energy_scale_list[1])):
                if Read_Energy_scale_list[1][k] != 'null' and Read_Energy_scale_list[1][k] * 100 <= \
                        Energy_20A_100mA_CT_list[i + 1][14]:
                    sheet.write(i + 1, 104, f'Passed')
                    continue
                else:
                    sheet.write(i + 1, 104, f'Failed')
                    sheet.write(i + 1, 105, f'{k + 15}列精度不达标或null')
                    break
        if Energy_20A_100mA_CT_list[i + 1][15] == '1E1p2w' and Energy_20A_100mA_CT_list[i + 1][14] != 'null':
            set_service_configuration(0)
            time.sleep(1)
            # Set_channle2_voltage_assignment(1)
            # Set_channle3_voltage_assignment(2)
            # set_device_reboot(1)
            # time.sleep(15)
            set_ac(Energy_20A_100mA_CT_list[i + 1][9], Energy_20A_100mA_CT_list[i + 1][8],
                   Energy_20A_100mA_CT_list[i + 1][7],
                   Energy_20A_100mA_CT_list[i + 1][12], Energy_20A_100mA_CT_list[i + 1][11],
                   Energy_20A_100mA_CT_list[i + 1][10], Energy_20A_100mA_CT_list[i + 1][3],
                   Energy_20A_100mA_CT_list[i + 1][2], Energy_20A_100mA_CT_list[i + 1][1],
                   Energy_20A_100mA_CT_list[i + 1][6],
                   Energy_20A_100mA_CT_list[i + 1][5], Energy_20A_100mA_CT_list[i + 1][4], 50)
            set_clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = read_energy_scale_new(Energy_20A_100mA_CT_list[i + 1][1],
                                                           Energy_20A_100mA_CT_list[i + 1][2],
                                                           Energy_20A_100mA_CT_list[i + 1][3],
                                                           Energy_20A_100mA_CT_list[i + 1][4],
                                                           Energy_20A_100mA_CT_list[i + 1][5],
                                                           Energy_20A_100mA_CT_list[i + 1][6],
                                                           Energy_20A_100mA_CT_list[i + 1][7],
                                                           Energy_20A_100mA_CT_list[i + 1][8],
                                                           Energy_20A_100mA_CT_list[i + 1][9],
                                                           Energy_20A_100mA_CT_list[i + 1][10],
                                                           Energy_20A_100mA_CT_list[i + 1][11],
                                                           Energy_20A_100mA_CT_list[i + 1][12],
                                                           Energy_20A_100mA_CT_list[i + 1][13],
                                                           Energy_20A_100mA_CT_list[i + 1][15])
            for j in range(len(Read_Energy_scale_list[0])):
                if Read_Energy_scale_list[1][j] != 'null':
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},{Read_Energy_scale_list[1][j]:.3%}')
                else:
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},null')
            for k in range(len(Read_Energy_scale_list[1])):
                if Read_Energy_scale_list[1][k] != 'null' and Read_Energy_scale_list[1][k] * 100 <= \
                        Energy_20A_100mA_CT_list[i + 1][14]:
                    sheet.write(i + 1, 104, f'Passed')
                    continue
                else:
                    sheet.write(i + 1, 104, f'Failed')
                    sheet.write(i + 1, 105, f'{k + 15}列精度不达标或null')
                    break
        if Energy_20A_100mA_CT_list[i + 1][15] == '3E3p4w' and Energy_20A_100mA_CT_list[i + 1][14] == 'null':
            set_service_configuration(4)
            time.sleep(1)
            set_device_reboot(1)
            time.sleep(15)
            set_ac(Energy_20A_100mA_CT_list[i + 1][9], Energy_20A_100mA_CT_list[i + 1][8],
                   Energy_20A_100mA_CT_list[i + 1][7],
                   Energy_20A_100mA_CT_list[i + 1][12], Energy_20A_100mA_CT_list[i + 1][11],
                   Energy_20A_100mA_CT_list[i + 1][10], Energy_20A_100mA_CT_list[i + 1][3],
                   Energy_20A_100mA_CT_list[i + 1][2], Energy_20A_100mA_CT_list[i + 1][1],
                   Energy_20A_100mA_CT_list[i + 1][6],
                   Energy_20A_100mA_CT_list[i + 1][5], Energy_20A_100mA_CT_list[i + 1][4], 50)
            set_clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = read_energy_scale_new(Energy_20A_100mA_CT_list[i + 1][1],
                                                           Energy_20A_100mA_CT_list[i + 1][2],
                                                           Energy_20A_100mA_CT_list[i + 1][3],
                                                           Energy_20A_100mA_CT_list[i + 1][4],
                                                           Energy_20A_100mA_CT_list[i + 1][5],
                                                           Energy_20A_100mA_CT_list[i + 1][6],
                                                           Energy_20A_100mA_CT_list[i + 1][7],
                                                           Energy_20A_100mA_CT_list[i + 1][8],
                                                           Energy_20A_100mA_CT_list[i + 1][9],
                                                           Energy_20A_100mA_CT_list[i + 1][10],
                                                           Energy_20A_100mA_CT_list[i + 1][11],
                                                           Energy_20A_100mA_CT_list[i + 1][12],
                                                           Energy_20A_100mA_CT_list[i + 1][13],
                                                           Energy_20A_100mA_CT_list[i + 1][15])
            for j in range(len(Read_Energy_scale_list[0])):
                sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]}')
            for j in range(len(Read_Energy_scale_list[0])):
                if Energy_20A_100mA_CT_list[i + 1][4] < 0.02 and Energy_20A_100mA_CT_list[i + 1][5] < 0.02 and \
                        Energy_20A_100mA_CT_list[i + 1][6] < 0.02:
                    if Read_Energy_scale_list[0][j] == 0:
                        sheet.write(i + 1, 104, f'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 104, f'Failed')
                        sheet.write(i + 1, 105, f'{j + 15}列能量数据预期为0')
                        break
                else:
                    if Read_Energy_scale_list[0][j] != 0:
                        sheet.write(i + 1, 104, f'Passed')
                        break
                    else:
                        sheet.write(i + 1, 104, f'Failed')
                        continue
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def sequence_component_precision_measure():
    Sequence_Component_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Sequence_Component')
    print(Sequence_Component_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Sequence_Component', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'A_amplitude')
    sheet.write(0, 2, 'B_amplitude')
    sheet.write(0, 3, 'C_amplitude')
    sheet.write(0, 4, 'A_angle')
    sheet.write(0, 5, 'B_angle')
    sheet.write(0, 6, 'C_angle')
    sheet.write(0, 7, '零序分量(模)')
    sheet.write(0, 8, '零序角度(°)')
    sheet.write(0, 9, '正序分量(模)')
    sheet.write(0, 10, '正序角度(°)')
    sheet.write(0, 11, '负序分量(模)')
    sheet.write(0, 12, '负序角度(°)')
    sheet.write(0, 13, 'VUF/CUF(%)')
    sheet.write(0, 14, '测试结果')
    for i in range(len(Sequence_Component_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Sequence_Component_list[i]))
            print('测试进度:{}'.format(Sequence_Component_list[i]))
        else:
            logging.info(
                '测试进度:{},执行时间:{}'.format(Sequence_Component_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Sequence_Component_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Sequence_Component_list) - 1:
            break
        sheet.write(i + 1, 0, Sequence_Component_list[i + 1][0])
        sheet.write(i + 1, 1, Sequence_Component_list[i + 1][1])
        sheet.write(i + 1, 2, Sequence_Component_list[i + 1][2])
        sheet.write(i + 1, 3, Sequence_Component_list[i + 1][3])
        sheet.write(i + 1, 4, Sequence_Component_list[i + 1][4])
        sheet.write(i + 1, 5, Sequence_Component_list[i + 1][5])
        sheet.write(i + 1, 6, Sequence_Component_list[i + 1][6])
        if Sequence_Component_list[i + 1][7] != 'null' and Sequence_Component_list[i + 1][8] == 'VUF':
            set_ac(Sequence_Component_list[i + 1][6], Sequence_Component_list[i + 1][5],
                   Sequence_Component_list[i + 1][4], 120, 240, 0, Sequence_Component_list[i + 1][3],
                   Sequence_Component_list[i + 1][2], Sequence_Component_list[i + 1][1], 1, 1, 1, 50)
            set_service_configuration(4)
            sequence_component_List = sequence_component_calculation(Sequence_Component_list[i + 1][1],
                                                                     Sequence_Component_list[i + 1][2],
                                                                     Sequence_Component_list[i + 1][3],
                                                                     Sequence_Component_list[i + 1][4],
                                                                     Sequence_Component_list[i + 1][5],
                                                                     Sequence_Component_list[i + 1][6])
            VUF = read_voltage_unbalance_factor_magnitude(sequence_component_List[6], times=10)
            sheet.write(i + 1, 13, f'VUF:{VUF}')
            if sequence_component_List[6] * 0.99 <= VUF <= sequence_component_List[6] * 1.01:
                sheet.write(i + 1, 14, f'Passed')
            else:
                sheet.write(i + 1, 14, f'Failed')
                sheet.write(i + 1, 15, f'VUF标准值:{sequence_component_List[6]}')
        if Sequence_Component_list[i + 1][7] != 'null' and Sequence_Component_list[i + 1][8] == 'CUF':
            set_ac(120, 240, 0, Sequence_Component_list[i + 1][6], Sequence_Component_list[i + 1][5],
                   Sequence_Component_list[i + 1][4], 50, 50, 50, Sequence_Component_list[i + 1][3],
                   Sequence_Component_list[i + 1][2], Sequence_Component_list[i + 1][1], 50)
            set_service_configuration(4)
            sequence_component_List = sequence_component_calculation(Sequence_Component_list[i + 1][1],
                                                                     Sequence_Component_list[i + 1][2],
                                                                     Sequence_Component_list[i + 1][3],
                                                                     Sequence_Component_list[i + 1][4],
                                                                     Sequence_Component_list[i + 1][5],
                                                                     Sequence_Component_list[i + 1][6])
            CUF = read_user_channel_1_current_unbalance_factor_magnitude(sequence_component_List[6], times=10)
            sheet.write(i + 1, 13, f'CUF:{CUF}')
            if sequence_component_List[6] * 0.99 <= CUF <= sequence_component_List[6] * 1.01:
                sheet.write(i + 1, 14, f'Passed')
            else:
                sheet.write(i + 1, 14, f'Failed')
                sheet.write(i + 1, 15, f'CUF标准值:{sequence_component_List[6]}')
        if Sequence_Component_list[i + 1][7] != 'null' and Sequence_Component_list[i + 1][8] == '相电压序分量':
            set_ac(Sequence_Component_list[i + 1][6], Sequence_Component_list[i + 1][5],
                   Sequence_Component_list[i + 1][4], 120, 240, 0, Sequence_Component_list[i + 1][3],
                   Sequence_Component_list[i + 1][2], Sequence_Component_list[i + 1][1], 1, 1, 1, 50)
            set_service_configuration(4)
            sequence_component_List = sequence_component_calculation(Sequence_Component_list[i + 1][1],
                                                                     Sequence_Component_list[i + 1][2],
                                                                     Sequence_Component_list[i + 1][3],
                                                                     Sequence_Component_list[i + 1][4],
                                                                     Sequence_Component_list[i + 1][5],
                                                                     Sequence_Component_list[i + 1][6])
            Acu4100_Voltage_Zero_Sequence = read_voltage_zero_sequence_magnitude(sequence_component_List[0], times=40)
            Acu4100_Voltage_Zero_Sequence_Angle = read_voltage_zero_sequence_angle(sequence_component_List[1], times=40)
            Acu4100_Voltage_Positive_Sequence = read_voltage_positive_sequence_magnitude(sequence_component_List[2],
                                                                                         times=40)
            Acu4100_Voltage_Positive_Angle = read_voltage_positive_sequence_angle(sequence_component_List[3], times=40)
            Acu4100_Voltage_Negative_Sequence = read_voltage_negative_sequence_magnitude(sequence_component_List[4],
                                                                                         times=40)
            Acu4100_Voltage_Negative_Angle = read_voltage_negative_sequence_angle(sequence_component_List[5], times=40)
            VUF = read_voltage_unbalance_factor_magnitude(sequence_component_List[6], times=40)
            Acu4100_Sequence_list = []
            Acu4100_Sequence_list.extend(
                [Acu4100_Voltage_Zero_Sequence, Acu4100_Voltage_Zero_Sequence_Angle, Acu4100_Voltage_Positive_Sequence,
                 Acu4100_Voltage_Positive_Angle, Acu4100_Voltage_Negative_Sequence, Acu4100_Voltage_Negative_Angle])
            sheet.write(i + 1, 7, f'{Acu4100_Voltage_Zero_Sequence}')
            sheet.write(i + 1, 8, f'{Acu4100_Voltage_Zero_Sequence_Angle}')
            sheet.write(i + 1, 9, f'{Acu4100_Voltage_Positive_Sequence}')
            sheet.write(i + 1, 10, f'{Acu4100_Voltage_Positive_Angle}')
            sheet.write(i + 1, 11, f'{Acu4100_Voltage_Negative_Sequence}')
            sheet.write(i + 1, 12, f'{Acu4100_Voltage_Negative_Angle}')
            sheet.write(i + 1, 13, f'{VUF}')
            for j in range(len(sequence_component_List)):
                if j < 6:
                    if sequence_component_List[j] == 0 and Acu4100_Sequence_list[j] <= 0.15:
                        sheet.write(i + 1, 14, f'Passed')
                        continue
                    elif sequence_component_List[j] * 0.95 <= Acu4100_Sequence_list[j] <= sequence_component_List[
                        j] * 1.05:
                        sheet.write(i + 1, 14, f'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 14, f'Failed')
                        sheet.write(i + 1, 15, f'标准值:{sequence_component_List}')
                        sheet.write(i + 1, 16, f'请检查:{Acu4100_Sequence_list[j]}值')
                        break
                else:
                    if sequence_component_List[6] * 0.99 <= VUF <= sequence_component_List[6] * 1.01:
                        sheet.write(i + 1, 14, f'Passed')
                        sheet.write(i + 1, 15, f'标准值:{sequence_component_List}')
                    else:
                        sheet.write(i + 1, 14, f'Failed')
                        sheet.write(i + 1, 15, f'标准值:{sequence_component_List}')
                        sheet.write(i + 1, 16, f'请检查VUF:{VUF}值')
        if Sequence_Component_list[i + 1][7] != 'null' and Sequence_Component_list[i + 1][8] == '用户序分量':
            set_ac(120, 240, 0, Sequence_Component_list[i + 1][6], Sequence_Component_list[i + 1][5],
                   Sequence_Component_list[i + 1][4], 50, 50, 50, Sequence_Component_list[i + 1][3],
                   Sequence_Component_list[i + 1][2], Sequence_Component_list[i + 1][1], 50)
            set_service_configuration(4)
            sequence_component_List = sequence_component_calculation(Sequence_Component_list[i + 1][1],
                                                                     Sequence_Component_list[i + 1][2],
                                                                     Sequence_Component_list[i + 1][3],
                                                                     Sequence_Component_list[i + 1][4],
                                                                     Sequence_Component_list[i + 1][5],
                                                                     Sequence_Component_list[i + 1][6])
            Acu4100_Voltage_Zero_Sequence = read_user_channel_1_current_zero_sequence_magnitude(
                sequence_component_List[0], times=40)
            Acu4100_Voltage_Zero_Sequence_Angle = read_user_channel_1_current_zero_sequence_angle(
                sequence_component_List[1], times=40)
            Acu4100_Voltage_Positive_Sequence = read_user_channel_1_current_positive_sequence_magnitude(
                sequence_component_List[2], times=40)
            Acu4100_Voltage_Positive_Angle = read_user_channel_1_current_positive_sequence_angle(
                sequence_component_List[3], times=40)
            Acu4100_Voltage_Negative_Sequence = read_user_channel_1_current_negative_sequence_magnitude(
                sequence_component_List[4], times=40)
            Acu4100_Voltage_Negative_Angle = read_user_channel_1_current_negative_sequence_angle(
                sequence_component_List[5], times=40)
            CUF = read_user_channel_1_current_unbalance_factor_magnitude(sequence_component_List[6], times=40)
            sheet.write(i + 1, 7, f'{Acu4100_Voltage_Zero_Sequence}')
            sheet.write(i + 1, 8, f'{Acu4100_Voltage_Zero_Sequence_Angle}')
            sheet.write(i + 1, 9, f'{Acu4100_Voltage_Positive_Sequence}')
            sheet.write(i + 1, 10, f'{Acu4100_Voltage_Positive_Angle}')
            sheet.write(i + 1, 11, f'{Acu4100_Voltage_Negative_Sequence}')
            sheet.write(i + 1, 12, f'{Acu4100_Voltage_Negative_Angle}')
            sheet.write(i + 1, 13, f'CUF:{CUF}')
            Acu4100_Sequence_list = []
            Acu4100_Sequence_list.extend(
                [Acu4100_Voltage_Zero_Sequence, Acu4100_Voltage_Zero_Sequence_Angle, Acu4100_Voltage_Positive_Sequence,
                 Acu4100_Voltage_Positive_Angle, Acu4100_Voltage_Negative_Sequence, Acu4100_Voltage_Negative_Angle])
            for j in range(len(sequence_component_List)):
                if j < 6:
                    if sequence_component_List[j] == 0 and Acu4100_Sequence_list[j] <= 0.15:
                        sheet.write(i + 1, 14, f'Passed')
                        continue
                    elif sequence_component_List[j] * 0.95 <= Acu4100_Sequence_list[j] <= sequence_component_List[
                        j] * 1.05:
                        sheet.write(i + 1, 14, f'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 14, f'Failed')
                        sheet.write(i + 1, 15, f'标准值:{sequence_component_List}')
                        sheet.write(i + 1, 16, f'请检查:{Acu4100_Sequence_list[j]}值')
                        break
                else:
                    if sequence_component_List[6] * 0.99 <= CUF <= sequence_component_List[6] * 1.01:
                        sheet.write(i + 1, 14, f'Passed')
                        sheet.write(i + 1, 15, f'标准值:{sequence_component_List}')
                    else:
                        sheet.write(i + 1, 14, f'Failed')
                        sheet.write(i + 1, 15, f'标准值:{sequence_component_List}')
                        sheet.write(i + 1, 16, f'请检查CUF:{CUF}值')
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def energy_2e_3w_delta_measure():
    Energy_E2_3W_Delta_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Energy_5A_333mV_CT1')
    print(Energy_E2_3W_Delta_list)
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Energy_E2_3W_Delta_measure', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Va输入值')
    sheet.write(0, 2, 'Vb输入值')
    sheet.write(0, 3, 'Vc输入值')

    sheet.write(0, 4, 'Ia输入值')
    sheet.write(0, 5, 'Ib输入值')
    sheet.write(0, 6, 'Ic输入值')

    sheet.write(0, 7, 'Va_ang输入值')
    sheet.write(0, 8, 'Vb_ang输入值')
    sheet.write(0, 9, 'Vc_ang输入值')

    sheet.write(0, 10, 'Ia_ang输入值')
    sheet.write(0, 11, 'Ib_ang输入值')
    sheet.write(0, 12, 'Ic_ang输入值')

    sheet.write(0, 13, '等待时间(min)')

    sheet.write(0, 14, 'System_P_E_import')
    sheet.write(0, 15, 'System_P_E_export')
    sheet.write(0, 16, 'System_P_E_net')
    sheet.write(0, 17, 'System_P_E_total')
    sheet.write(0, 18, 'System_Q_E_import')
    sheet.write(0, 19, 'System_Q_E_export')
    sheet.write(0, 20, 'System_Q_E_net')
    sheet.write(0, 21, 'System_Q_E_total')
    sheet.write(0, 22, 'System_S_E')

    sheet.write(0, 23, 'User_1_P_E_import')
    sheet.write(0, 24, 'User_1_P_E_export')
    sheet.write(0, 25, 'User_1_P_E_net')
    sheet.write(0, 26, 'User_1_P_E_total')
    sheet.write(0, 27, 'User_1_Q_E_import')
    sheet.write(0, 28, 'User_1_Q_E_export')
    sheet.write(0, 29, 'User_1_Q_E_net')
    sheet.write(0, 30, 'User_1_Q_E_total')
    sheet.write(0, 31, 'User_1_S_E')

    sheet.write(0, 32, '测试结果')
    for i in range(len(Energy_E2_3W_Delta_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(Energy_E2_3W_Delta_list[i]))
            print('测试进度:{}'.format(Energy_E2_3W_Delta_list[i]))
        else:
            logging.info(
                '测试进度:{},执行时间:{}'.format(Energy_E2_3W_Delta_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(Energy_E2_3W_Delta_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(Energy_E2_3W_Delta_list) - 1:
            break
        print(Energy_E2_3W_Delta_list[i + 1])
        print(Energy_E2_3W_Delta_list[i + 1][14])
        sheet.write(i + 1, 0, Energy_E2_3W_Delta_list[i + 1][0])
        sheet.write(i + 1, 1, Energy_E2_3W_Delta_list[i + 1][1])
        sheet.write(i + 1, 2, Energy_E2_3W_Delta_list[i + 1][2])
        sheet.write(i + 1, 3, Energy_E2_3W_Delta_list[i + 1][3])
        sheet.write(i + 1, 4, Energy_E2_3W_Delta_list[i + 1][4])
        sheet.write(i + 1, 5, Energy_E2_3W_Delta_list[i + 1][5])
        sheet.write(i + 1, 6, Energy_E2_3W_Delta_list[i + 1][6])
        sheet.write(i + 1, 7, Energy_E2_3W_Delta_list[i + 1][7])
        sheet.write(i + 1, 8, Energy_E2_3W_Delta_list[i + 1][8])
        sheet.write(i + 1, 9, Energy_E2_3W_Delta_list[i + 1][9])
        sheet.write(i + 1, 10, Energy_E2_3W_Delta_list[i + 1][10])
        sheet.write(i + 1, 11, Energy_E2_3W_Delta_list[i + 1][11])
        sheet.write(i + 1, 12, Energy_E2_3W_Delta_list[i + 1][12])
        sheet.write(i + 1, 13, Energy_E2_3W_Delta_list[i + 1][13])
        if Energy_E2_3W_Delta_list[i + 1][15] == '2E3WDelta' and Energy_E2_3W_Delta_list[i + 1][14] != 'null':
            set_service_configuration(2)
            time.sleep(1)
            set_ac(Energy_E2_3W_Delta_list[i + 1][9], Energy_E2_3W_Delta_list[i + 1][8],
                   Energy_E2_3W_Delta_list[i + 1][7],
                   Energy_E2_3W_Delta_list[i + 1][12], Energy_E2_3W_Delta_list[i + 1][11],
                   Energy_E2_3W_Delta_list[i + 1][10], Energy_E2_3W_Delta_list[i + 1][3],
                   Energy_E2_3W_Delta_list[i + 1][2], Energy_E2_3W_Delta_list[i + 1][1],
                   Energy_E2_3W_Delta_list[i + 1][6],
                   Energy_E2_3W_Delta_list[i + 1][5], Energy_E2_3W_Delta_list[i + 1][4], 50)
            set_clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_E2_3W_Delta_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_E2_3W_Delta_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)

            line_to_line_voltage = line_to_line_voltage_calculate(Energy_E2_3W_Delta_list[i + 1][1],
                                                                  Energy_E2_3W_Delta_list[i + 1][2],
                                                                  Energy_E2_3W_Delta_list[i + 1][3],
                                                                  Energy_E2_3W_Delta_list[i + 1][7],
                                                                  Energy_E2_3W_Delta_list[i + 1][8],
                                                                  Energy_E2_3W_Delta_list[i + 1][9])
            Vab = line_to_line_voltage[0]
            Vbc = line_to_line_voltage[1]
            Vca = line_to_line_voltage[2]

            input_Power_standard = e2_3w_delta_power_standard_value_new(Vab, Vbc, Vca,
                                                                        Energy_E2_3W_Delta_list[i + 1][4],
                                                                        Energy_E2_3W_Delta_list[i + 1][5],
                                                                        Energy_E2_3W_Delta_list[i + 1][6],
                                                                        Energy_E2_3W_Delta_list[i + 1][7],
                                                                        Energy_E2_3W_Delta_list[i + 1][10],
                                                                        Energy_E2_3W_Delta_list[i + 1][8],
                                                                        Energy_E2_3W_Delta_list[i + 1][11],
                                                                        Energy_E2_3W_Delta_list[i + 1][9],
                                                                        Energy_E2_3W_Delta_list[i + 1][12])
            User_P = input_Power_standard[4][0]
            User_Q = input_Power_standard[4][1]
            User_S = input_Power_standard[4][2]
            if User_P > 0:
                User_P_E_import = abs(User_P * (Energy_E2_3W_Delta_list[i + 1][13] / 60))
                User_P_E_export = 0
            else:
                User_P_E_import = 0
                User_P_E_export = abs(User_P * (Energy_E2_3W_Delta_list[i + 1][13] / 60))
            User_P_E_net = User_P_E_import - User_P_E_export
            User_P_E_total = User_P_E_import + User_P_E_export

            if User_Q > 0:
                User_Q_E_import = abs(User_Q * (Energy_E2_3W_Delta_list[i + 1][13] / 60))
                User_Q_E_export = 0
            else:
                User_Q_E_import = 0
                User_Q_E_export = abs(User_Q * (Energy_E2_3W_Delta_list[i + 1][13] / 60))
            User_Q_E_net = User_Q_E_import - User_Q_E_export
            User_Q_E_total = User_Q_E_import + User_Q_E_export

            User_S_E = User_S * (Energy_E2_3W_Delta_list[i + 1][13] / 60)

            System_Energy_list = read_system_energy()
            User_Channel_1_Energy = read_user_channel_1_energy()
            Read_Energy_list = System_Energy_list + User_Channel_1_Energy

            User_power_list = [User_P_E_import, User_P_E_export, User_P_E_net, User_P_E_total, User_Q_E_import,
                               User_Q_E_export, User_Q_E_net, User_Q_E_total, User_S_E]

            Energy_scale_list = []
            for n in range(len(User_power_list)):
                if User_power_list[n] != 0:
                    Energy_scale = abs(
                        (System_Energy_list[n] - User_power_list[n]) / User_power_list[n])
                    Energy_scale_list.append(Energy_scale)
                else:
                    Energy_scale = 0
                    Energy_scale_list.append(Energy_scale)
            for m in range(len(User_power_list)):
                if User_power_list[m] != 0:
                    Energy_scale = abs(
                        (User_Channel_1_Energy[m] - User_power_list[m]) / User_power_list[m])
                    Energy_scale_list.append(Energy_scale)
                else:
                    Energy_scale = 0
                    Energy_scale_list.append(Energy_scale)

            Read_Energy_scale_list = (Read_Energy_list, Energy_scale_list)
            for j in range(len(Read_Energy_scale_list[0])):
                if Read_Energy_scale_list[1][j] != 'null':
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},{Read_Energy_scale_list[1][j]:.3%}')
                else:
                    sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},null')
            for k in range(len(Read_Energy_scale_list[1])):
                print(Read_Energy_scale_list[1])
                print(i)
                print(Energy_E2_3W_Delta_list[i + 1][14])
                if Read_Energy_scale_list[1][k] != 'null' and Read_Energy_scale_list[1][k] * 100 <= \
                        Energy_E2_3W_Delta_list[i + 1][14]:
                    sheet.write(i + 1, 32, f'Passed')
                    continue
                else:
                    sheet.write(i + 1, 32, f'Failed')
                    sheet.write(i + 1, 32, f'{k + 15}列精度不达标或null')
                    break
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


if __name__ == '__main__':
    print('====================Precision Measure Start====================')
    print('======================{}======================'.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    start_time = time.time()
    # my_workbook = xlwt.Workbook()
    switch_device_screen_interface(0x01)
    time.sleep(5)
    set_gear_switching_mode('00000000')
    time.sleep(5)
    # frequency_precision_measure()
    # line_to_neutral_voltage_precision_measure()
    # line_to_line_voltage_precision_measure()
    # current_5a_333mv_ct_precision_measure()
    # current_20a_100ma_ct_precision_measure()
    # power_5a_333mv_ct_precision_measure()
    # power_20a_100ma_ct_precision_measure()
    # phase_voltage_angle_precision_measure()
    # input1_current_angle_precision_measure()
    # load_nature_measure()
    energy_5a_333mv_ct_measure()
    # energy_20a_100ma_ct_measure()
    # sequence_component_precision_measure()
    # energy_2e_3w_delta_measure()
    ModbusClient.close()
    # my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))
    ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
    time.sleep(5)
    switch_device_screen_interface(0x00)
    print('====================测试总耗时:{}===================='.format(time.time() - start_time))
    print('====================={}====================='.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    print('====================Precision Measure End====================')
