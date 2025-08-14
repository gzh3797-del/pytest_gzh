#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:Acuvim2v3_Precision_Measure.py
功能描述:能量测试
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""

import os
import time

import openpyxl

from comm.source_control import switch_device_screen_interface, set_gear_switching_mode, set_ac
from test_case.AcuvimSeries.acuvimseries_modbus_get import HandleMemory
from test_case.Acuvim2v3.testcase_names import ENERGY_MEASURE_COLUMNS
from tools.excel_operate import data_read
from tools.log import Log

Log(str(__file__).split("\\")[-1])

filepath = r'./test_case/Acuvim2v3/acuvim2v3_test_case_0722.xlsx'
sheet_name_mV = "test_case_mV"
sheet_name_mA = "test_case_mA"
sheet_name_5A = "test_case_5A"

save_filedir = rf"./energy_measure_{time.strftime('%Y%m%d')}"
if not os.path.exists(save_filedir):
    os.makedirs(save_filedir, exist_ok=True)


def up_source_ac():
    ret = set_ac(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    return ret


class EnergyMeasure:
    def __init__(self):
        self.handle_memory = HandleMemory(slave_id=1)

    def select_test_type(self, test_type, wire_type):
        """
        选择测试数据类型
        :param test_type: 0:mV, 1:mA, 2:5A
        :param wire_type: 0:1e2w1p, 1:2e3w1p, 2:2e3wN, 3:2e3wD, 4:3e3wD, 5:2.5e4wY, 6:3e4wY,
        :return:
        """
        if test_type == 0:
            self.select_wire_type(filepath, sheet_name_mV, wire_type)
        elif test_type == 1:
            self.select_wire_type(filepath, sheet_name_mA, wire_type)
        else:
            self.select_wire_type(filepath, sheet_name_5A, wire_type)

    def set_wire_type(self, voltage_wire_value, current_wire_value):
        self.handle_memory.set_wire_mode_by_voltage(voltage_wire_mode=voltage_wire_value)
        self.handle_memory.set_wire_mode_by_current(current_wire_mode=current_wire_value)

    @staticmethod
    def get_input_list_by_wire_type(source_input_list, wire_type):
        input_list = []
        input_list.append(source_input_list[0])
        if wire_type == 6:
            for i in range(1, len(source_input_list)):
                if source_input_list[i][15] == "3E4wY":
                    input_list.append(source_input_list[i])
        elif wire_type == 5:
            for i in range(1, len(source_input_list)):
                if source_input_list[i][15] == "2.5E4wY":
                    input_list.append(source_input_list[i])
        elif wire_type == 4:
            for i in range(1, len(source_input_list)):
                if source_input_list[i][15] == "3E3wD":
                    input_list.append(source_input_list[i])
        elif wire_type == 3:
            for i in range(1, len(source_input_list)):
                if source_input_list[i][15] == "2E3wD":
                    input_list.append(source_input_list[i])
        elif wire_type == 2:
            for i in range(1, len(source_input_list)):
                if source_input_list[i][15] == "2E3wN":
                    input_list.append(source_input_list[i])
        elif wire_type == 1:
            for i in range(1, len(source_input_list)):
                if source_input_list[i][15] == "2E3w1p":
                    input_list.append(source_input_list[i])
        elif wire_type == 0:
            for i in range(1, len(source_input_list)):
                if source_input_list[i][15] == "1E2w1p":
                    input_list.append(source_input_list[i])
        return input_list

    def select_wire_type(self, filepath, sheet_name, wire_type):
        """
        选择接线方式
        :param filepath: 测试文件名
        :param sheet_name: sheet名
        :param wire_type: 0:1e2w1p, 1:2e3w1p, 2:2e3wN, 3:2e3wD, 4:3e3wD, 5:2.5e4wY, 6:3e4wY,
        :return:
        # 3 element 4 wire Wye       3LN     3CT    6
        # 2.5 element 4 wire Wye     3LN-2.5 3CT    5
        # 3 element 3 wire Delta     3LL     3CT    4
        # 2 element 3 wire Delta     2LL     3CT    3
        # 2 element 3 wire network   2LL     2CT    2
        # 2 element 3 wire 1 phase   1LL     2CT    1
        # 1 element 2 wire           1LN     1CT    0
        """
        if wire_type == 0:
            self.set_wire_type(voltage_wire_value=1, current_wire_value=1)
            source_input_list = data_read(filepath, sheet_name)
            input_list = self.get_input_list_by_wire_type(source_input_list, wire_type)
            # fast_precision_measure_by_1E2W1P(input_list)
        elif wire_type == 1:
            self.set_wire_type(voltage_wire_value=4, current_wire_value=2)
            source_input_list = data_read(filepath, sheet_name)
            input_list = self.get_input_list_by_wire_type(source_input_list, wire_type)
            # fast_precision_measure_by_2E3W1P(input_list)
        elif wire_type == 2:
            self.set_wire_type(voltage_wire_value=2, current_wire_value=2)
            source_input_list = data_read(filepath, sheet_name)
            input_list = self.get_input_list_by_wire_type(source_input_list, wire_type)
            # fast_precision_measure_by_2E3WN(input_list)
        elif wire_type == 3:
            self.set_wire_type(voltage_wire_value=2, current_wire_value=0)
            source_input_list = data_read(filepath, sheet_name)
            input_list = self.get_input_list_by_wire_type(source_input_list, wire_type)
            # fast_precision_measure_by_2E3WD(input_list)
        elif wire_type == 4:
            self.set_wire_type(voltage_wire_value=3, current_wire_value=0)
            source_input_list = data_read(filepath, sheet_name)
            input_list = self.get_input_list_by_wire_type(source_input_list, wire_type)
            # fast_precision_measure_by_3E3WD(input_list)
        elif wire_type == 5:
            self.set_wire_type(voltage_wire_value=5, current_wire_value=0)
            source_input_list = data_read(filepath, sheet_name)
            input_list = self.get_input_list_by_wire_type(source_input_list, wire_type)
            # fast_precision_measure_by_2p5E4WY(input_list)
        else:
            self.set_wire_type(voltage_wire_value=0, current_wire_value=0)
            source_input_list = data_read(filepath, sheet_name)
            input_list = self.get_input_list_by_wire_type(source_input_list, wire_type)
            filename = f"Precision_Measure_3E4WY_{time.strftime('%Y%m%d%H%M%S')}.xlsx"
            file_path = os.path.join(str({save_filedir}), str(filename))
            self.energy_measure_by_3e4wy(file_path, input_list)

    @staticmethod
    def write_excel_title(ws):
        for i in range(len(ENERGY_MEASURE_COLUMNS)):
            j = i + 1
            ws.cell(1, j, f'{ENERGY_MEASURE_COLUMNS[i]}')

    def energy_measure_by_3e4wy(self, file_path, input_list):
        # 读取工作簿
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        self.write_excel_title(ws)
        pass
        # sheet.write(0, 0, '测试用例')
        # sheet.write(0, 1, 'Va输入值')
        # sheet.write(0, 2, 'Vb输入值')
        # sheet.write(0, 3, 'Vc输入值')
        # sheet.write(0, 4, 'Ia输入值')
        # sheet.write(0, 5, 'Ib输入值')
        # sheet.write(0, 6, 'Ic输入值')
        # sheet.write(0, 7, 'Va_ang输入值')
        # sheet.write(0, 8, 'Vb_ang输入值')
        # sheet.write(0, 9, 'Vc_ang输入值')
        # sheet.write(0, 10, 'Ia_ang输入值')
        # sheet.write(0, 11, 'Ib_ang输入值')
        # sheet.write(0, 12, 'Ic_ang输入值')
        # sheet.write(0, 13, '等待时间(min)')
        # sheet.write(0, 14, 'Phase_A_P_E_import')
        # sheet.write(0, 15, 'Phase_A_P_E_export')
        # sheet.write(0, 16, 'Phase_A_P_E_net')
        # sheet.write(0, 17, 'Phase_A_P_E_total')
        # sheet.write(0, 18, 'Phase_A_Q_E_import')
        # sheet.write(0, 19, 'Phase_A_Q_E_export')
        # sheet.write(0, 20, 'Phase_A_Q_E_net')
        # sheet.write(0, 21, 'Phase_A_Q_E_total')
        # sheet.write(0, 22, 'Phase_A_S_E')
        # sheet.write(0, 23, 'Phase_B_P_E_import')
        # sheet.write(0, 24, 'Phase_B_P_E_export')
        # sheet.write(0, 25, 'Phase_B_P_E_net')
        # sheet.write(0, 26, 'Phase_B_P_E_total')
        # sheet.write(0, 27, 'Phase_B_Q_E_import')
        # sheet.write(0, 28, 'Phase_B_Q_E_export')
        # sheet.write(0, 29, 'Phase_B_Q_E_net')
        # sheet.write(0, 30, 'Phase_B_Q_E_total')
        # sheet.write(0, 31, 'Phase_B_S_E')
        # sheet.write(0, 32, 'Phase_C_P_E_import')
        # sheet.write(0, 33, 'Phase_C_P_E_export')
        # sheet.write(0, 34, 'Phase_C_P_E_net')
        # sheet.write(0, 35, 'Phase_C_P_E_total')
        # sheet.write(0, 36, 'Phase_C_Q_E_import')
        # sheet.write(0, 37, 'Phase_C_Q_E_export')
        # sheet.write(0, 38, 'Phase_C_Q_E_net')
        # sheet.write(0, 39, 'Phase_C_Q_E_total')
        # sheet.write(0, 40, 'Phase_C_S_E')
        # sheet.write(0, 41, 'System_P_E_import')
        # sheet.write(0, 42, 'System_P_E_export')
        # sheet.write(0, 43, 'System_P_E_net')
        # sheet.write(0, 44, 'System_P_E_total')
        # sheet.write(0, 45, 'System_Q_E_import')
        # sheet.write(0, 46, 'System_Q_E_export')
        # sheet.write(0, 47, 'System_Q_E_net')
        # sheet.write(0, 48, 'System_Q_E_total')
        # sheet.write(0, 49, 'System_S_E')
        # sheet.write(0, 50, 'Input_1_P_E_import')
        # sheet.write(0, 51, 'Input_1_P_E_export')
        # sheet.write(0, 52, 'Input_1_P_E_net')
        # sheet.write(0, 53, 'Input_1_P_E_total')
        # sheet.write(0, 54, 'Input_1_Q_E_import')
        # sheet.write(0, 55, 'Input_1_Q_E_export')
        # sheet.write(0, 56, 'Input_1_Q_E_net')
        # sheet.write(0, 57, 'Input_1_Q_E_total')
        # sheet.write(0, 58, 'Input_1_S_E')
        # sheet.write(0, 59, 'Input_2_P_E_import')
        # sheet.write(0, 60, 'Input_2_P_E_export')
        # sheet.write(0, 61, 'Input_2_P_E_net')
        # sheet.write(0, 62, 'Input_2_P_E_total')
        # sheet.write(0, 63, 'Input_2_Q_E_import')
        # sheet.write(0, 64, 'Input_2_Q_E_export')
        # sheet.write(0, 65, 'Input_2_Q_E_net')
        # sheet.write(0, 66, 'Input_2_Q_E_total')
        # sheet.write(0, 67, 'Input_2_S_E')
        # sheet.write(0, 68, 'Input_3_P_E_import')
        # sheet.write(0, 69, 'Input_3_P_E_export')
        # sheet.write(0, 70, 'Input_3_P_E_net')
        # sheet.write(0, 71, 'Input_3_P_E_total')
        # sheet.write(0, 72, 'Input_3_Q_E_import')
        # sheet.write(0, 73, 'Input_3_Q_E_export')
        # sheet.write(0, 74, 'Input_3_Q_E_net')
        # sheet.write(0, 75, 'Input_3_Q_E_total')
        # sheet.write(0, 76, 'Input_3_S_E')
        # sheet.write(0, 77, 'User_1_P_E_import')
        # sheet.write(0, 78, 'User_1_P_E_export')
        # sheet.write(0, 79, 'User_1_P_E_net')
        # sheet.write(0, 80, 'User_1_P_E_total')
        # sheet.write(0, 81, 'User_1_Q_E_import')
        # sheet.write(0, 82, 'User_1_Q_E_export')
        # sheet.write(0, 83, 'User_1_Q_E_net')
        # sheet.write(0, 84, 'User_1_Q_E_total')
        # sheet.write(0, 85, 'User_1_S_E')
        # sheet.write(0, 86, 'User_2_P_E_import')
        # sheet.write(0, 87, 'User_2_P_E_export')
        # sheet.write(0, 88, 'User_2_P_E_net')
        # sheet.write(0, 89, 'User_2_P_E_total')
        # sheet.write(0, 90, 'User_2_Q_E_import')
        # sheet.write(0, 91, 'User_2_Q_E_export')
        # sheet.write(0, 92, 'User_2_Q_E_net')
        # sheet.write(0, 93, 'User_2_Q_E_total')
        # sheet.write(0, 94, 'User_2_S_E')
        # sheet.write(0, 95, 'User_3_P_E_import')
        # sheet.write(0, 96, 'User_3_P_E_export')
        # sheet.write(0, 97, 'User_3_P_E_net')
        # sheet.write(0, 98, 'User_3_P_E_total')
        # sheet.write(0, 99, 'User_3_Q_E_import')
        # sheet.write(0, 100, 'User_3_Q_E_export')
        # sheet.write(0, 101, 'User_3_Q_E_net')
        # sheet.write(0, 102, 'User_3_Q_E_total')
        # sheet.write(0, 103, 'User_3_S_E')
        #
        # sheet.write(0, 104, '测试结果')
        # for i in range(len(Energy_5A_333mV_CT_list)):
        #     if i == 0:
        #         logging.info('测试进度:{}'.format(Energy_5A_333mV_CT_list[i]))
        #         print('测试进度:{}'.format(Energy_5A_333mV_CT_list[i]))
        #     else:
        #         logging.info(
        #             '测试进度:{},执行时间:{}'.format(Energy_5A_333mV_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        #         print('测试进度:{},执行时间:{}'.format(Energy_5A_333mV_CT_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        #     if i == len(Energy_5A_333mV_CT_list) - 1:
        #         break
        #     sheet.write(i + 1, 0, Energy_5A_333mV_CT_list[i + 1][0])
        #     sheet.write(i + 1, 1, Energy_5A_333mV_CT_list[i + 1][1])
        #     sheet.write(i + 1, 2, Energy_5A_333mV_CT_list[i + 1][2])
        #     sheet.write(i + 1, 3, Energy_5A_333mV_CT_list[i + 1][3])
        #     sheet.write(i + 1, 4, Energy_5A_333mV_CT_list[i + 1][4])
        #     sheet.write(i + 1, 5, Energy_5A_333mV_CT_list[i + 1][5])
        #     sheet.write(i + 1, 6, Energy_5A_333mV_CT_list[i + 1][6])
        #     sheet.write(i + 1, 7, Energy_5A_333mV_CT_list[i + 1][7])
        #     sheet.write(i + 1, 8, Energy_5A_333mV_CT_list[i + 1][8])
        #     sheet.write(i + 1, 9, Energy_5A_333mV_CT_list[i + 1][9])
        #     sheet.write(i + 1, 10, Energy_5A_333mV_CT_list[i + 1][10])
        #     sheet.write(i + 1, 11, Energy_5A_333mV_CT_list[i + 1][11])
        #     sheet.write(i + 1, 12, Energy_5A_333mV_CT_list[i + 1][12])
        #     sheet.write(i + 1, 13, Energy_5A_333mV_CT_list[i + 1][13])
        #     if Energy_5A_333mV_CT_list[i + 1][15] == '3E3p4w' and Energy_5A_333mV_CT_list[i + 1][14] != 'null':
        #         Set_Service_Configuration(4)
        #         time.sleep(1)
        #         Set_Device_Reboot(1)
        #         time.sleep(15)
        #         set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
        #                Energy_5A_333mV_CT_list[i + 1][7],
        #                Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
        #                Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
        #                Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
        #                Energy_5A_333mV_CT_list[i + 1][6],
        #                Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
        #         Set_Clear_energy(1)
        #         thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
        #         thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
        #         thread_a.start()
        #         thread_b.start()
        #         thread_a.join()
        #         thread_b.join()
        #         ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
        #         Read_Energy_scale_list = Read_Energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
        #                                                        Energy_5A_333mV_CT_list[i + 1][2],
        #                                                        Energy_5A_333mV_CT_list[i + 1][3],
        #                                                        Energy_5A_333mV_CT_list[i + 1][4],
        #                                                        Energy_5A_333mV_CT_list[i + 1][5],
        #                                                        Energy_5A_333mV_CT_list[i + 1][6],
        #                                                        Energy_5A_333mV_CT_list[i + 1][7],
        #                                                        Energy_5A_333mV_CT_list[i + 1][8],
        #                                                        Energy_5A_333mV_CT_list[i + 1][9],
        #                                                        Energy_5A_333mV_CT_list[i + 1][10],
        #                                                        Energy_5A_333mV_CT_list[i + 1][11],
        #                                                        Energy_5A_333mV_CT_list[i + 1][12],
        #                                                        Energy_5A_333mV_CT_list[i + 1][13],
        #                                                        Energy_5A_333mV_CT_list[i + 1][15])
        #         for j in range(len(Read_Energy_scale_list[0])):
        #             if Read_Energy_scale_list[1][j] != 'null':
        #                 sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},{Read_Energy_scale_list[1][j]:.3%}')
        #             else:
        #                 sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},null')
        #         for k in range(len(Read_Energy_scale_list[1])):
        #             if Read_Energy_scale_list[1][k] != 'null' and Read_Energy_scale_list[1][k] * 100 <= \
        #                     Energy_5A_333mV_CT_list[i + 1][14]:
        #                 sheet.write(i + 1, 104, f'Passed')
        #                 continue
        #             else:
        #                 sheet.write(i + 1, 104, f'Failed')
        #                 sheet.write(i + 1, 105, f'{k + 15}列精度不达标或null')
        #                 break
        #     if Energy_5A_333mV_CT_list[i + 1][15] == '1E1p2w' and Energy_5A_333mV_CT_list[i + 1][14] != 'null':
        #         Set_Service_Configuration(0)
        #         time.sleep(1)
        #         # Set_channle2_voltage_assignment(1)
        #         # Set_channle3_voltage_assignment(2)
        #         Set_Device_Reboot(1)
        #         time.sleep(15)
        #         set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
        #                Energy_5A_333mV_CT_list[i + 1][7],
        #                Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
        #                Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
        #                Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
        #                Energy_5A_333mV_CT_list[i + 1][6],
        #                Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
        #         Set_Clear_energy(1)
        #         thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
        #         thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
        #         thread_a.start()
        #         thread_b.start()
        #         thread_a.join()
        #         thread_b.join()
        #         ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
        #         Read_Energy_scale_list = Read_Energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
        #                                                        Energy_5A_333mV_CT_list[i + 1][2],
        #                                                        Energy_5A_333mV_CT_list[i + 1][3],
        #                                                        Energy_5A_333mV_CT_list[i + 1][4],
        #                                                        Energy_5A_333mV_CT_list[i + 1][5],
        #                                                        Energy_5A_333mV_CT_list[i + 1][6],
        #                                                        Energy_5A_333mV_CT_list[i + 1][7],
        #                                                        Energy_5A_333mV_CT_list[i + 1][8],
        #                                                        Energy_5A_333mV_CT_list[i + 1][9],
        #                                                        Energy_5A_333mV_CT_list[i + 1][10],
        #                                                        Energy_5A_333mV_CT_list[i + 1][11],
        #                                                        Energy_5A_333mV_CT_list[i + 1][12],
        #                                                        Energy_5A_333mV_CT_list[i + 1][13],
        #                                                        Energy_5A_333mV_CT_list[i + 1][15])
        #         for j in range(len(Read_Energy_scale_list[0])):
        #             if Read_Energy_scale_list[1][j] != 'null':
        #                 sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},{Read_Energy_scale_list[1][j]:.3%}')
        #             else:
        #                 sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]},null')
        #         for k in range(len(Read_Energy_scale_list[1])):
        #             if Read_Energy_scale_list[1][k] != 'null' and Read_Energy_scale_list[1][k] * 100 <= \
        #                     Energy_5A_333mV_CT_list[i + 1][14]:
        #                 sheet.write(i + 1, 104, f'Passed')
        #                 continue
        #             else:
        #                 sheet.write(i + 1, 104, f'Failed')
        #                 sheet.write(i + 1, 105, f'{k + 15}列精度不达标或null')
        #                 break
        #     if Energy_5A_333mV_CT_list[i + 1][15] == '3E3p4w' and Energy_5A_333mV_CT_list[i + 1][14] == 'null':
        #         Set_Service_Configuration(4)
        #         time.sleep(1)
        #         Set_Device_Reboot(1)
        #         time.sleep(15)
        #         set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
        #                Energy_5A_333mV_CT_list[i + 1][7],
        #                Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
        #                Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
        #                Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
        #                Energy_5A_333mV_CT_list[i + 1][6],
        #                Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
        #         Set_Clear_energy(1)
        #         thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
        #         thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
        #         thread_a.start()
        #         thread_b.start()
        #         thread_a.join()
        #         thread_b.join()
        #         ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
        #         Read_Energy_scale_list = Read_Energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
        #                                                        Energy_5A_333mV_CT_list[i + 1][2],
        #                                                        Energy_5A_333mV_CT_list[i + 1][3],
        #                                                        Energy_5A_333mV_CT_list[i + 1][4],
        #                                                        Energy_5A_333mV_CT_list[i + 1][5],
        #                                                        Energy_5A_333mV_CT_list[i + 1][6],
        #                                                        Energy_5A_333mV_CT_list[i + 1][7],
        #                                                        Energy_5A_333mV_CT_list[i + 1][8],
        #                                                        Energy_5A_333mV_CT_list[i + 1][9],
        #                                                        Energy_5A_333mV_CT_list[i + 1][10],
        #                                                        Energy_5A_333mV_CT_list[i + 1][11],
        #                                                        Energy_5A_333mV_CT_list[i + 1][12],
        #                                                        Energy_5A_333mV_CT_list[i + 1][13],
        #                                                        Energy_5A_333mV_CT_list[i + 1][15])
        #         for j in range(len(Read_Energy_scale_list[0])):
        #             sheet.write(i + 1, j + 14, f'{Read_Energy_scale_list[0][j]}')
        #         for j in range(len(Read_Energy_scale_list[0])):
        #             if Energy_5A_333mV_CT_list[i + 1][4] < 0.005 and Energy_5A_333mV_CT_list[i + 1][5] < 0.005 and \
        #                     Energy_5A_333mV_CT_list[i + 1][6] < 0.005:
        #                 if Read_Energy_scale_list[0][j] == 0:
        #                     sheet.write(i + 1, 104, f'Passed')
        #                     continue
        #                 else:
        #                     sheet.write(i + 1, 104, f'Failed')
        #                     sheet.write(i + 1, 105, f'{j + 15}列能量数据预期为0')
        #                     break
        #             else:
        #                 if Read_Energy_scale_list[0][j] != 0:
        #                     sheet.write(i + 1, 104, f'Passed')
        #                     break
        #                 else:
        #                     sheet.write(i + 1, 104, f'Failed')
        #                     continue
        # my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


if __name__ == '__main__':
    print(f"====================Precision Measure Start====================")
    print(f"======================{time.strftime('%Y_%m_%d %H:%M:%S')}======================")
    start_time = time.time()
    switch_device_screen_interface(inter=0x01)  # 切换至交流界面
    set_gear_switching_mode('00000000')  # 档位切换归零
    time.sleep(5)
    energy_measure = EnergyMeasure()

    # mV测量
    # select_test_case(test_type=0, wire_type=6)

    # mA测量
    # select_test_case(test_type=1, wire_type=6)

    # 5A测量
    energy_measure.select_test_type(test_type=2, wire_type=6)

    # 关闭ModbusClient客户端连接
    energy_measure.handle_memory.modbus_client.close()
    up_source_ac()
    switch_device_screen_interface(inter=0x00)  # 切换至默认界面
    print(f"====================测试总耗时:{time.time() - start_time}====================")
    print(f"====================={time.strftime('%Y_%m_%d %H:%M:%S')}=====================")
    print(f"====================Precision Measure End====================")
