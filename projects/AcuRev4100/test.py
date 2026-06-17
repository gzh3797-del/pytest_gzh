import time

from AcuRev4100_modbus_get import *
import threading

# volt_cur_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'test_data')
Log(str(__file__).split("\\")[-1])
def Energy_5A_333mV_CT_measure():
    Energy_5A_333mV_CT_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Energy_5A_333mV_CT')
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
            Set_Service_Configuration(4)
            time.sleep(1)
            Set_Device_Reboot(1)
            time.sleep(15)
            set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
                   Energy_5A_333mV_CT_list[i + 1][7],
                   Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
                   Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
                   Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
                   Energy_5A_333mV_CT_list[i + 1][6],
                   Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
            Set_Clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = Read_Energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
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
            Set_Service_Configuration(0)
            time.sleep(1)
            # Set_channle2_voltage_assignment(1)
            # Set_channle3_voltage_assignment(2)
            Set_Device_Reboot(1)
            time.sleep(15)
            set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
                   Energy_5A_333mV_CT_list[i + 1][7],
                   Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
                   Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
                   Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
                   Energy_5A_333mV_CT_list[i + 1][6],
                   Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
            Set_Clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = Read_Energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
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
            Set_Service_Configuration(4)
            time.sleep(1)
            Set_Device_Reboot(1)
            time.sleep(15)
            set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
                   Energy_5A_333mV_CT_list[i + 1][7],
                   Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
                   Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
                   Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
                   Energy_5A_333mV_CT_list[i + 1][6],
                   Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
            Set_Clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = Read_Energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
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

def Energy_5A_333mV_CT_measure1():
    Energy_5A_333mV_CT_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Energy_5A_333mV_CT1')
    my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('Energy_5A_333mV_CT1', cell_overwrite_ok=True)
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
            Set_Service_Configuration(4)
            time.sleep(1)
            Set_Device_Reboot(1)
            time.sleep(15)
            set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
                   Energy_5A_333mV_CT_list[i + 1][7],
                   Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
                   Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
                   Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
                   Energy_5A_333mV_CT_list[i + 1][6],
                   Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
            Set_Clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = Read_Energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
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
            Set_Service_Configuration(0)
            time.sleep(1)
            # Set_channle2_voltage_assignment(1)
            # Set_channle3_voltage_assignment(2)
            Set_Device_Reboot(1)
            time.sleep(15)
            set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
                   Energy_5A_333mV_CT_list[i + 1][7],
                   Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
                   Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
                   Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
                   Energy_5A_333mV_CT_list[i + 1][6],
                   Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
            Set_Clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = Read_Energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
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
            Set_Service_Configuration(4)
            time.sleep(1)
            Set_Device_Reboot(1)
            time.sleep(15)
            set_ac(Energy_5A_333mV_CT_list[i + 1][9], Energy_5A_333mV_CT_list[i + 1][8],
                   Energy_5A_333mV_CT_list[i + 1][7],
                   Energy_5A_333mV_CT_list[i + 1][12], Energy_5A_333mV_CT_list[i + 1][11],
                   Energy_5A_333mV_CT_list[i + 1][10], Energy_5A_333mV_CT_list[i + 1][3],
                   Energy_5A_333mV_CT_list[i + 1][2], Energy_5A_333mV_CT_list[i + 1][1],
                   Energy_5A_333mV_CT_list[i + 1][6],
                   Energy_5A_333mV_CT_list[i + 1][5], Energy_5A_333mV_CT_list[i + 1][4], 50)
            Set_Clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_5A_333mV_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = Read_Energy_scale_new(Energy_5A_333mV_CT_list[i + 1][1],
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
def Energy_20A_100mA_CT_measure():
    Energy_20A_100mA_CT_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Energy_20A_100mA_CT')
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
            Set_Service_Configuration(4)
            time.sleep(1)
            Set_Device_Reboot(1)
            time.sleep(15)
            set_ac(Energy_20A_100mA_CT_list[i + 1][9], Energy_20A_100mA_CT_list[i + 1][8],
                   Energy_20A_100mA_CT_list[i + 1][7],
                   Energy_20A_100mA_CT_list[i + 1][12], Energy_20A_100mA_CT_list[i + 1][11],
                   Energy_20A_100mA_CT_list[i + 1][10], Energy_20A_100mA_CT_list[i + 1][3],
                   Energy_20A_100mA_CT_list[i + 1][2], Energy_20A_100mA_CT_list[i + 1][1],
                   Energy_20A_100mA_CT_list[i + 1][6],
                   Energy_20A_100mA_CT_list[i + 1][5], Energy_20A_100mA_CT_list[i + 1][4], 50)
            Set_Clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = Read_Energy_scale_new(Energy_20A_100mA_CT_list[i + 1][1],
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
            Set_Service_Configuration(0)
            time.sleep(1)
            # Set_channle2_voltage_assignment(1)
            # Set_channle3_voltage_assignment(2)
            Set_Device_Reboot(1)
            time.sleep(15)
            set_ac(Energy_20A_100mA_CT_list[i + 1][9], Energy_20A_100mA_CT_list[i + 1][8],
                   Energy_20A_100mA_CT_list[i + 1][7],
                   Energy_20A_100mA_CT_list[i + 1][12], Energy_20A_100mA_CT_list[i + 1][11],
                   Energy_20A_100mA_CT_list[i + 1][10], Energy_20A_100mA_CT_list[i + 1][3],
                   Energy_20A_100mA_CT_list[i + 1][2], Energy_20A_100mA_CT_list[i + 1][1],
                   Energy_20A_100mA_CT_list[i + 1][6],
                   Energy_20A_100mA_CT_list[i + 1][5], Energy_20A_100mA_CT_list[i + 1][4], 50)
            Set_Clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = Read_Energy_scale_new(Energy_20A_100mA_CT_list[i + 1][1],
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
            Set_Service_Configuration(4)
            time.sleep(1)
            Set_Device_Reboot(1)
            time.sleep(15)
            set_ac(Energy_20A_100mA_CT_list[i + 1][9], Energy_20A_100mA_CT_list[i + 1][8],
                   Energy_20A_100mA_CT_list[i + 1][7],
                   Energy_20A_100mA_CT_list[i + 1][12], Energy_20A_100mA_CT_list[i + 1][11],
                   Energy_20A_100mA_CT_list[i + 1][10], Energy_20A_100mA_CT_list[i + 1][3],
                   Energy_20A_100mA_CT_list[i + 1][2], Energy_20A_100mA_CT_list[i + 1][1],
                   Energy_20A_100mA_CT_list[i + 1][6],
                   Energy_20A_100mA_CT_list[i + 1][5], Energy_20A_100mA_CT_list[i + 1][4], 50)
            Set_Clear_energy(1)
            thread_a = threading.Thread(target=wait_minutes, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_b = threading.Thread(target=hold_rs485_connect, args=(Energy_20A_100mA_CT_list[i + 1][13],))
            thread_a.start()
            thread_b.start()
            thread_a.join()
            thread_b.join()
            ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
            Read_Energy_scale_list = Read_Energy_scale_new(Energy_20A_100mA_CT_list[i + 1][1],
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


def Sequence_Component_precision_measure():
    Sequence_Component_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Sequence_Component')
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
            Set_Service_Configuration(4)
            sequence_component_List = sequence_component_calculation(Sequence_Component_list[i + 1][1],
                                                                     Sequence_Component_list[i + 1][2],
                                                                     Sequence_Component_list[i + 1][3],
                                                                     Sequence_Component_list[i + 1][4],
                                                                     Sequence_Component_list[i + 1][5],
                                                                     Sequence_Component_list[i + 1][6])
            VUF = Read_Voltage_Unbalance_Factor_Magnitude(sequence_component_List[6], times=10)
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
            Set_Service_Configuration(4)
            sequence_component_List = sequence_component_calculation(Sequence_Component_list[i + 1][1],
                                                                     Sequence_Component_list[i + 1][2],
                                                                     Sequence_Component_list[i + 1][3],
                                                                     Sequence_Component_list[i + 1][4],
                                                                     Sequence_Component_list[i + 1][5],
                                                                     Sequence_Component_list[i + 1][6])
            CUF = Read_User_Channel_1_Current_Unbalance_Factor_Magnitude(sequence_component_List[6], times=10)
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
            Set_Service_Configuration(4)
            sequence_component_List = sequence_component_calculation(Sequence_Component_list[i + 1][1],
                                                                     Sequence_Component_list[i + 1][2],
                                                                     Sequence_Component_list[i + 1][3],
                                                                     Sequence_Component_list[i + 1][4],
                                                                     Sequence_Component_list[i + 1][5],
                                                                     Sequence_Component_list[i + 1][6])
            Acu4100_Voltage_Zero_Sequence = Read_Voltage_Zero_Sequence_Magnitude(sequence_component_List[0], times=40)
            Acu4100_Voltage_Zero_Sequence_Angle = Read_Voltage_Zero_Sequence_Angle(sequence_component_List[1], times=40)
            Acu4100_Voltage_Positive_Sequence = Read_Voltage_Positive_Sequence_Magnitude(sequence_component_List[2],
                                                                                         times=40)
            Acu4100_Voltage_Positive_Angle = Read_Voltage_Positive_Sequence_Angle(sequence_component_List[3], times=40)
            Acu4100_Voltage_Negative_Sequence = Read_Voltage_Negative_Sequence_Magnitude(sequence_component_List[4],
                                                                                         times=40)
            Acu4100_Voltage_Negative_Angle = Read_Voltage_Negative_Sequence_Angle(sequence_component_List[5], times=40)
            VUF = Read_Voltage_Unbalance_Factor_Magnitude(sequence_component_List[6], times=40)
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
            Set_Service_Configuration(4)
            sequence_component_List = sequence_component_calculation(Sequence_Component_list[i + 1][1],
                                                                     Sequence_Component_list[i + 1][2],
                                                                     Sequence_Component_list[i + 1][3],
                                                                     Sequence_Component_list[i + 1][4],
                                                                     Sequence_Component_list[i + 1][5],
                                                                     Sequence_Component_list[i + 1][6])
            Acu4100_Voltage_Zero_Sequence = Read_User_Channel_1_Current_Zero_Sequence_Magnitude(
                sequence_component_List[0], times=40)
            Acu4100_Voltage_Zero_Sequence_Angle = Read_User_Channel_1_Current_Zero_Sequence_Angle(
                sequence_component_List[1], times=40)
            Acu4100_Voltage_Positive_Sequence = Read_User_Channel_1_Current_Positive_Sequence_Magnitude(
                sequence_component_List[2], times=40)
            Acu4100_Voltage_Positive_Angle = Read_User_Channel_1_Current_Positive_Sequence_Angle(
                sequence_component_List[3], times=40)
            Acu4100_Voltage_Negative_Sequence = Read_User_Channel_1_Current_Negative_Sequence_Magnitude(
                sequence_component_List[4], times=40)
            Acu4100_Voltage_Negative_Angle = Read_User_Channel_1_Current_Negative_Sequence_Angle(
                sequence_component_List[5], times=40)
            CUF = Read_User_Channel_1_Current_Unbalance_Factor_Magnitude(sequence_component_List[6], times=40)
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
def line_to_neutral_voltage_precision_measure():
    voltage_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Line_to_Neutral_Voltage')
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
            Phase_A = Read_Phase_A_Voltage(voltage_list[i + 1][1], times=40)
            Phase_B = Read_Phase_B_Voltage(voltage_list[i + 1][2], times=40)
            Phase_C = Read_Phase_C_Voltage(voltage_list[i + 1][3], times=40)
            Read_Average_Vol = Read_Average_ln_Voltage(Average_Vol, times=40)
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
            Phase_A = Read_Phase_A_Voltage(voltage_list[i + 1][1], times=40)
            Phase_B = Read_Phase_B_Voltage(voltage_list[i + 1][2], times=40)
            Phase_C = Read_Phase_C_Voltage(voltage_list[i + 1][3], times=40)
            Read_Average_Vol = Read_Average_ln_Voltage(Average_Vol, times=40)
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
                else:
                    sheet.write(i + 1, 15, 'Failed')
            else:
                if Phase_A != 0 and Phase_B != 0 and Phase_C != 0 and Read_Average_Vol != 0:
                    sheet.write(i + 1, 15, 'Passed')
                else:
                    sheet.write(i + 1, 15, 'Failed')

        if voltage_list[i + 1][7] != 'null' and voltage_list[i + 1][1] != 'null' and voltage_list[i + 1][
            2] != 'null' and voltage_list[i + 1][3] != 'null' and voltage_list[i + 1][8] != 'null':
            if voltage_list[i + 1][8] == 'ABC':
                Set_Phase_Order(0)
                sheet.write(i + 1, 16, '相序设置：ABC')
            if voltage_list[i + 1][8] == 'ACB':
                Set_Phase_Order(1)
                sheet.write(i + 1, 16, '相序设置：ACB')
            ret = set_ac(voltage_list[i + 1][6], voltage_list[i + 1][5], voltage_list[i + 1][4], 120, 240, 0,
                         voltage_list[i + 1][3], voltage_list[i + 1][2], voltage_list[i + 1][1], 1, 1, 1, 50)
            Average_Vol = (voltage_list[i + 1][1] + voltage_list[i + 1][2] + voltage_list[i + 1][3]) / 3
            Phase_A = Read_Phase_A_Voltage(voltage_list[i + 1][1], times=40)
            Phase_B = Read_Phase_B_Voltage(voltage_list[i + 1][2], times=40)
            Phase_C = Read_Phase_C_Voltage(voltage_list[i + 1][3], times=40)
            Read_Average_Vol = Read_Average_ln_Voltage(Average_Vol, times=40)
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
            Set_Phase_Order(0)
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
    line_to_neutral_voltage_precision_measure()
    Energy_5A_333mV_CT_measure()
    Energy_5A_333mV_CT_measure1()
    ModbusClient.close()
    # my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))
    ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
    time.sleep(5)
    switch_device_screen_interface(0x00)
    print('====================测试总耗时:{}===================='.format(time.time() - start_time))
    print('====================={}====================='.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    print('====================Precision Measure End====================')