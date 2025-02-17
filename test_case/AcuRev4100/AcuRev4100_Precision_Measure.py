from AcuRev4100_modbus_get import *

# volt_cur_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'test_data')
Log(str(__file__).split("\\")[-1])


# ModbusClient = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])


def frequency_precision_measure():
    frequency_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'frequency')
    print(frequency_list)
    # my_workbook = xlwt.Workbook()
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
            # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 1, 1, 1, frequency_list[i + 1][1])
            frequency = read_frequency(frequency_list[i + 1][1], 10)
            scale = abs(frequency - frequency_list[i + 1][1]) / frequency_list[i + 1][1]
            sheet.write(i + 1, 2, frequency)
            sheet.write(i + 1, 3, f'{scale:.3%}')
            if scale <= frequency_list[i + 1][2] / 100:
                sheet.write(i + 1, 4, 'Passed')
            else:
                sheet.write(i + 1, 4, 'Failed')
        if frequency_list[i + 1][1] != 'null' and frequency_list[i + 1][2] == 'null':
            # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, 1, 1, 1, frequency_list[i + 1][1])
            frequency = read_frequency(frequency_list[i + 1][1], 10)
            sheet.write(i + 1, 2, frequency)
            if frequency != 0:
                sheet.write(i + 1, 4, 'Passed')
            else:
                sheet.write(i + 1, 4, 'Failed')
    # mes.close()
    # my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


def line_to_neutral_voltage_precision_measure():
    voltage_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Line_to_Neutral_Voltage')
    print(voltage_list)
    # my_workbook = xlwt.Workbook()
    sheet = my_workbook.add_sheet('line_to_neutral_voltage', cell_overwrite_ok=True)
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, 'Van输入值')
    sheet.write(0, 2, 'Vbn输入值')
    sheet.write(0, 3, 'Vcn输入值')
    sheet.write(0, 4, '电表实际Van值')
    sheet.write(0, 5, '电表实际Vbn值')
    sheet.write(0, 6, '电表实际Vcn值')
    sheet.write(0, 7, '电表实际Vlnavg值')
    sheet.write(0, 8, 'Van精度')
    sheet.write(0, 9, 'Vbn精度')
    sheet.write(0, 10, 'Vcn精度')
    sheet.write(0, 11, 'Vlnavg精度')
    sheet.write(0, 12, '测试结果')
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
        if voltage_list[i + 1][4] != 'null' and voltage_list[i + 1][1] != 'null' and voltage_list[i + 1][
            2] != 'null' and voltage_list[i + 1][3] != 'null':
            # ret = set_ac(120, 240, 0, 120, 240, 0, voltage_list[i + 1][3], voltage_list[i + 1][2],
            #              voltage_list[i + 1][1], 1, 1, 1, 50)
            Average_Vol = (voltage_list[i + 1][1] + voltage_list[i + 1][2] + voltage_list[i + 1][3]) / 3
            Phase_A = Read_Phase_A_Voltage(voltage_list[i + 1][1], 10)
            Phase_B = Read_Phase_B_Voltage(voltage_list[i + 1][2], 10)
            Phase_C = Read_Phase_C_Voltage(voltage_list[i + 1][3], 10)
            Read_Average_Vol = Read_Average_ln_Voltage(Average_Vol, 10)
            scale_A = abs(Phase_A - voltage_list[i + 1][1]) / voltage_list[i + 1][1]
            scale_B = abs(Phase_B - voltage_list[i + 1][2]) / voltage_list[i + 1][2]
            scale_C = abs(Phase_B - voltage_list[i + 1][3]) / voltage_list[i + 1][3]
            scale_Average_Vol = abs(Read_Average_Vol - Average_Vol) / Average_Vol
            sheet.write(i + 1, 4, Phase_A)
            sheet.write(i + 1, 5, Phase_B)
            sheet.write(i + 1, 6, Phase_C)
            sheet.write(i + 1, 7, Read_Average_Vol)
            sheet.write(i + 1, 8, f'{scale_A:.2%}')
            sheet.write(i + 1, 9, f'{scale_B:.2%}')
            sheet.write(i + 1, 10, f'{scale_C:.2%}')
            sheet.write(i + 1, 11, f'{scale_Average_Vol:.2%}')
            if scale_A <= voltage_list[i + 1][4] / 100 and scale_B <= voltage_list[i + 1][4] / 100 and scale_C <= \
                    voltage_list[i + 1][4] / 100:
                sheet.write(i + 1, 12, 'Passed')
            else:
                sheet.write(i + 1, 12, 'Failed')
        if voltage_list[i + 1][4] == 'null' and voltage_list[i + 1][1] != 'null' and voltage_list[i + 1][
            2] != 'null' and voltage_list[i + 1][3] != 'null':
            # ret = set_ac(120, 240, 0, 120, 240, 0, voltage_list[i + 1][3], voltage_list[i + 1][2],
            #              voltage_list[i + 1][1], 1, 1, 1, 50)
            Average_Vol = (voltage_list[i + 1][1] + voltage_list[i + 1][2] + voltage_list[i + 1][3]) / 3
            Phase_A = Read_Phase_A_Voltage(voltage_list[i + 1][1], 10)
            Phase_B = Read_Phase_B_Voltage(voltage_list[i + 1][2], 10)
            Phase_C = Read_Phase_C_Voltage(voltage_list[i + 1][3], 10)
            Read_Average_Vol = Read_Average_ln_Voltage(Average_Vol, 10)
            sheet.write(i + 1, 4, Phase_A)
            sheet.write(i + 1, 5, Phase_B)
            sheet.write(i + 1, 6, Phase_C)
            sheet.write(i + 1, 7, Read_Average_Vol)
            sheet.write(i + 1, 8, 'null')
            sheet.write(i + 1, 9, 'null')
            sheet.write(i + 1, 10, 'null')
            sheet.write(i + 1, 11, 'null')
            if voltage_list[i + 1][1] < 10 and voltage_list[i + 1][2] < 10 and voltage_list[i + 1][3] < 10:
                if Phase_A == Phase_B == Phase_C == Read_Average_Vol == 0:
                    sheet.write(i + 1, 12, 'Passed')
                else:
                    sheet.write(i + 1, 12, 'Failed')
            else:
                if Phase_A != 0 and Phase_B != 0 and Phase_C != 0 and Read_Average_Vol != 0:
                    sheet.write(i + 1, 12, 'Passed')
                else:
                    sheet.write(i + 1, 12, 'Failed')


def line_to_line_voltage_precision_measure():
    voltage_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Line_to_Line_Voltage')
    print(voltage_list)
    # my_workbook = xlwt.Workbook()
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
        if voltage_list[i + 1][7] != 'null':
            # ret = set_ac(voltage_list[i + 1][6], voltage_list[i + 1][5], voltage_list[i + 1][4], 120, 240, 0,
            #              voltage_list[i + 1][3], voltage_list[i + 1][2], voltage_list[i + 1][1], 1, 1, 1, 50)
            line_voltage_list = line_to_line_voltage_calculate(voltage_list[i + 1][1], voltage_list[i + 1][2],
                                                               voltage_list[i + 1][3],
                                                               voltage_list[i + 1][4], voltage_list[i + 1][5],
                                                               voltage_list[i + 1][6])
            Average_line_voltage = (line_voltage_list[0] + line_voltage_list[1] + line_voltage_list[2]) / 3
            Phase_AB_Voltage = Read_Phase_AB_Voltage(line_voltage_list[0], 10)
            Phase_BC_Voltage = Read_Phase_BC_Voltage(line_voltage_list[1], 10)
            Phase_CA_Voltage = Read_Phase_CA_Voltage(line_voltage_list[2], 10)
            Average_ll_Voltage = Read_Average_ll_Voltage(Average_line_voltage, 10)
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
            if scale_AB <= voltage_list[i + 1][7] / 100 and scale_BC <= voltage_list[i + 1][7] / 100 and scale_CA <= \
                    voltage_list[i + 1][7] / 100:
                sheet.write(i + 1, 15, 'Passed')
            else:
                sheet.write(i + 1, 15, 'Failed')
        if voltage_list[i + 1][7] == 'null':
            # ret = set_ac(voltage_list[i + 1][6], voltage_list[i + 1][5], voltage_list[i + 1][4], 120, 240, 0,
            #              voltage_list[i + 1][3], voltage_list[i + 1][2], voltage_list[i + 1][1], 1, 1, 1, 50)
            line_voltage_list = line_to_line_voltage_calculate(voltage_list[i + 1][1], voltage_list[i + 1][2],
                                                               voltage_list[i + 1][3],
                                                               voltage_list[i + 1][4], voltage_list[i + 1][5],
                                                               voltage_list[i + 1][6])
            Average_line_voltage = (line_voltage_list[0] + line_voltage_list[1] + line_voltage_list[2]) / 3
            Phase_AB_Voltage = Read_Phase_AB_Voltage(line_voltage_list[0], 10)
            Phase_BC_Voltage = Read_Phase_BC_Voltage(line_voltage_list[1], 10)
            Phase_CA_Voltage = Read_Phase_CA_Voltage(line_voltage_list[2], 10)
            Average_ll_Voltage = Read_Average_ll_Voltage(Average_line_voltage, 10)
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


def Current_5A_333mV_CT_precision_measure():
    Current_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Current_5A_333mV_CT')
    print(Current_list)
    j = 30
    for i, Current in enumerate(Current_list):
        if 'Input and user' in Current:
            j = i
            break
    # my_workbook = xlwt.Workbook()
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
                3] != 'null' and Current_list[i + 1][4] != 'null':
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = Read_Phase_A_Current(Current_list[i + 1][1], 10)
                Phase_B_Current = Read_Phase_B_Current(Current_list[i + 1][2], 10)
                Phase_C_Current = Read_Phase_C_Current(Current_list[i + 1][3], 10)
                Iavg = Read_System_Average_Current(Average_Current, 10)
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
                if scale_A_Current <= Current_list[i + 1][4] / 100 and scale_B_Current <= Current_list[i + 1][
                    4] / 100 and scale_C_Current <= Current_list[i + 1][4] / 100 and scale_Iavg <= Current_list[i + 1][
                    4] / 100:
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
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = Read_Phase_A_Current(Current_list[i + 1][1], 10)
                Phase_B_Current = Read_Phase_B_Current(Current_list[i + 1][2], 10)
                Phase_C_Current = Read_Phase_C_Current(Current_list[i + 1][3], 10)
                Iavg = Read_System_Average_Current(Average_Current, 10)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                scale_A_Current = 1
                scale_B_Current = 1
                scale_C_Current = 1
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
                if scale_A_Current <= Current_list[i + 1][4] / 100 and scale_B_Current <= Current_list[i + 1][
                    4] / 100 and scale_C_Current <= Current_list[i + 1][4] / 100 and scale_Iavg <= Current_list[i + 1][
                    4] / 100:
                    sheet.write(i + 1, 12, 'Passed')
                else:
                    sheet.write(i + 1, 12, 'Failed')
            if Current_list[i + 1][4] == 'null':
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = Read_Phase_A_Current(Current_list[i + 1][1], 10)
                Phase_B_Current = Read_Phase_B_Current(Current_list[i + 1][2], 10)
                Phase_C_Current = Read_Phase_C_Current(Current_list[i + 1][3], 10)
                Iavg = Read_System_Average_Current(Average_Current, 10)
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
            sheet.write(j, 10, 'Input1 精度')
            sheet.write(j, 11, 'Input2 精度')
            sheet.write(j, 12, 'Input3 精度')
            sheet.write(j, 13, 'User1 精度')
            sheet.write(j, 14, 'User1 精度')
            sheet.write(j, 15, 'User1 精度')
            sheet.write(j, 16, '测试结果')
            sheet.write(i + 2, 0, Current_list[i + 1][0])
            sheet.write(i + 2, 1, Current_list[i + 1][1])
            sheet.write(i + 2, 2, Current_list[i + 1][2])
            sheet.write(i + 2, 3, Current_list[i + 1][3])
            if Current_list[i + 1][4] != 'null':
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Input_Channel_1_Current = Read_Input_Channel_1_Current(Current_list[i + 1][1], 10)
                Input_Channel_2_Current = Read_Input_Channel_2_Current(Current_list[i + 1][2], 10)
                Input_Channel_3_Current = Read_Input_Channel_3_Current(Current_list[i + 1][3], 10)
                User_Channel_1_Current = Read_User_Channel_1_Current(Current_list[i + 1][1], 10)
                User_Channel_2_Current = Read_User_Channel_2_Current(Current_list[i + 1][2], 10)
                User_Channel_3_Current = Read_User_Channel_3_Current(Current_list[i + 1][3], 10)
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
                scale_User_Channel_1_Current = abs(User_Channel_1_Current - Current_list[i + 1][1]) / \
                                               Current_list[i + 1][1]
                scale_User_Channel_2_Current = abs(User_Channel_2_Current - Current_list[i + 1][2]) / \
                                               Current_list[i + 1][2]
                scale_User_Channel_3_Current = abs(User_Channel_3_Current - Current_list[i + 1][3]) / \
                                               Current_list[i + 1][3]
                sheet.write(i + 2, 10, f'{scale_Input_Channel_1_Current:.2%}')
                sheet.write(i + 2, 11, f'{scale_Input_Channel_2_Current:.2%}')
                sheet.write(i + 2, 12, f'{scale_Input_Channel_3_Current:.2%}')
                sheet.write(i + 2, 13, f'{scale_User_Channel_1_Current:.2%}')
                sheet.write(i + 2, 14, f'{scale_User_Channel_2_Current:.2%}')
                sheet.write(i + 2, 15, f'{scale_User_Channel_3_Current:.2%}')
                if scale_Input_Channel_1_Current <= Current_list[i + 1][4] / 100 and scale_Input_Channel_2_Current <= \
                        Current_list[i + 1][4] / 100 and scale_Input_Channel_3_Current <= Current_list[i + 1][
                    4] / 100 and scale_User_Channel_1_Current <= Current_list[i + 1][
                    4] / 100 and scale_User_Channel_2_Current <= Current_list[i + 1][
                    4] / 100 and scale_User_Channel_3_Current <= Current_list[i + 1][4] / 100:
                    sheet.write(i + 2, 16, 'Passed')
                else:
                    sheet.write(i + 2, 16, 'Failed')
            if Current_list[i + 1][4] == 'null':
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Input_Channel_1_Current = Read_Input_Channel_1_Current(Current_list[i + 1][1], 10)
                Input_Channel_2_Current = Read_Input_Channel_2_Current(Current_list[i + 1][2], 10)
                Input_Channel_3_Current = Read_Input_Channel_3_Current(Current_list[i + 1][3], 10)
                User_Channel_1_Current = Read_User_Channel_1_Current(Current_list[i + 1][1], 10)
                User_Channel_2_Current = Read_User_Channel_2_Current(Current_list[i + 1][2], 10)
                User_Channel_3_Current = Read_User_Channel_3_Current(Current_list[i + 1][3], 10)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                sheet.write(i + 2, 10, f'null')
                sheet.write(i + 2, 11, f'null')
                sheet.write(i + 2, 12, f'null')
                sheet.write(i + 2, 13, f'null')
                sheet.write(i + 2, 14, f'null')
                sheet.write(i + 2, 15, f'null')
                if Current_list[i + 1][1] >= 0.005 and Current_list[i + 1][2] >= 0.005 and Current_list[i + 1][
                    3] >= 0.005:
                    if Input_Channel_1_Current != 0 and Input_Channel_2_Current != 0 and Input_Channel_3_Current != 0 and User_Channel_1_Current != 0 and User_Channel_2_Current != 0 and User_Channel_3_Current != 0:
                        sheet.write(i + 2, 16, 'Passed')
                    else:
                        sheet.write(i + 2, 16, 'Failed')
                if Current_list[i + 1][1] == 0 and Current_list[i + 1][2] == 0 and Current_list[i + 1][
                    3] == 0:
                    if Input_Channel_1_Current == Input_Channel_2_Current == Input_Channel_3_Current == User_Channel_1_Current == User_Channel_2_Current == User_Channel_3_Current == 0:
                        sheet.write(i + 2, 16, 'Passed')
                    else:
                        sheet.write(i + 2, 16, 'Failed')


def Current_20A_100mA_CT_precision_measure():
    Current_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Current_20A_100mA_CT')
    print(Current_list)
    j = 30
    for i, Current in enumerate(Current_list):
        if 'Input and user' in Current:
            j = i
            break
    # my_workbook = xlwt.Workbook()
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
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = Read_Phase_A_Current(Current_list[i + 1][1], 10)
                Phase_B_Current = Read_Phase_B_Current(Current_list[i + 1][2], 10)
                Phase_C_Current = Read_Phase_C_Current(Current_list[i + 1][3], 10)
                Iavg = Read_System_Average_Current(Average_Current, 10)
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
                if scale_A_Current <= Current_list[i + 1][4] / 100 and scale_B_Current <= Current_list[i + 1][
                    4] / 100 and scale_C_Current <= Current_list[i + 1][4] / 100 and scale_Iavg <= Current_list[i + 1][
                    4] / 100:
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
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = Read_Phase_A_Current(Current_list[i + 1][1], 10)
                Phase_B_Current = Read_Phase_B_Current(Current_list[i + 1][2], 10)
                Phase_C_Current = Read_Phase_C_Current(Current_list[i + 1][3], 10)
                Iavg = Read_System_Average_Current(Average_Current, 10)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                scale_A_Current = 1
                scale_B_Current = 1
                scale_C_Current = 1
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
                if scale_A_Current <= Current_list[i + 1][4] / 100 and scale_B_Current <= Current_list[i + 1][
                    4] / 100 and scale_C_Current <= Current_list[i + 1][4] / 100 and scale_Iavg <= Current_list[i + 1][
                    4] / 100:
                    sheet.write(i + 1, 12, 'Passed')
                else:
                    sheet.write(i + 1, 12, 'Failed')
            if Current_list[i + 1][4] == 'null':
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Average_Current = (Current_list[i + 1][1] + Current_list[i + 1][2] + Current_list[i + 1][3]) / 3
                Phase_A_Current = Read_Phase_A_Current(Current_list[i + 1][1], 10)
                Phase_B_Current = Read_Phase_B_Current(Current_list[i + 1][2], 10)
                Phase_C_Current = Read_Phase_C_Current(Current_list[i + 1][3], 10)
                Iavg = Read_System_Average_Current(Average_Current, 10)
                sheet.write(i + 1, 4, Phase_A_Current)
                sheet.write(i + 1, 5, Phase_B_Current)
                sheet.write(i + 1, 6, Phase_C_Current)
                sheet.write(i + 1, 7, Iavg)
                sheet.write(i + 1, 8, 'null')
                sheet.write(i + 1, 9, 'null')
                sheet.write(i + 1, 10, 'null')
                sheet.write(i + 1, 11, 'null')
                if Current_list[i + 1][1] < 0.02 and Current_list[i + 1][2] < 0.02 and Current_list[i + 1][3] < 0.02:
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
            sheet.write(j, 10, 'Input1 精度')
            sheet.write(j, 11, 'Input2 精度')
            sheet.write(j, 12, 'Input3 精度')
            sheet.write(j, 13, 'User1 精度')
            sheet.write(j, 14, 'User1 精度')
            sheet.write(j, 15, 'User1 精度')
            sheet.write(j, 16, '测试结果')
            sheet.write(i + 2, 0, Current_list[i + 1][0])
            sheet.write(i + 2, 1, Current_list[i + 1][1])
            sheet.write(i + 2, 2, Current_list[i + 1][2])
            sheet.write(i + 2, 3, Current_list[i + 1][3])
            if Current_list[i + 1][4] != 'null':
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Input_Channel_1_Current = Read_Input_Channel_1_Current(Current_list[i + 1][1], 10)
                Input_Channel_2_Current = Read_Input_Channel_2_Current(Current_list[i + 1][2], 10)
                Input_Channel_3_Current = Read_Input_Channel_3_Current(Current_list[i + 1][3], 10)
                User_Channel_1_Current = Read_User_Channel_1_Current(Current_list[i + 1][1], 10)
                User_Channel_2_Current = Read_User_Channel_2_Current(Current_list[i + 1][2], 10)
                User_Channel_3_Current = Read_User_Channel_3_Current(Current_list[i + 1][3], 10)
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
                scale_User_Channel_1_Current = abs(User_Channel_1_Current - Current_list[i + 1][1]) / \
                                               Current_list[i + 1][1]
                scale_User_Channel_2_Current = abs(User_Channel_2_Current - Current_list[i + 1][2]) / \
                                               Current_list[i + 1][2]
                scale_User_Channel_3_Current = abs(User_Channel_3_Current - Current_list[i + 1][3]) / \
                                               Current_list[i + 1][3]
                sheet.write(i + 2, 10, f'{scale_Input_Channel_1_Current:.2%}')
                sheet.write(i + 2, 11, f'{scale_Input_Channel_2_Current:.2%}')
                sheet.write(i + 2, 12, f'{scale_Input_Channel_3_Current:.2%}')
                sheet.write(i + 2, 13, f'{scale_User_Channel_1_Current:.2%}')
                sheet.write(i + 2, 14, f'{scale_User_Channel_2_Current:.2%}')
                sheet.write(i + 2, 15, f'{scale_User_Channel_3_Current:.2%}')
                if scale_Input_Channel_1_Current <= Current_list[i + 1][4] / 100 and scale_Input_Channel_2_Current <= \
                        Current_list[i + 1][4] / 100 and scale_Input_Channel_3_Current <= Current_list[i + 1][
                    4] / 100 and scale_User_Channel_1_Current <= Current_list[i + 1][
                    4] / 100 and scale_User_Channel_2_Current <= Current_list[i + 1][
                    4] / 100 and scale_User_Channel_3_Current <= Current_list[i + 1][4] / 100:
                    sheet.write(i + 2, 16, 'Passed')
                else:
                    sheet.write(i + 2, 16, 'Failed')
            if Current_list[i + 1][4] == 'null':
                # ret = set_ac(120, 240, 0, 120, 240, 0, 50, 50, 50, Current_list[i + 1][3], Current_list[i + 1][2],
                #              Current_list[i + 1][1], 50)
                Input_Channel_1_Current = Read_Input_Channel_1_Current(Current_list[i + 1][1], 10)
                Input_Channel_2_Current = Read_Input_Channel_2_Current(Current_list[i + 1][2], 10)
                Input_Channel_3_Current = Read_Input_Channel_3_Current(Current_list[i + 1][3], 10)
                User_Channel_1_Current = Read_User_Channel_1_Current(Current_list[i + 1][1], 10)
                User_Channel_2_Current = Read_User_Channel_2_Current(Current_list[i + 1][2], 10)
                User_Channel_3_Current = Read_User_Channel_3_Current(Current_list[i + 1][3], 10)
                sheet.write(i + 2, 4, Input_Channel_1_Current)
                sheet.write(i + 2, 5, Input_Channel_2_Current)
                sheet.write(i + 2, 6, Input_Channel_3_Current)
                sheet.write(i + 2, 7, User_Channel_1_Current)
                sheet.write(i + 2, 8, User_Channel_2_Current)
                sheet.write(i + 2, 9, User_Channel_3_Current)
                sheet.write(i + 2, 10, f'null')
                sheet.write(i + 2, 11, f'null')
                sheet.write(i + 2, 12, f'null')
                sheet.write(i + 2, 13, f'null')
                sheet.write(i + 2, 14, f'null')
                sheet.write(i + 2, 15, f'null')
                if Current_list[i + 1][1] >= 0.02 and Current_list[i + 1][2] >= 0.02 and Current_list[i + 1][
                    3] >= 0.02:
                    if Input_Channel_1_Current != 0 and Input_Channel_2_Current != 0 and Input_Channel_3_Current != 0 and User_Channel_1_Current != 0 and User_Channel_2_Current != 0 and User_Channel_3_Current != 0:
                        sheet.write(i + 2, 16, 'Passed')
                    else:
                        sheet.write(i + 2, 16, 'Failed')
                if Current_list[i + 1][1] == 0 and Current_list[i + 1][2] == 0 and Current_list[i + 1][
                    3] == 0:
                    if Input_Channel_1_Current == Input_Channel_2_Current == Input_Channel_3_Current == User_Channel_1_Current == User_Channel_2_Current == User_Channel_3_Current == 0:
                        sheet.write(i + 2, 16, 'Passed')
                    else:
                        sheet.write(i + 2, 16, 'Failed')


def Phase_Voltage_Angle_precision_measure():
    Voltage_Angle_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Phase_Voltage_Angle')
    print(Voltage_Angle_list)
    # my_workbook = xlwt.Workbook()
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
            # ret = set_ac(Voltage_Angle_list[i + 1][3], Voltage_Angle_list[i + 1][2], Voltage_Angle_list[i + 1][1], 120,
            #              240, 0, 100, 100, 100, 1, 1, 1, 50)
            Phase_A_Voltage_Angle = Read_Phase_A_Voltage_Angle(Voltage_Angle_list[i + 1][1], 10)
            Phase_B_Voltage_Angle = Read_Phase_B_Voltage_Angle(Voltage_Angle_list[i + 1][2], 10)
            Phase_C_Voltage_Angle = Read_Phase_C_Voltage_Angle(Voltage_Angle_list[i + 1][3], 10)
            sheet.write(i + 1, 4, Phase_A_Voltage_Angle)
            sheet.write(i + 1, 5, Phase_B_Voltage_Angle)
            sheet.write(i + 1, 6, Phase_C_Voltage_Angle)
            scale_B_Voltage_Angle = abs(Phase_B_Voltage_Angle - Voltage_Angle_list[i + 1][2]) / \
                                    Voltage_Angle_list[i + 1][2]
            scale_C_Voltage_Angle = abs(Phase_C_Voltage_Angle - Voltage_Angle_list[i + 1][3]) / \
                                    Voltage_Angle_list[i + 1][3]
            sheet.write(i + 1, 7, 'null')
            sheet.write(i + 1, 8, f'{scale_B_Voltage_Angle:.2%}')
            sheet.write(i + 1, 9, f'{scale_C_Voltage_Angle:.2%}')
            if Phase_A_Voltage_Angle == 0 and scale_B_Voltage_Angle <= Voltage_Angle_list[i + 1][
                4] / 100 and scale_C_Voltage_Angle <= Voltage_Angle_list[i + 1][4] / 100:
                sheet.write(i + 1, 10, 'Passed')
            else:
                sheet.write(i + 1, 10, 'Failed')


def Input1_Current_Angle_precision_measure():
    Input_Current_Angle = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Input_Current_Angle')
    print(Input_Current_Angle)
    # my_workbook = xlwt.Workbook()
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
            # ret = set_ac(120, 240, 0, 120, 240, Input_Current_Angle[i + 1][1], 100, 100, 100, 1, 1, 1, 50)
            Input_Channel_1_Current_Phase_Angle = Read_Input_Channel_1_Current_Phase_Angle(
                Input_Current_Angle[i + 1][1], 10)
            sheet.write(i + 1, 2, Input_Channel_1_Current_Phase_Angle)
            if Input_Current_Angle[i + 1][1] == 0:
                sheet.write(i + 1, 3, 'null')
                if Input_Channel_1_Current_Phase_Angle == 0:
                    sheet.write(i + 1, 4, 'Passed')
                else:
                    sheet.write(i + 1, 4, 'Failed')
            else:
                scale_Input1_Current_Angle = abs(
                    Input_Channel_1_Current_Phase_Angle - Input_Current_Angle[i + 1][1] / Input_Current_Angle[i + 1][1])
                sheet.write(i + 1, 3, f'{scale_Input1_Current_Angle:.2%}')
                if scale_Input1_Current_Angle <= Input_Current_Angle[i + 1][4] / 100:
                    sheet.write(i + 1, 4, 'Passed')
                else:
                    sheet.write(i + 1, 4, 'Failed')


def Power_5A_333mV_CT_precision_measure():
    Power_5A_333mV_CT_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Power_5A_333mV_CT')
    print(Power_5A_333mV_CT_list)
    # my_workbook = xlwt.Workbook()
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
        if Power_5A_333mV_CT_list[i + 1][14] == 'user1' and Power_5A_333mV_CT_list[i + 1][13] != 'null' or \
                Power_5A_333mV_CT_list[i + 1][15] == 'True Reactive Power':
            # set_ac(Power_5A_333mV_CT_list[i + 1][9], Power_5A_333mV_CT_list[i + 1][8], Power_5A_333mV_CT_list[i + 1][7],
            #        Power_5A_333mV_CT_list[i + 1][12], Power_5A_333mV_CT_list[i + 1][11],
            #        Power_5A_333mV_CT_list[i + 1][10], Power_5A_333mV_CT_list[i + 1][3],
            #        Power_5A_333mV_CT_list[i + 1][2], Power_5A_333mV_CT_list[i + 1][1], Power_5A_333mV_CT_list[i + 1][6],
            #        Power_5A_333mV_CT_list[i + 1][5], Power_5A_333mV_CT_list[i + 1][4], 50)
            AcuRev4100_Power = Read_AcuRev4100_Power(Power_5A_333mV_CT_list[i + 1][1], Power_5A_333mV_CT_list[i + 1][2],
                                                     Power_5A_333mV_CT_list[i + 1][3], Power_5A_333mV_CT_list[i + 1][4],
                                                     Power_5A_333mV_CT_list[i + 1][5], Power_5A_333mV_CT_list[i + 1][6],
                                                     Power_5A_333mV_CT_list[i + 1][7], Power_5A_333mV_CT_list[i + 1][8],
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
                    if k != 'null' and scale_list[k] <= Power_5A_333mV_CT_list[i + 1][13] / 100:
                        sheet.write(i + 1, 52, 'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 52, 'Failed')
                        break
            else:
                if Power_5A_333mV_CT_list[i+1][7] - Power_5A_333mV_CT_list[i+1][10] == 0:
                    if AcuRev4100_Power[1][0] == AcuRev4100_Power[5][0] == AcuRev4100_Power[9][0] == \
                            AcuRev4100_Power[13][0] == AcuRev4100_Power[17][0] == AcuRev4100_Power[21][0] == \
                            AcuRev4100_Power[25][0] == 0:
                        sheet.write(i + 1, 52, 'Passed')
                else:
                    for k in range(len(scale_list)):
                        if k != 'null' and scale_list[k] <= Power_5A_333mV_CT_list[i + 1][13] / 100:
                            sheet.write(i + 1, 52, 'Passed')
                            continue
                        else:
                            sheet.write(i + 1, 52, 'Failed')
                            break

        if Power_5A_333mV_CT_list[i + 1][14] == 'user1' and Power_5A_333mV_CT_list[i + 1][13] == 'null':
            # set_ac(Power_5A_333mV_CT_list[i + 1][9], Power_5A_333mV_CT_list[i + 1][8], Power_5A_333mV_CT_list[i + 1][7],
            #        Power_5A_333mV_CT_list[i + 1][12], Power_5A_333mV_CT_list[i + 1][11],
            #        Power_5A_333mV_CT_list[i + 1][10], Power_5A_333mV_CT_list[i + 1][3],
            #        Power_5A_333mV_CT_list[i + 1][2], Power_5A_333mV_CT_list[i + 1][1], Power_5A_333mV_CT_list[i + 1][6],
            #        Power_5A_333mV_CT_list[i + 1][5], Power_5A_333mV_CT_list[i + 1][4], 50)
            AcuRev4100_Power = Read_AcuRev4100_Power(Power_5A_333mV_CT_list[i + 1][1], Power_5A_333mV_CT_list[i + 1][2],
                                                     Power_5A_333mV_CT_list[i + 1][3], Power_5A_333mV_CT_list[i + 1][4],
                                                     Power_5A_333mV_CT_list[i + 1][5], Power_5A_333mV_CT_list[i + 1][6],
                                                     Power_5A_333mV_CT_list[i + 1][7], Power_5A_333mV_CT_list[i + 1][8],
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
                        sheet.write(i + 1, 52, 'Passed')
                        break
                    else:
                        sheet.write(i + 1, 52, 'Failed')
                        break
            else:
                for k in range(len(Power_list)):
                    if Power_list[k] == 0:
                        sheet.write(i + 1, 52, 'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 52, 'Failed')
                        break
        if Power_5A_333mV_CT_list[i + 1][14] == 'user1,user2,user3' and Power_5A_333mV_CT_list[i + 1][13] != 'null':
            # set_ac(Power_5A_333mV_CT_list[i + 1][9], Power_5A_333mV_CT_list[i + 1][8], Power_5A_333mV_CT_list[i + 1][7],
            #        Power_5A_333mV_CT_list[i + 1][12], Power_5A_333mV_CT_list[i + 1][11],
            #        Power_5A_333mV_CT_list[i + 1][10], Power_5A_333mV_CT_list[i + 1][3],
            #        Power_5A_333mV_CT_list[i + 1][2], Power_5A_333mV_CT_list[i + 1][1], Power_5A_333mV_CT_list[i + 1][6],
            #        Power_5A_333mV_CT_list[i + 1][5], Power_5A_333mV_CT_list[i + 1][4], 50)
            AcuRev4100_Power = Read_AcuRev4100_Power(Power_5A_333mV_CT_list[i + 1][1], Power_5A_333mV_CT_list[i + 1][2],
                                                     Power_5A_333mV_CT_list[i + 1][3], Power_5A_333mV_CT_list[i + 1][4],
                                                     Power_5A_333mV_CT_list[i + 1][5], Power_5A_333mV_CT_list[i + 1][6],
                                                     Power_5A_333mV_CT_list[i + 1][7], Power_5A_333mV_CT_list[i + 1][8],
                                                     Power_5A_333mV_CT_list[i + 1][9],
                                                     Power_5A_333mV_CT_list[i + 1][10],
                                                     Power_5A_333mV_CT_list[i + 1][11],
                                                     Power_5A_333mV_CT_list[i + 1][12],
                                                     Power_5A_333mV_CT_list[i + 1][14])
            scale_list = []
            for j in range(len(AcuRev4100_Power)):
                sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]:.2%}')
                scale_list.append(AcuRev4100_Power[j][1])
            for k in range(len(scale_list)):
                if scale_list[k] <= Power_5A_333mV_CT_list[i + 1][13] / 100:
                    sheet.write(i + 1, 52, 'Passed')
                    continue
                else:
                    sheet.write(i + 1, 52, 'Failed')
                    break


def Power_20A_100mA_CT_precision_measure():
    Power_20A_100mA_CT_list = data_read(r'./test_case/AcuRev4100/4100_test_case.xlsx', 'Power_20A_100mA_CT')
    print(Power_20A_100mA_CT_list)
    # my_workbook = xlwt.Workbook()
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
        if Power_20A_100mA_CT_list[i + 1][14] == 'user1' and Power_20A_100mA_CT_list[i + 1][13] != 'null' or \
                Power_20A_100mA_CT_list[i + 1][15] == 'True Reactive Power':
            # set_ac(Power_20A_100mA_CT_list[i + 1][9], Power_20A_100mA_CT_list[i + 1][8], Power_20A_100mA_CT_list[i + 1][7],
            #        Power_20A_100mA_CT_list[i + 1][12], Power_20A_100mA_CT_list[i + 1][11],
            #        Power_20A_100mA_CT_list[i + 1][10], Power_20A_100mA_CT_list[i + 1][3],
            #        Power_20A_100mA_CT_list[i + 1][2], Power_20A_100mA_CT_list[i + 1][1], Power_20A_100mA_CT_list[i + 1][6],
            #        Power_20A_100mA_CT_list[i + 1][5], Power_20A_100mA_CT_list[i + 1][4], 50)
            AcuRev4100_Power = Read_AcuRev4100_Power(Power_20A_100mA_CT_list[i + 1][1], Power_20A_100mA_CT_list[i + 1][2],
                                                     Power_20A_100mA_CT_list[i + 1][3], Power_20A_100mA_CT_list[i + 1][4],
                                                     Power_20A_100mA_CT_list[i + 1][5], Power_20A_100mA_CT_list[i + 1][6],
                                                     Power_20A_100mA_CT_list[i + 1][7], Power_20A_100mA_CT_list[i + 1][8],
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
                    if k != 'null' and scale_list[k] <= Power_20A_100mA_CT_list[i + 1][13] / 100:
                        sheet.write(i + 1, 52, 'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 52, 'Failed')
                        break
            else:
                if Power_20A_100mA_CT_list[i + 1][7] - Power_20A_100mA_CT_list[i + 1][10] == 0:
                    if AcuRev4100_Power[1][0] == AcuRev4100_Power[5][0] == AcuRev4100_Power[9][0] == \
                            AcuRev4100_Power[13][0] == AcuRev4100_Power[17][0] == AcuRev4100_Power[21][0] == \
                            AcuRev4100_Power[25][0] == 0:
                        sheet.write(i + 1, 52, 'Passed')
                else:
                    for k in range(len(scale_list)):
                        if k != 'null' and scale_list[k] <= Power_20A_100mA_CT_list[i + 1][13] / 100:
                            sheet.write(i + 1, 52, 'Passed')
                            continue
                        else:
                            sheet.write(i + 1, 52, 'Failed')
                            break

        if Power_20A_100mA_CT_list[i + 1][14] == 'user1' and Power_20A_100mA_CT_list[i + 1][13] == 'null':
            # set_ac(Power_20A_100mA_CT_list[i + 1][9], Power_20A_100mA_CT_list[i + 1][8], Power_20A_100mA_CT_list[i + 1][7],
            #        Power_20A_100mA_CT_list[i + 1][12], Power_20A_100mA_CT_list[i + 1][11],
            #        Power_20A_100mA_CT_list[i + 1][10], Power_20A_100mA_CT_list[i + 1][3],
            #        Power_20A_100mA_CT_list[i + 1][2], Power_20A_100mA_CT_list[i + 1][1], Power_20A_100mA_CT_list[i + 1][6],
            #        Power_20A_100mA_CT_list[i + 1][5], Power_20A_100mA_CT_list[i + 1][4], 50)
            AcuRev4100_Power = Read_AcuRev4100_Power(Power_20A_100mA_CT_list[i + 1][1], Power_20A_100mA_CT_list[i + 1][2],
                                                     Power_20A_100mA_CT_list[i + 1][3], Power_20A_100mA_CT_list[i + 1][4],
                                                     Power_20A_100mA_CT_list[i + 1][5], Power_20A_100mA_CT_list[i + 1][6],
                                                     Power_20A_100mA_CT_list[i + 1][7], Power_20A_100mA_CT_list[i + 1][8],
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
                        sheet.write(i + 1, 52, 'Passed')
                        break
                    else:
                        sheet.write(i + 1, 52, 'Failed')
                        break
            else:
                for k in range(len(Power_list)):
                    if Power_list[k] == 0:
                        sheet.write(i + 1, 52, 'Passed')
                        continue
                    else:
                        sheet.write(i + 1, 52, 'Failed')
                        break
        if Power_20A_100mA_CT_list[i + 1][14] == 'user1,user2,user3' and Power_20A_100mA_CT_list[i + 1][13] != 'null':
            # set_ac(Power_20A_100mA_CT_list[i + 1][9], Power_20A_100mA_CT_list[i + 1][8], Power_20A_100mA_CT_list[i + 1][7],
            #        Power_20A_100mA_CT_list[i + 1][12], Power_20A_100mA_CT_list[i + 1][11],
            #        Power_20A_100mA_CT_list[i + 1][10], Power_20A_100mA_CT_list[i + 1][3],
            #        Power_20A_100mA_CT_list[i + 1][2], Power_20A_100mA_CT_list[i + 1][1], Power_20A_100mA_CT_list[i + 1][6],
            #        Power_20A_100mA_CT_list[i + 1][5], Power_20A_100mA_CT_list[i + 1][4], 50)
            AcuRev4100_Power = Read_AcuRev4100_Power(Power_20A_100mA_CT_list[i + 1][1], Power_20A_100mA_CT_list[i + 1][2],
                                                     Power_20A_100mA_CT_list[i + 1][3], Power_20A_100mA_CT_list[i + 1][4],
                                                     Power_20A_100mA_CT_list[i + 1][5], Power_20A_100mA_CT_list[i + 1][6],
                                                     Power_20A_100mA_CT_list[i + 1][7], Power_20A_100mA_CT_list[i + 1][8],
                                                     Power_20A_100mA_CT_list[i + 1][9],
                                                     Power_20A_100mA_CT_list[i + 1][10],
                                                     Power_20A_100mA_CT_list[i + 1][11],
                                                     Power_20A_100mA_CT_list[i + 1][12],
                                                     Power_20A_100mA_CT_list[i + 1][14])
            scale_list = []
            for j in range(len(AcuRev4100_Power)):
                sheet.write(i + 1, j + 13, f'{AcuRev4100_Power[j][0]},{AcuRev4100_Power[j][1]:.2%}')
                scale_list.append(AcuRev4100_Power[j][1])
            for k in range(len(scale_list)):
                if scale_list[k] <= Power_20A_100mA_CT_list[i + 1][13] / 100:
                    sheet.write(i + 1, 52, 'Passed')
                    continue
                else:
                    sheet.write(i + 1, 52, 'Failed')
                    break


if __name__ == '__main__':
    print('====================Precision Measure Start====================')
    print('======================{}======================'.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    start_time = time.time()
    my_workbook = xlwt.Workbook()
    frequency_precision_measure()
    line_to_neutral_voltage_precision_measure()
    line_to_line_voltage_precision_measure()
    Current_5A_333mV_CT_precision_measure()
    Current_20A_100mA_CT_precision_measure()
    Power_5A_333mV_CT_precision_measure()
    Power_20A_100mA_CT_precision_measure()
    Phase_Voltage_Angle_precision_measure()
    Input1_Current_Angle_precision_measure()
    ModbusClient.close()
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))
    print('====================测试总耗时:{}===================='.format(time.time() - start_time))
    print('====================={}====================='.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    print('====================Precision Measure End====================')
