"""
文件名称: acuvim_precision_measure.py
功能描述: AcuvimⅡ测量数据精度测试
创建日期: 2025-05-15
作者: 苏博
版本: v1.0
修改记录: 2025-05-15 苏博 创建脚本
"""
import csv
import time
from comm.source_control import *
from tools.excel_operate import data_read
from comm.acuvim_modbus_get_attr import *
import xlwt
from math import cos, sin
from tools.log import Log

Log(str(__file__).split("\\")[-1])

my_workbook = xlwt.Workbook()
sheet = my_workbook.add_sheet('precision_measure', cell_overwrite_ok=True)
energy_sheet = my_workbook.add_sheet('energy_precision_measure', cell_overwrite_ok=True)
test_data = data_read('/comm/test_data/acuvim_test_case.xlsx', 'test_case')
precision_requirement = data_read('/comm/test_data/acuvim_test_case.xlsx', 'precision_requirement')


def write_data_csv(row, data: list):
    if row == 0:
        with open("../../comm/test_data/AcuvimⅡ_Precision_Measure_realtime_basicpara.csv", "a", newline="", encoding="gbk") as rb:
            writer = csv.writer(rb)
            writer.writerow(data)
        return
    with open("../../comm/test_data/AcuvimⅡ_Precision_Measure_energy.csv", "a", newline="", encoding="gbk") as rb:
        writer = csv.writer(rb)
        writer.writerow(data)
    return


def write_excel_topic():
    topic_list = ['测试用例', 'A相电压输入值', 'B相电压输入值', 'C相电压输入值', 'A相电流输入值',
                  'B相电流输入值',
                  'C相电流输入值', 'A相电压相位输入值', 'B相电压相位输入值', 'C相电压相位输入值',
                  'A相电流相位输入值',
                  'B相电流相位输入值', 'C相电流相位输入值', 'A相有功功率输入值', 'B相有功功率输入值',
                  'C相有功功率输入值',
                  'A相无功功率输入值', 'B相无功功率输入值', 'C相无功功率输入值', 'A相视在功率输入值',
                  'B相视在功率输入值',
                  'C相视在功率输入值',
                  'A相电压测量值', 'B相电压测量值', 'C相电压测量值', 'A相电流测量值', 'B相电流测量值',
                  'C相电流测量值',
                  'A相电压相位测量值', 'B相电压相位测量值', 'C相电压相位测量值', 'A相电流相位测量值',
                  'B相电流相位测量值', 'C相电流相位测量值', 'A相有功功率测量值', 'B相有功功率测量值',
                  'C相有功功率测量值', 'A相无功功率测量值', 'B相无功功率测量值', 'C相无功功率测量值',
                  'A相视在功率测量值', 'B相视在功率测量值', 'C相视在功率测量值',
                  'A相电压精度', 'B相电压精度', 'C相电压精度', 'A相电流精度', 'B相电流精度', 'C相电流精度',
                  'A相电压相位精度', 'B相电压相位精度', 'C相电压相位精度', 'A相电流相位精度', 'B相电流相位精度',
                  'C相电流相位精度', 'A相有功功率精度', 'B相有功功率精度', 'C相有功功率精度', 'A相无功功率精度',
                  'B相无功功率精度', 'C相无功功率精度', 'A相视在功率精度', 'B相视在功率精度', 'C相视在功率精度',
                  '测试结果']
    write_data_csv(row=0, data=topic_list)
    for i in range(0, len(topic_list)):
        sheet.write(0, i, topic_list[i])
    topic_list = ['测试用例', 'A相电压输入值', 'B相电压输入值', 'C相电压输入值', 'A相电流输入值', 'B相电流输入值',
                  'C相电流输入值', 'A相电压相位输入值', 'B相电压相位输入值', 'C相电压相位输入值',
                  'A相电流相位输入值', 'B相电流相位输入值', 'C相电流相位输入值',
                  'A相有功功率输入值',
                  'B相有功功率输入值',
                  'C相有功功率输入值',
                  'A相无功功率输入值',
                  'B相无功功率输入值',
                  'C相无功功率输入值',
                  'A相视在功率输入值',
                  'B相视在功率输入值',
                  'C相视在功率输入值',
                  'A相进线有功能量输入值',
                  'A相出线有功能量输入值',
                  'B相进线有功能量输入值',
                  'B相出线有功能量输入值',
                  'C相进线有功能量输入值',
                  'C相出线有功能量输入值',
                  'A相进线无功能量输入值',
                  'A相出线无功能量输入值',
                  'B相进线无功能量输入值',
                  'B相出线无功能量输入值',
                  'C相进线无功能量输入值',
                  'C相出线无功能量输入值',
                  'A相进线视在能量输入值',
                  'A相出线视在能量输入值',
                  'B相进线视在能量输入值',
                  'B相出线视在能量输入值',
                  'C相进线视在能量输入值',
                  'C相出线视在能量输入值',
                  'A相进线有功能量测量值',
                  'A相出线有功能量测量值',
                  'B相进线有功能量测量值',
                  'B相出线有功能量测量值',
                  'C相进线有功能量测量值',
                  'C相出线有功能量测量值',
                  'A相进线无功能量测量值',
                  'A相出线无功能量测量值',
                  'B相进线无功能量测量值',
                  'B相出线无功能量测量值',
                  'C相进线无功能量测量值',
                  'C相出线无功能量测量值',
                  'A相进线视在能量测量值',
                  'A相出线视在能量测量值',
                  'B相进线视在能量测量值',
                  'B相出线视在能量测量值',
                  'C相进线视在能量测量值',
                  'C相出线视在能量测量值',
                  'A相进线有功能量精度',
                  'A相出线有功能量精度',
                  'B相进线有功能量精度', 'B相出线有功能量精度',
                  'C相进线有功能量精度', 'C相出线有功能量精度',
                  'A相进线无功能量精度', 'A相出线无功能量精度', 'B相进线无功能量精度', 'B相出线无功能量精度',
                  'C相进线无功能量精度', 'C相出线无功能量精度',
                  'A相进线视在能量精度', 'A相出线视在能量精度', 'B相进线视在能量精度', 'B相出线视在能量精度',
                  'C相进线视在能量精度', 'C相出线视在能量精度',
                  '测试结果']
    write_data_csv(row=1, data=topic_list)
    for i in range(0, len(topic_list)):
        energy_sheet.write(0, i, topic_list[i])


def calculate_pqs(result):
    result.append(result[0] * result[3] * (cos(result[6] - result[9])))
    result.append(result[1] * result[4] * (cos(result[7] - result[10])))
    result.append(result[2] * result[5] * (cos(result[8] - result[11])))
    result.append(result[0] * result[3] * (sin(result[6] - result[9])))
    result.append(result[1] * result[4] * (sin(result[7] - result[10])))
    result.append(result[2] * result[5] * (sin(result[8] - result[11])))
    result.append(result[0] * result[3])
    result.append(result[1] * result[4])
    result.append(result[2] * result[5])


def calculate_import_energy(result, line):
    for i in range(0, 18, 2):
        result.append(abs(result[i // 2 + 12]) * test_data[line][13])
        result.append(0.0)


def calculate_export_energy(result, line):
    for i in range(1, 18, 2):
        result.append(0.0)
        result.append(abs(result[i // 2 + 12]) * test_data[line][13])


def append_realtime_basicpara(result, realtime_basicpara, phase_angle):
    result.append(realtime_basicpara['Ua_rms'])
    result.append(realtime_basicpara['Ub_rms'])
    result.append(realtime_basicpara['Uc_rms'])
    result.append(realtime_basicpara['Ia_rms'])
    result.append(realtime_basicpara['Ib_rms'])
    result.append(realtime_basicpara['Ic_rms'])
    result.append(phase_angle['Ua_phase_angle'])
    result.append(phase_angle['Ub_phase_angle'])
    result.append(phase_angle['Uc_phase_angle'])
    result.append(phase_angle['Ia_phase_angle'])
    result.append(phase_angle['Ib_phase_angle'])
    result.append(phase_angle['Ic_phase_angle'])
    result.append(realtime_basicpara['Pa_rms'])
    result.append(realtime_basicpara['Pb_rms'])
    result.append(realtime_basicpara['Pc_rms'])
    result.append(realtime_basicpara['Qa_rms'])
    result.append(realtime_basicpara['Qb_rms'])
    result.append(realtime_basicpara['Qc_rms'])
    result.append(realtime_basicpara['Sa_rms'])
    result.append(realtime_basicpara['Sa_rms'])
    result.append(realtime_basicpara['Sa_rms'])


def add_test_accuracy(result, mode=0):
    if mode == 0:
        for i in range(0, 21):
            if result[i] == 0:
                result.append(0)
                continue
            result.append(abs(result[i] - result[i + 21]) / abs(result[i]) * 100)
        return
    for i in range(0, 18):
        if result[i + 21] == 0:
            result.append(0)
            continue
        result.append(abs(result[i + 21] - result[i + 21 + 18]) / abs(result[i + 21]) * 100)


def check_test_result(result, mode=0):
    ret = []
    if mode == 0:
        for i in range(0, 2):
            if abs(result[42 + i]) <= precision_requirement[1][1]:
                ret.append("passed")
            ret.append("failed")
        for i in range(0, 2):
            if abs(result[45 + i]) <= precision_requirement[2][1]:
                ret.append("passed")
            ret.append("failed")
        for i in range(0, 5):
            if abs(result[48 + i]) <= precision_requirement[3][1]:
                ret.append("passed")
            ret.append("failed")
        for i in range(0, 8):
            if abs(result[54 + i]) <= precision_requirement[4][1]:
                ret.append("passed")
            ret.append("failed")
        if "failed" in ret:
            result.append("failed")
            return
        result.append('passed')
        return
    for i in range(0, 18):
        if abs(result[57 + i]) <= precision_requirement[5][1]:
            ret.append("passed")
        ret.append("failed")
    if 'failed' in ret:
        result.append("failed")
        return
    result.append('passed')
    return


def add_energy_measure(result):
    ret = get_energy_continue()
    result.append(ret['Epa_imp'])
    result.append(ret['Epa_exp'])
    result.append(ret['Epb_imp'])
    result.append(ret['Epb_exp'])
    result.append(ret['Epc_imp'])
    result.append(ret['Epc_exp'])
    result.append(ret['Eqa_imp'])
    result.append(ret['Eqa_exp'])
    result.append(ret['Eqb_imp'])
    result.append(ret['Eqb_exp'])
    result.append(ret['Eqc_imp'])
    result.append(ret['Eqc_exp'])
    ret = get_apparent_energy()
    result.append(ret['Esa_imp'])
    result.append(ret['Esa_exp'])
    result.append(ret['Esb_imp'])
    result.append(ret['Esb_exp'])
    result.append(ret['Esc_imp'])
    result.append(ret['Esc_exp'])


def generate_test_report():
    num = 0
    write_excel_topic()
    print("测试进度：{}".format(test_data[0]))
    for i in range(1, len(test_data)):
        if test_data[i][13] == 0:
            num += 1
            result = []
            sheet.write(i, 0, test_data[i][0])
            for j in range(1, 13):
                result.append(test_data[i][j])
            # set_ac(ua=test_data[i][1], ub=test_data[i][2], uc=test_data[i][3], ia=test_data[i][4], ib=test_data[i][5],
            #        ic=test_data[i][6], qua=test_data[i][7], qub=test_data[i][8], quc=test_data[i][9],
            #        qia=test_data[i][10], qib=test_data[i][11], qic=test_data[i][12], f=60)
            realtime_basicpara = {
                'Ua_rms': get_ua(test_data[i][1]),
                'Ub_rms': get_ub(test_data[i][2]),
                'Uc_rms': get_uc(test_data[i][3]),
                'Ia_rms': get_ia(test_data[i][4]),
                'Ib_rms': get_ib(test_data[i][5]),
                'Ic_rms': get_ic(test_data[i][6]),
                'Pa_rms': get_pa(result[0] * result[3] * (cos(result[6] - result[9]))),
                'Pb_rms': get_pb(result[1] * result[4] * (cos(result[7] - result[10]))),
                'Pc_rms': get_pc(result[2] * result[5] * (cos(result[8] - result[11]))),
                'Qa_rms': get_qa(result[0] * result[3] * (sin(result[6] - result[9]))),
                'Qb_rms': get_qb(result[1] * result[4] * (sin(result[7] - result[10]))),
                'Qc_rms': get_qc(result[2] * result[5] * (sin(result[8] - result[11]))),
                'Sa_rms': get_sa(result[0] * result[3]),
                'Sb_rms': get_sb(result[1] * result[4]),
                'Sc_rms': get_sc(result[2] * result[5]),
            }
            phase_angle = {
                'Ua_phase_angle': 0.0,
                'Ub_phase_angle': get_ub_phase(test_data[i][8]),
                'Uc_phase_angle': get_uc_phase(test_data[i][9]),
                'Ia_phase_angle': get_ia_phase(test_data[i][10]),
                'Ib_phase_angle': get_ib_phase(test_data[i][11]),
                'Ic_phase_angle': get_ic_phase(test_data[i][12]),
            }

            calculate_pqs(result=result)
            append_realtime_basicpara(result=result, realtime_basicpara=realtime_basicpara, phase_angle=phase_angle)
            add_test_accuracy(result=result)
            check_test_result(result=result)
            logging.info("{} result is:{}".format(test_data[i][0], result))
            result.insert(0, test_data[i][0])
            write_data_csv(row=0, data=result)
            for index, element in enumerate(result):
                sheet.write(i, index + 1, element)
            print("测试进度：{}".format(test_data[i]))
            continue
        else:
            result = []
            energy_sheet.write(i - num, 0, test_data[i][0])
            for j in range(1, 13):
                result.append(test_data[i][j])
            calculate_pqs(result=result)
            assert clear_energy() is True
            cur_time = time.time()
            # set_ac(ua=test_data[i][1], ub=test_data[i][2], uc=test_data[i][3], ia=test_data[i][4],
            #        ib=test_data[i][5],
            #        ic=test_data[i][6], qua=test_data[i][7], qub=test_data[i][8], quc=test_data[i][9],
            #        qia=test_data[i][10], qib=test_data[i][11], qic=test_data[i][12], f=60)
            # while True:
            #     if (time.time() - cur_time) / 3600 >= test_data[i][13]:
            #         break
            # set_ac(ua=0, ub=0, uc=0, ia=0, ib=0, ic=0, qua=0, qub=0, quc=0, qia=0, qib=0, qic=0, f=0)
            if test_data[i][4] >= 0:
                calculate_import_energy(result=result, line=i)
                add_energy_measure(result=result)
                add_test_accuracy(result=result, mode=test_data[i][13])
                check_test_result(result=result, mode=test_data[i][13])
                for index, element in enumerate(result):
                    energy_sheet.write(i - num, index + 1, element)
                result.insert(0, test_data[i][0])
                write_data_csv(row=1, data=result)
                print("测试进度：{}".format(test_data[i]))
                continue
            calculate_export_energy(result=result, line=i)
            add_energy_measure(result=result)
            add_test_accuracy(result=result, mode=test_data[i][13])
            check_test_result(result=result, mode=test_data[i][13])
            for index, element in enumerate(result):
                energy_sheet.write(i - num, index + 1, element)
            result.insert(0, test_data[i][0])
            write_data_csv(row=1, data=result)
            print("测试进度：{}".format(test_data[i]))

    client.close()
    my_workbook.save('AcuvimⅡ_Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


if __name__ == '__main__':
    print('====================Precision Measure Start====================')
    print('======================{}======================'.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    start_time = time.time()
    generate_test_report()
    print('====================测试总耗时:{}===================='.format(time.time() - start_time))
    print('====================={}====================='.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    print('====================Precision Measure End====================')

