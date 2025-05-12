from comm.modbus_rtu_tcp import *
from comm.source_control import *
from tools.log import Log
import xlwt
from tools.excel_operate import data_read
from modbus_config import modbus_config

volt_cur_list = data_read(r'../../../comm/test_data/test_case.xlsx', 'test_data')
Log(str(__file__).split("\\")[-1])
my_workbook = xlwt.Workbook()
sheet = my_workbook.add_sheet('precision para', cell_overwrite_ok=True)
mes = ModbusRtuOrTcp(conn_mode=modbus_config['conn_mode'])


def read_vol(standard_value, times=1):
    vol_list = []
    for v in range(times):
        voltages = mes.read_measurement(address=12288, count=2, slave=1)
        logging.info('voltage ret is:{}'.format(voltages))
        reg = hex(voltages[0]).replace('0x', '').zfill(4) + hex(voltages[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        voltage_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        vol_list.append(voltage_measu)
    vol_list.sort()
    if abs(vol_list[-1] - standard_value) > abs(vol_list[0] - standard_value):
        return vol_list[-1]
    return vol_list[0]


def read_cur(standard_value, times=1):
    vol_list = []
    for v in range(times):
        voltages = mes.read_measurement(address=12290, count=2, slave=1)
        logging.info('voltage ret is:{}'.format(voltages))
        reg = hex(voltages[0]).replace('0x', '').zfill(4) + hex(voltages[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        voltage_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        vol_list.append(voltage_measu)
    vol_list.sort()
    if abs(vol_list[-1] - standard_value) > abs(vol_list[0] - standard_value):
        return vol_list[-1]
    return vol_list[0]


def read_pow(standard_value, times=1):
    vol_list = []
    for v in range(times):
        voltages = mes.read_measurement(address=12292, count=2, slave=1)
        logging.info('voltage ret is:{}'.format(voltages))
        reg = hex(voltages[0]).replace('0x', '').zfill(4) + hex(voltages[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        voltage_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        vol_list.append(voltage_measu)
    vol_list.sort()
    if abs(vol_list[-1] - standard_value) > abs(vol_list[0] - standard_value):
        return vol_list[-1]
    return vol_list[0]


def read_energy_charge():
    energy_charge: list = mes.read_measurement(address=16384, count=32, slave=1)
    import_energy = hex(energy_charge[0]).replace('0x', '').zfill(4) + hex(energy_charge[1]).replace('0x', '').zfill(
        4) + hex(energy_charge[2]).replace('0x', '').zfill(4) + hex(energy_charge[3]).replace('0x', '').zfill(4)
    integer_num = int(import_energy, 16)
    import_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    export_energy = hex(energy_charge[4]).replace('0x', '').zfill(4) + hex(energy_charge[5]).replace('0x', '').zfill(
        4) + hex(energy_charge[6]).replace('0x', '').zfill(4) + hex(energy_charge[7]).replace('0x', '').zfill(4)
    integer_num = int(export_energy, 16)
    export_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    net_energy = hex(energy_charge[8]).replace('0x', '').zfill(4) + hex(energy_charge[9]).replace('0x', '').zfill(
        4) + hex(energy_charge[10]).replace('0x', '').zfill(4) + hex(energy_charge[11]).replace('0x', '').zfill(4)
    integer_num = int(net_energy, 16)
    net_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    total_energy = hex(energy_charge[12]).replace('0x', '').zfill(4) + hex(energy_charge[13]).replace('0x', '').zfill(
        4) + hex(energy_charge[14]).replace('0x', '').zfill(4) + hex(energy_charge[15]).replace('0x', '').zfill(4)
    integer_num = int(total_energy, 16)
    total_energy = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    import_charge = hex(energy_charge[16]).replace('0x', '').zfill(4) + hex(energy_charge[17]).replace('0x', '').zfill(
        4) + hex(energy_charge[18]).replace('0x', '').zfill(4) + hex(energy_charge[19]).replace('0x', '').zfill(4)
    integer_num = int(import_charge, 16)
    import_charge = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    export_charge = hex(energy_charge[20]).replace('0x', '').zfill(4) + hex(energy_charge[21]).replace('0x', '').zfill(
        4) + hex(energy_charge[22]).replace('0x', '').zfill(4) + hex(energy_charge[23]).replace('0x', '').zfill(4)
    integer_num = int(export_charge, 16)
    export_charge = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    net_charge = hex(energy_charge[24]).replace('0x', '').zfill(4) + hex(energy_charge[25]).replace('0x', '').zfill(
        4) + hex(energy_charge[26]).replace('0x', '').zfill(4) + hex(energy_charge[27]).replace('0x', '').zfill(4)
    integer_num = int(net_charge, 16)
    net_charge = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    total_charge = hex(energy_charge[28]).replace('0x', '').zfill(4) + hex(energy_charge[29]).replace('0x', '').zfill(
        4) + hex(energy_charge[30]).replace('0x', '').zfill(4) + hex(energy_charge[31]).replace('0x', '').zfill(4)
    integer_num = int(total_charge, 16)
    total_charge = struct.unpack('<d', struct.pack('<Q', integer_num))[0]
    logging.info(
        'import_energy, export_energy, net_energy, total_energy, import_charge, export_charge, net_charge, total_charge ret is:{}'.format(
            (import_energy, export_energy, net_energy, total_energy, import_charge, export_charge, net_charge,
             total_charge)))
    return (import_energy, export_energy, net_energy, total_energy, import_charge, export_charge, net_charge,
            total_charge)


def clear_energy():
    ret = mes.write_registers(address=8192, values=[1], slave=1)
    if '(8192,1)' not in str(ret):
        logging.error('clear energy fail, ret is:{}'.format(ret))
        return False
    ret = mes.read_measurement(address=8192, count=1, slave=1)
    if ret[0] == 0:
        return True
    return False


def clear_charge():
    ret = mes.write_registers(address=8193, values=[1], slave=1)
    if '(8193,1)' not in str(ret):
        logging.error('clear charge fail, ret is:{}'.format(ret))
        return False
    ret = mes.read_measurement(address=8193, count=1, slave=1)
    if ret[0] == 0:
        return True
    return False


def get_enable_cable_loss_compensation():
    ret = mes.read_measurement(address=4130, count=1, slave=1)
    return ret[0]


def get_cable_resistance():
    ret = mes.read_measurement(address=4131, count=1, slave=1)
    return ret[0]


def set_cable_loss_compensation_enable(value):
    ret = mes.write_registers(address=4130, values=[value], slave=1)
    if '(4130,1)' not in str(ret):
        logging.error('set cable loss compensation enable ret is:{}'.format(ret))
        return False
    ret = get_enable_cable_loss_compensation()
    if ret != value:
        logging.info('cable loss compensation enable ret is:{}'.format(ret[0]))
        return False
    return True


def set_cable_resistance(value):
    value *= 10000
    ret = mes.write_registers(address=4131, values=[int(value)], slave=1)
    if '(4131,1)' not in str(ret):
        logging.error('set cable resistance ret is:{}'.format(ret))
        return False
    ret = get_cable_resistance()
    if ret / 10000 != value:
        logging.info('cable resistance ret is:{}'.format(ret))
        return False
    return True


def run():
    sheet.write(0, 0, '测试用例')
    sheet.write(0, 1, '电压输入值')
    sheet.write(0, 2, '电流输入值')
    sheet.write(0, 3, '功率输入值')
    sheet.write(0, 4, '进线能量输入值')
    sheet.write(0, 5, '进线电荷输入值')
    sheet.write(0, 6, '出线能量输入值')
    sheet.write(0, 7, '出线电荷输入值')
    sheet.write(0, 8, '电压测量值')
    sheet.write(0, 9, '电流测量值')
    sheet.write(0, 10, '功率测量值')
    sheet.write(0, 11, '进线能量测量值')
    sheet.write(0, 12, '进线电荷测量值')
    sheet.write(0, 13, '出线能量测量值')
    sheet.write(0, 14, '出线电荷测量值')
    sheet.write(0, 15, '电压精度')
    sheet.write(0, 16, '电流精度')
    sheet.write(0, 17, '功率精度')
    sheet.write(0, 18, '进线能量精度')
    sheet.write(0, 19, '进线电荷精度')
    sheet.write(0, 20, '出线能量精度')
    sheet.write(0, 21, '出线电荷精度')
    sheet.write(0, 22, '测试结果')
    for i in range(len(volt_cur_list)):
        if i == 0:
            logging.info('测试进度:{}'.format(volt_cur_list[i]))
            print('测试进度:{}'.format(volt_cur_list[i]))
        else:
            logging.info('测试进度:{},执行时间:{}'.format(volt_cur_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
            print('测试进度:{},执行时间:{}'.format(volt_cur_list[i], time.strftime('%Y_%m_%d %H:%M:%S')))
        if i == len(volt_cur_list) - 1:
            break
        sheet.write(i + 1, 0, volt_cur_list[i + 1][0])
        sheet.write(i + 1, 1, volt_cur_list[i + 1][1])
        sheet.write(i + 1, 2, volt_cur_list[i + 1][2])
        if volt_cur_list[i + 1][2] == 'null':
            sheet.write(i + 1, 9, 'null')
            sheet.write(i + 1, 10, 'null')
            sheet.write(i + 1, 16, 'null')
            sheet.write(i + 1, 17, 'null')
            sheet.write(i + 1, 3, 'null')
            sheet.write(i + 1, 4, 'null')
            sheet.write(i + 1, 5, 'null')
            sheet.write(i + 1, 6, 'null')
            sheet.write(i + 1, 7, 'null')
            sheet.write(i + 1, 11, 'null')
            sheet.write(i + 1, 12, 'null')
            sheet.write(i + 1, 13, 'null')
            sheet.write(i + 1, 14, 'null')
            sheet.write(i + 1, 18, 'null')
            sheet.write(i + 1, 19, 'null')
            sheet.write(i + 1, 20, 'null')
            sheet.write(i + 1, 21, 'null')
            sour_output(voltage=volt_cur_list[i + 1][1], current=0)
            vol_meas = read_vol(volt_cur_list[i + 1][1], times=20)
            vol_precision = abs(volt_cur_list[i + 1][1] - vol_meas) / abs(volt_cur_list[i + 1][1]) * 100
            sheet.write(i + 1, 8, vol_meas)
            sheet.write(i + 1, 15, vol_precision)
            sheet.write(i + 1, 22, 'passed' if vol_precision < volt_cur_list[i + 1][4] else 'failed')
            sour_stop()
            time.sleep(3)
            continue

        if volt_cur_list[i + 1][1] == 'null':
            sheet.write(i + 1, 8, 'null')
            sheet.write(i + 1, 10, 'null')
            sheet.write(i + 1, 15, 'null')
            sheet.write(i + 1, 17, 'null')
            sheet.write(i + 1, 3, 'null')
            sheet.write(i + 1, 4, 'null')
            sheet.write(i + 1, 5, 'null')
            sheet.write(i + 1, 6, 'null')
            sheet.write(i + 1, 7, 'null')
            sheet.write(i + 1, 11, 'null')
            sheet.write(i + 1, 12, 'null')
            sheet.write(i + 1, 13, 'null')
            sheet.write(i + 1, 14, 'null')
            sheet.write(i + 1, 18, 'null')
            sheet.write(i + 1, 19, 'null')
            sheet.write(i + 1, 20, 'null')
            sheet.write(i + 1, 21, 'null')
            if volt_cur_list[i + 1][2] < 0:
                time.sleep(3)
                sour_output(voltage=0, current=volt_cur_list[i + 1][2])
            else:
                time.sleep(3)
                sour_output(voltage=0, current=volt_cur_list[i + 1][2])
            cur_meas = read_cur(volt_cur_list[i + 1][2], times=20)
            if volt_cur_list[i + 1][2] == 0:
                sheet.write(i + 1, 9, cur_meas)
                sheet.write(i + 1, 16, 'null')
                sheet.write(i + 1, 22, 'passed' if cur_meas == 0 else 'failed')
                continue
            cur_precision = abs(abs(volt_cur_list[i + 1][2]) - abs(cur_meas)) / abs(volt_cur_list[i + 1][2]) * 100
            if abs(volt_cur_list[i + 1][2]) <= 3.9 or abs(volt_cur_list[i + 1][2]) > 650:
                sheet.write(i + 1, 9, cur_meas)
                sheet.write(i + 1, 16, cur_precision)
                sheet.write(i + 1, 22, 'passed')
                continue
            sheet.write(i + 1, 9, cur_meas)
            sheet.write(i + 1, 16, cur_precision)
            sheet.write(i + 1, 22, 'passed' if cur_precision < volt_cur_list[i + 1][4] else 'failed')
            sour_stop()
            continue

        if volt_cur_list[i + 1][3] == 'null':
            sheet.write(i + 1, 4, 'null')
            sheet.write(i + 1, 5, 'null')
            sheet.write(i + 1, 6, 'null')
            sheet.write(i + 1, 7, 'null')
            sheet.write(i + 1, 11, 'null')
            sheet.write(i + 1, 12, 'null')
            sheet.write(i + 1, 13, 'null')
            sheet.write(i + 1, 14, 'null')
            sheet.write(i + 1, 18, 'null')
            sheet.write(i + 1, 19, 'null')
            sheet.write(i + 1, 20, 'null')
            sheet.write(i + 1, 21, 'null')
            pow_standard = volt_cur_list[i + 1][1] * volt_cur_list[i + 1][2] / 1000
            if volt_cur_list[i + 1][2] < 0:
                sour_output(voltage=volt_cur_list[i + 1][1], current=volt_cur_list[i + 1][2],
                            current_direction='反向')
            else:
                sour_output(voltage=volt_cur_list[i + 1][1], current=volt_cur_list[i + 1][2])
            vol_meas = read_vol(volt_cur_list[i + 1][1], times=20)
            vol_precision = abs(volt_cur_list[i + 1][1] - vol_meas) / abs(volt_cur_list[i + 1][1]) * 100
            sheet.write(i + 1, 8, vol_meas)
            sheet.write(i + 1, 15, vol_precision)
            cur_meas = read_cur(volt_cur_list[i + 1][2], times=20)
            cur_precision = abs(abs(volt_cur_list[i + 1][2]) - abs(cur_meas)) / abs(volt_cur_list[i + 1][2]) * 100
            if abs(volt_cur_list[i + 1][2]) <= 3.9:
                sheet.write(i + 1, 9, cur_meas)
                sheet.write(i + 1, 16, cur_precision)
                pow_meas = read_pow(pow_standard, times=20)
                pow_precision = abs(abs(pow_standard) - abs(pow_meas)) / abs(pow_standard) * 100
                sheet.write(i + 1, 10, pow_meas)
                sheet.write(i + 1, 17, pow_precision)
                sheet.write(i + 1, 22, 'passed')
            else:
                sheet.write(i + 1, 9, cur_meas)
                sheet.write(i + 1, 16, cur_precision)
                pow_meas = read_pow(pow_standard, times=20)
                pow_precision = abs(abs(pow_standard) - abs(pow_meas)) / abs(pow_standard) * 100
                sheet.write(i + 1, 10, pow_meas)
                sheet.write(i + 1, 17, pow_precision)
                precision = [float(i) for i in volt_cur_list[i + 1][4].split(',')]
                sheet.write(i + 1, 22,
                            'passed' if pow_precision < precision[2] and cur_precision < precision[
                                1] and vol_precision <
                                        precision[0] else 'failed')
            sour_stop()
            time.sleep(3)
            power = float(volt_cur_list[i + 1][1]) * float(volt_cur_list[i + 1][2]) / 1000
            sheet.write(i + 1, 3, power)
            continue

        if volt_cur_list[i + 1][3] != 'null' and volt_cur_list[i + 1][5] == 'null':
            j = 0
            pow_standard = volt_cur_list[i + 1][1] * volt_cur_list[i + 1][2] / 1000
            vol_meas = []
            cur_meas = []
            pow_meas = []
            vol_precision = 0
            cur_precision = 0
            pow_precision = 0
            if volt_cur_list[i + 1][2] < 0:
                clear_energy()
                clear_charge()
                cur_time = time.time()
                sour_output(voltage=volt_cur_list[i + 1][1], current=volt_cur_list[i + 1][2],
                            current_direction='反向')
                while time.time() <= cur_time + volt_cur_list[i + 1][3] * 3600:
                    vol_meas.append(read_vol(volt_cur_list[i + 1][1]))
                    cur_meas.append(read_cur(volt_cur_list[i + 1][2]))
                    pow_meas.append(read_pow(pow_standard))
                    j += 1
                sour_stop()
                vol_preci_max = max(vol_meas) if abs(max(vol_meas) - volt_cur_list[i + 1][1]) > abs(
                    min(vol_meas) - volt_cur_list[i + 1][1]) else min(vol_meas)
                vol_precision = abs(volt_cur_list[i + 1][1] - abs(vol_preci_max)) / abs(volt_cur_list[i + 1][1]) * 100
                sheet.write(i + 1, 8, vol_preci_max)
                sheet.write(i + 1, 15, vol_precision)
                cur_preci_max = max(cur_meas) if abs(abs(max(cur_meas)) - abs(volt_cur_list[i + 1][2])) > abs(
                    abs(min(cur_meas)) - abs(volt_cur_list[i + 1][2])) else min(cur_meas)
                cur_precision = abs(abs(volt_cur_list[i + 1][2]) - abs(cur_preci_max)) / abs(
                    volt_cur_list[i + 1][2]) * 100
                if abs(volt_cur_list[i + 1][2]) <= 3.9:
                    sheet.write(i + 1, 9, cur_preci_max)
                    sheet.write(i + 1, 16, cur_precision)
                    pow_preci_max = max(pow_meas) if abs(abs(max(pow_meas)) - abs(pow_standard)) > abs(
                        abs(min(pow_meas)) - abs(pow_standard)) else min(pow_meas)
                    pow_precision = abs(abs(pow_standard) - abs(pow_preci_max)) / abs(pow_standard) * 100
                    sheet.write(i + 1, 10, pow_preci_max)
                    sheet.write(i + 1, 17, pow_precision)
                else:
                    sheet.write(i + 1, 9, cur_preci_max)
                    sheet.write(i + 1, 16, cur_precision)
                    pow_preci_max = max(pow_meas) if abs(abs(max(pow_meas)) - abs(pow_standard)) > abs(
                        abs(min(pow_meas)) - abs(pow_standard)) else min(pow_meas)
                    pow_precision = abs(abs(pow_standard) - abs(pow_preci_max)) / abs(pow_standard) * 100
                    sheet.write(i + 1, 10, pow_preci_max)
                    sheet.write(i + 1, 17, pow_precision)
                time.sleep(10)
                energy = float(volt_cur_list[i + 1][1]) * float(volt_cur_list[i + 1][2]) * float(
                    volt_cur_list[i + 1][3]) / 1000
                charge = float(volt_cur_list[i + 1][2]) * float(volt_cur_list[i + 1][3])
                sheet.write(i + 1, 4, 0)
                sheet.write(i + 1, 5, 0)
                sheet.write(i + 1, 6, abs(energy))
                sheet.write(i + 1, 7, abs(charge))
                energy_charge = read_energy_charge()
                sheet.write(i + 1, 11, energy_charge[0])
                sheet.write(i + 1, 12, energy_charge[4])
                sheet.write(i + 1, 13, energy_charge[1])
                sheet.write(i + 1, 14, energy_charge[5])
                sheet.write(i + 1, 18, 0)
                sheet.write(i + 1, 19, 0)
                energy_precision = abs(abs(energy) - abs(energy_charge[1])) / abs(energy) * 100
                charge_precision = abs(abs(charge) - abs(energy_charge[5])) / abs(charge) * 100
                sheet.write(i + 1, 20, energy_precision)
                sheet.write(i + 1, 21, charge_precision)
                precision = [float(i) for i in volt_cur_list[i + 1][4].split(',')]
                sheet.write(i + 1, 22, 'passed' if energy_precision < precision[0] and charge_precision < precision[
                    1] else 'failed')
            else:
                clear_energy()
                clear_charge()
                cur_time = time.time()
                sour_output(voltage=volt_cur_list[i + 1][1], current=volt_cur_list[i + 1][2])
                while time.time() <= cur_time + volt_cur_list[i + 1][3] * 3600:
                    vol_meas.append(read_vol(volt_cur_list[i + 1][1]))
                    cur_meas.append(read_cur(volt_cur_list[i + 1][2]))
                    pow_meas.append(read_pow(pow_standard))
                    j += 1
                sour_stop()
                vol_preci_max = max(vol_meas) if abs(max(vol_meas) - volt_cur_list[i + 1][1]) > abs(
                    min(vol_meas) - volt_cur_list[i + 1][1]) else min(vol_meas)
                vol_precision = abs(volt_cur_list[i + 1][1] - abs(vol_preci_max)) / abs(volt_cur_list[i + 1][1]) * 100
                sheet.write(i + 1, 8, vol_preci_max)
                sheet.write(i + 1, 15, vol_precision)
                cur_preci_max = max(cur_meas) if abs(abs(max(cur_meas)) - abs(volt_cur_list[i + 1][2])) > abs(
                    abs(min(cur_meas)) - abs(volt_cur_list[i + 1][2])) else min(cur_meas)
                cur_precision = abs(abs(volt_cur_list[i + 1][2]) - abs(cur_preci_max)) / abs(
                    volt_cur_list[i + 1][2]) * 100
                if abs(volt_cur_list[i + 1][2]) <= 3.9:
                    sheet.write(i + 1, 9, cur_preci_max)
                    sheet.write(i + 1, 16, cur_precision)
                    pow_preci_max = max(pow_meas) if abs(abs(max(pow_meas)) - abs(pow_standard)) > abs(
                        abs(min(pow_meas)) - abs(pow_standard)) else min(pow_meas)
                    pow_precision = abs(abs(pow_standard) - abs(pow_preci_max)) / abs(pow_standard) * 100
                    sheet.write(i + 1, 10, pow_preci_max)
                    sheet.write(i + 1, 17, pow_precision)
                else:
                    sheet.write(i + 1, 9, cur_preci_max)
                    sheet.write(i + 1, 16, cur_precision)
                    pow_preci_max = max(pow_meas) if abs(abs(max(pow_meas)) - abs(pow_standard)) > abs(
                        abs(min(pow_meas)) - abs(pow_standard)) else min(pow_meas)
                    pow_precision = abs(abs(pow_standard) - abs(pow_preci_max)) / abs(pow_standard) * 100
                    sheet.write(i + 1, 10, pow_preci_max)
                    sheet.write(i + 1, 17, pow_precision)
                time.sleep(10)
                energy = float(volt_cur_list[i + 1][1]) * float(volt_cur_list[i + 1][2]) * float(
                    volt_cur_list[i + 1][3]) / 1000
                charge = float(volt_cur_list[i + 1][2]) * float(volt_cur_list[i + 1][3])
                sheet.write(i + 1, 4, abs(energy))
                sheet.write(i + 1, 5, abs(charge))
                sheet.write(i + 1, 6, 0)
                sheet.write(i + 1, 7, 0)
                energy_charge = read_energy_charge()
                sheet.write(i + 1, 11, energy_charge[0])
                sheet.write(i + 1, 12, energy_charge[4])
                sheet.write(i + 1, 13, energy_charge[1])
                sheet.write(i + 1, 14, energy_charge[5])
                energy_precision = abs(abs(energy) - abs(energy_charge[0])) / abs(energy) * 100
                charge_precision = abs(abs(charge) - abs(energy_charge[4])) / abs(charge) * 100
                sheet.write(i + 1, 18, energy_precision)
                sheet.write(i + 1, 19, charge_precision)
                sheet.write(i + 1, 20, 0)
                sheet.write(i + 1, 21, 0)
                precision = [float(i) for i in volt_cur_list[i + 1][4].split(',')]
                sheet.write(i + 1, 22, 'passed' if energy_precision < precision[0] and charge_precision < precision[
                    1] else 'failed')
            power = float(volt_cur_list[i + 1][1]) * float(volt_cur_list[i + 1][2]) / 1000
            sheet.write(i + 1, 3, power)
        if volt_cur_list[i + 1][5] != 'null':
            j = 0
            set_cable_loss_compensation_enable(1)
            set_cable_resistance(volt_cur_list[i + 1][5])
            print(volt_cur_list[i + 1][5])
            cable_vol = volt_cur_list[i + 1][1] - abs(volt_cur_list[i + 1][2]) * volt_cur_list[i + 1][5]
            pow_standard = cable_vol * volt_cur_list[i + 1][2] / 1000
            vol_meas = []
            cur_meas = []
            pow_meas = []
            vol_precision = 0
            cur_precision = 0
            pow_precision = 0
            if volt_cur_list[i + 1][2] < 0:
                clear_energy()
                clear_charge()
                cur_time = time.time()
                sour_output(voltage=volt_cur_list[i + 1][1], current=volt_cur_list[i + 1][2],
                            current_direction='反向')
                while time.time() <= cur_time + volt_cur_list[i + 1][3] * 3600:
                    vol_meas.append(read_vol(cable_vol))
                    cur_meas.append(read_cur(volt_cur_list[i + 1][2]))
                    pow_meas.append(read_pow(pow_standard))
                    # print(vol_meas[j], cur_meas[j], pow_meas[j])
                    j += 1
                sour_stop()
                vol_preci_max = max(vol_meas) if abs(max(vol_meas) - cable_vol) > abs(
                    min(vol_meas) - cable_vol) else min(vol_meas)
                vol_precision = abs(cable_vol - abs(vol_preci_max)) / abs(cable_vol) * 100
                sheet.write(i + 1, 8, vol_preci_max)
                sheet.write(i + 1, 15, vol_precision)
                cur_preci_max = max(cur_meas) if abs(abs(max(cur_meas)) - abs(volt_cur_list[i + 1][2])) > abs(
                    abs(min(cur_meas)) - abs(volt_cur_list[i + 1][2])) else min(cur_meas)
                cur_precision = abs(abs(volt_cur_list[i + 1][2]) - abs(cur_preci_max)) / abs(
                    volt_cur_list[i + 1][2]) * 100
                if abs(volt_cur_list[i + 1][2]) <= 3.9:
                    sheet.write(i + 1, 9, cur_preci_max)
                    sheet.write(i + 1, 16, cur_precision)
                    pow_preci_max = max(pow_meas) if abs(abs(max(pow_meas)) - abs(pow_standard)) > abs(
                        abs(min(pow_meas)) - abs(pow_standard)) else min(pow_meas)
                    pow_precision = abs(abs(pow_standard) - abs(pow_preci_max)) / abs(pow_standard) * 100
                    sheet.write(i + 1, 10, pow_preci_max)
                    sheet.write(i + 1, 17, pow_precision)
                else:
                    sheet.write(i + 1, 9, cur_preci_max)
                    sheet.write(i + 1, 16, cur_precision)
                    pow_preci_max = max(pow_meas) if abs(abs(max(pow_meas)) - abs(pow_standard)) > abs(
                        abs(min(pow_meas)) - abs(pow_standard)) else min(pow_meas)
                    pow_precision = abs(abs(pow_standard) - abs(pow_preci_max)) / abs(pow_standard) * 100
                    sheet.write(i + 1, 10, pow_preci_max)
                    sheet.write(i + 1, 17, pow_precision)
                time.sleep(10)
                energy = float(cable_vol) * float(volt_cur_list[i + 1][2]) * float(volt_cur_list[i + 1][3]) / 1000
                charge = float(volt_cur_list[i + 1][2]) * float(volt_cur_list[i + 1][3])
                sheet.write(i + 1, 4, 0)
                sheet.write(i + 1, 5, 0)
                sheet.write(i + 1, 6, abs(energy))
                sheet.write(i + 1, 7, abs(charge))
                energy_charge = read_energy_charge()
                sheet.write(i + 1, 11, energy_charge[0])
                sheet.write(i + 1, 12, energy_charge[4])
                sheet.write(i + 1, 13, energy_charge[1])
                sheet.write(i + 1, 14, energy_charge[5])
                sheet.write(i + 1, 18, 0)
                sheet.write(i + 1, 19, 0)
                energy_precision = abs(abs(energy) - abs(energy_charge[1])) / abs(energy) * 100
                charge_precision = abs(abs(charge) - abs(energy_charge[5])) / abs(charge) * 100
                sheet.write(i + 1, 20, energy_precision)
                sheet.write(i + 1, 21, charge_precision)
                precision = [float(i) for i in volt_cur_list[i + 1][4].split(',')]
                sheet.write(i + 1, 22, 'passed' if energy_precision < precision[3] and charge_precision < precision[
                    4] and vol_precision < precision[0] and cur_precision < precision[1] and pow_precision < precision[
                                                       2] else 'failed')
            else:
                clear_energy()
                clear_charge()
                cur_time = time.time()
                sour_output(voltage=volt_cur_list[i + 1][1], current=volt_cur_list[i + 1][2])
                while time.time() <= cur_time + volt_cur_list[i + 1][3] * 3600:
                    vol_meas.append(read_vol(cable_vol))
                    cur_meas.append(read_cur(volt_cur_list[i + 1][2]))
                    pow_meas.append(read_pow(pow_standard))
                    # print(vol_meas[j], cur_meas[j], pow_meas[j])
                    j += 1
                sour_stop()
                vol_preci_max = max(vol_meas) if abs(max(vol_meas) - cable_vol) > abs(
                    min(vol_meas) - cable_vol) else min(vol_meas)
                vol_precision = abs(cable_vol - abs(vol_preci_max)) / abs(cable_vol) * 100
                sheet.write(i + 1, 8, vol_preci_max)
                sheet.write(i + 1, 15, vol_precision)
                cur_preci_max = max(cur_meas) if abs(abs(max(cur_meas)) - abs(volt_cur_list[i + 1][2])) > abs(
                    abs(min(cur_meas)) - abs(volt_cur_list[i + 1][2])) else min(cur_meas)
                cur_precision = abs(abs(volt_cur_list[i + 1][2]) - abs(cur_preci_max)) / abs(
                    volt_cur_list[i + 1][2]) * 100
                if abs(volt_cur_list[i + 1][2]) <= 3.9:
                    sheet.write(i + 1, 9, cur_preci_max)
                    sheet.write(i + 1, 16, cur_precision)
                    pow_preci_max = max(pow_meas) if abs(abs(max(pow_meas)) - abs(pow_standard)) > abs(
                        abs(min(pow_meas)) - abs(pow_standard)) else min(pow_meas)
                    pow_precision = abs(abs(pow_standard) - abs(pow_preci_max)) / abs(pow_standard) * 100
                    sheet.write(i + 1, 10, pow_preci_max)
                    sheet.write(i + 1, 17, pow_precision)
                else:
                    sheet.write(i + 1, 9, cur_preci_max)
                    sheet.write(i + 1, 16, cur_precision)
                    pow_preci_max = max(pow_meas) if abs(abs(max(pow_meas)) - abs(pow_standard)) > abs(
                        abs(min(pow_meas)) - abs(pow_standard)) else min(pow_meas)
                    pow_precision = abs(abs(pow_standard) - abs(pow_preci_max)) / abs(pow_standard) * 100
                    sheet.write(i + 1, 10, pow_preci_max)
                    sheet.write(i + 1, 17, pow_precision)
                time.sleep(10)
                energy = float(cable_vol) * float(volt_cur_list[i + 1][2]) * float(volt_cur_list[i + 1][3]) / 1000
                charge = float(volt_cur_list[i + 1][2]) * float(volt_cur_list[i + 1][3])
                sheet.write(i + 1, 4, abs(energy))
                sheet.write(i + 1, 5, abs(charge))
                sheet.write(i + 1, 6, 0)
                sheet.write(i + 1, 7, 0)
                energy_charge = read_energy_charge()
                sheet.write(i + 1, 11, energy_charge[0])
                sheet.write(i + 1, 12, energy_charge[4])
                sheet.write(i + 1, 13, energy_charge[1])
                sheet.write(i + 1, 14, energy_charge[5])
                energy_precision = abs(abs(energy) - abs(energy_charge[0])) / abs(energy) * 100
                charge_precision = abs(abs(charge) - abs(energy_charge[4])) / abs(charge) * 100
                sheet.write(i + 1, 18, energy_precision)
                sheet.write(i + 1, 19, charge_precision)
                sheet.write(i + 1, 20, 0)
                sheet.write(i + 1, 21, 0)
                precision = [float(i) for i in volt_cur_list[i + 1][4].split(',')]
                sheet.write(i + 1, 22, 'passed' if energy_precision < precision[3] and charge_precision < precision[
                    4] and vol_precision < precision[0] and cur_precision < precision[1] and pow_precision < precision[
                                                       2] else 'failed')
            power = float(volt_cur_list[i + 1][1]) * float(volt_cur_list[i + 1][2]) / 1000
            sheet.write(i + 1, 3, power)
            set_cable_loss_compensation_enable(0)
    sour_stop()
    mes.close()
    my_workbook.save('Precision_Measure_{}.xls'.format(time.strftime('%Y%m%d%H%M%S')))


if __name__ == '__main__':
    print('====================Precision Measure Start====================')
    print('======================{}======================'.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    start_time = time.time()
    run()
    print('====================测试总耗时:{}===================='.format(time.time() - start_time))
    print('====================={}====================='.format(time.strftime('%Y_%m_%d %H:%M:%S')))
    print('====================Precision Measure End====================')
