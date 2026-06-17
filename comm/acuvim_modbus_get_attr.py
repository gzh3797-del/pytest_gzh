"""
文件名称: acuvim_modbus_get_attr.py
功能描述: AcuvimⅡ测量数据获取
创建日期: 2025-05-14
作者: 苏博
版本: v1.0
修改记录: 2025-05-14 苏博 创建脚本
"""

from comm.modbus_rtu_tcp import ModbusRtuOrTcp
import struct
import time
import numpy as np
import logging

client = ModbusRtuOrTcp()


def clear_energy():
    """
    清理能量
    :return: True表示清理成功，False表示清理失败
    """
    ret = client.write_registers(address=0x1016, values=[1], slave=1)
    if '(4118,1)' not in str(ret):
        logging.error('clear energy fail, ret is:{}'.format(ret))
        return False
    ret = client.read_measurement(address=0x1016, count=1, slave=1)
    client.close()
    if ret[0] == 0:
        return True
    return False


def get_energy():
    energy_key_list = ['Ep_Imp', 'Ep_Exp', 'Eq_Imp', 'Eq_Exp', 'Ep_sum', 'Ep_net', 'Eq_sum', 'Eq_net', 'Es']
    energy_value_list = []
    ret = {}

    energy: list = client.read_measurement(address=0x4048, count=18, slave=1)

    for index in range(0, len(energy) - 1, 2):
        reg = hex(energy[index]).replace('0x', '').zfill(4) + hex(energy[index + 1]).replace('0x', '').zfill(4)
        integer_num = int(reg, 16)
        energy_value_list.append(float(struct.unpack('<I', struct.pack('<I', integer_num))[0] / 10))
    for key, value in zip(energy_key_list, energy_value_list):
        ret[key] = value
    return ret


def get_energy_continue():
    energy_key_list = ['Epa_imp', 'Epa_exp', 'Epb_imp', 'Epb_exp', 'Epc_imp', 'Epc_exp', 'Eqa_imp', 'Eqa_exp',
                       'Eqb_imp', 'Eqb_exp', 'Eqc_imp', 'Eqc_exp', 'Esa', 'Esb', 'Esc']
    energy_value_list = []
    ret = {}
    energy: list = client.read_measurement(address=0x4620, count=30, slave=1)
    for index in range(0, len(energy) - 1, 2):
        reg = hex(energy[index]).replace('0x', '').zfill(4) + hex(energy[index + 1]).replace('0x', '').zfill(4)
        integer_num = int(reg, 16)
        energy_value_list.append(float(struct.unpack('<I', struct.pack('<I', integer_num))[0] / 10))
    for key, value in zip(energy_key_list, energy_value_list):
        ret[key] = value
    return ret


def get_apparent_energy():
    energy_key_list = ['Es_imp ', 'Esa_imp', 'Esb_imp', 'Esc_imp', 'Es_exp', 'Esa_exp', 'Esb_exp', 'Esc_exp']
    energy_value_list = []
    ret = {}
    energy: list = client.read_measurement(address=0x4900, count=16, slave=1)
    for index in range(0, len(energy) - 1, 2):
        reg = hex(energy[index]).replace('0x', '').zfill(4) + hex(energy[index + 1]).replace('0x', '').zfill(4)
        integer_num = int(reg, 16)
        energy_value_list.append(float(struct.unpack('<I', struct.pack('<I', integer_num))[0] / 10))
    for key, value in zip(energy_key_list, energy_value_list):
        ret[key] = value
    return ret


def get_realtime_basicpara():
    energy_key_list = ["Freq_rms", "Ua_rms", "Ub_rms", "Uc_rms", "Uvag_rms", "Uab_rms", "Ubc_rms", "Uca_rms",
                       "Ulag_rms", "Ia_rms", "Ib_rms", "Ic_rms", "Ivag_rms", "In_rms", "Pa_rms", "Pb_rms", "Pc_rms",
                       "P_rms", "Qa_rms", "Qb_rms", "Qc_rms", "Q_rms", "Sa_rms", "Sb_rms", "Sc_rms", "S_rms", "PFa_rms",
                       "PFb_rms", "PFc_rms", "PF_rms", "AI1 value， fast read", "AI2 value， fast read",
                       "AI3 value， fast read", "AI4 value， fast read"]
    energy_value_list = []
    ret = {}
    energy: list = client.read_measurement(address=0x3000, count=68, slave=1)
    for index in range(0, len(energy) - 1, 2):
        reg = hex(energy[index]).replace('0x', '').zfill(4) + hex(energy[index + 1]).replace('0x', '').zfill(4)
        integer_num = int(reg, 16)
        energy_value_list.append(float(struct.unpack('!f', struct.pack('!I', integer_num))[0]))
    for key, value in zip(energy_key_list, energy_value_list):
        ret[key] = value
    return ret


def get_phase_angle():
    energy_key_list = ['Ua_phase_angle', 'Ub_phase_angle', 'Uc_phase_angle', 'Ia_phase_angle', 'Ib_phase_angle',
                       'Ic_phase_angle']
    energy_value_list = [0.0]
    ret = {}
    energy: list = client.read_measurement(address=0x42A0, count=5, slave=1)
    for index in range(0, len(energy)):
        energy_value_list.append(energy[index] / 10)
    for key, value in zip(energy_key_list, energy_value_list):
        ret[key] = value
    return ret


def get_float_attr(address, standard_value):
    ret_list = []
    for v in range(5):
        time.sleep(0.1)
        ret = client.read_measurement(address=address, count=2, slave=1)
        reg = hex(ret[0]).replace('0x', '').zfill(4) + hex(ret[1]).replace('0x', '').zfill(4)
        hex_num = reg.replace('0x', '')
        integer_num = int(hex_num, 16)
        voltage_measu = struct.unpack('!f', struct.pack('!I', integer_num))[0]
        ret_list.append(voltage_measu)
    ret_list.sort()
    mean = np.mean(ret_list)
    var = np.var(ret_list)
    if abs(ret_list[-1] - standard_value) > abs(ret_list[0] - standard_value):
        return ret_list[-1], mean, var
    return ret_list[0], mean, var


def get_word_attr(address, standard_value):
    ret_list = []
    for v in range(5):
        time.sleep(0.1)
        ret = client.read_measurement(address=address, count=1, slave=1)
        ret_list.append(ret[0] / 10)
    ret_list.sort()
    mean = np.mean(ret_list)
    var = np.var(ret_list)
    if abs(ret_list[-1] - standard_value) > abs(ret_list[0] - standard_value):
        return ret_list[-1], mean, var
    return ret_list[0], mean, var


def get_ua(standard_value):
    return get_float_attr(address=0x3002, standard_value=standard_value)[0]


def get_ub(standard_value):
    return get_float_attr(address=0x3004, standard_value=standard_value)[0]


def get_uc(standard_value):
    return get_float_attr(address=0x3006, standard_value=standard_value)[0]


def get_ia(standard_value):
    return get_float_attr(address=0x3012, standard_value=standard_value)[0]


def get_ib(standard_value):
    return get_float_attr(address=0x3014, standard_value=standard_value)[0]


def get_ic(standard_value):
    return get_float_attr(address=0x3016, standard_value=standard_value)[0]


def get_pa(standard_value):
    return get_float_attr(address=0x301C, standard_value=standard_value)[0]


def get_pb(standard_value):
    return get_float_attr(address=0x301E, standard_value=standard_value)[0]


def get_pc(standard_value):
    return get_float_attr(address=0x3020, standard_value=standard_value)[0]


def get_qa(standard_value):
    return get_float_attr(address=0x3024, standard_value=standard_value)[0]


def get_qb(standard_value):
    return get_float_attr(address=0x3026, standard_value=standard_value)[0]


def get_qc(standard_value):
    return get_float_attr(address=0x3028, standard_value=standard_value)[0]


def get_sa(standard_value):
    return get_float_attr(address=0x302C, standard_value=standard_value)[0]


def get_sb(standard_value):
    return get_float_attr(address=0x302E, standard_value=standard_value)[0]


def get_sc(standard_value):
    return get_float_attr(address=0x3030, standard_value=standard_value)[0]


def get_ub_phase(standard_value):
    return get_word_attr(address=0x42A0, standard_value=standard_value)[0]


def get_uc_phase(standard_value):
    return get_word_attr(address=0x42A1, standard_value=standard_value)[0]


def get_ia_phase(standard_value):
    return get_word_attr(address=0x42A2, standard_value=standard_value)[0]


def get_ib_phase(standard_value):
    return get_word_attr(address=0x42A3, standard_value=standard_value)[0]


def get_ic_phase(standard_value):
    return get_word_attr(address=0x42A4, standard_value=standard_value)[0]
