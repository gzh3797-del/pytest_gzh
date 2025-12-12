import socket
import time
import logging
from modbus_config import modbus_config

import sys
import os
from time import sleep
from tools.log import Log


class SourceControlError(Exception):
    def __init__(self, msg):
        self.msg = msg


class SourCon:
    def __init__(self, timeout=5):
        self.udp_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.udp_socket.settimeout(timeout)
        self.udp_socket.bind((modbus_config['local']['ip'], modbus_config['local']['port']))
        self.dest_addr = (modbus_config['source']['ip'], modbus_config['source']['port'])

    def send(self, hex_data):
        send_data = hex_data.encode('gbk')
        logging.info(send_data.decode('gbk'))
        ret = self.udp_socket.sendto(send_data, self.dest_addr)
        return ret

    def recv(self):
        try:
            recv_data = self.udp_socket.recvfrom(1024)
        except TimeoutError:
            raise SourceControlError(
                'Source control timeout. Check whether the software of the control source is turned on.')
        return recv_data[0].decode('gbk')

    def close(self):
        self.udp_socket.close()


def sour_para_conf(input_method='直接', pluse_cons=500):
    data = ''
    if input_method == '直接':
        data = '''<参数配置>
        电流接入方式:直接接入式;
        供电方式:电源供电;
        额定电压:1000;
        标定电流:500;
        分流器额定:18mV;
        被检表阻抗:0.0000277Ω;
        脉冲常数:{};
        校验圈数:自动;
        校验秒数:1;
        <End>'''.format(pluse_cons)
    elif input_method == '间接':
        data = '''<参数配置>
        电流接入方式:间接接入式;
        供电方式:电源供电;
        额定电压:1000;
        标定电流:650;
        分流器额定:18mV;
        被检表阻抗:0.0000277Ω;
        脉冲常数:21600;
        校验圈数:自动;
        校验秒数:1;
        <End>'''
    else:
        logging.info('电流输入方式错误，请重新配置')
    re = SourCon()
    re.send(data)
    re.close()


def sour_output(voltage: float, current: float, stable_time=10):
    vol = voltage / 1000 * 100
    cur = current / 500 * 100
    data = '''<源输出>
    电压检定点:{}%;
    电流检定点:{}%;
    电压纹波比例:0%;
    电流纹波比例:0%;
    电压纹波相位:0度;
    电流纹波相位:0度;
    纹波频率:300Hz;
    电能方向:正向;
    <End>'''.format(str(vol), str(cur))
    re = SourCon()
    re.send(data)
    re.close()
    time.sleep(stable_time)


def mv_sour_output(voltage: float, current: float, shunt_rate=18, stable_time=10, current_direction='正向',
                   mv_flag=True):
    """
    控制源输出参数
    :param voltage: 源输出的电压
    :param current: 源输出的电流
    :param shunt_rate: Shunt的额定电压，仅在mV信号时使用，非mV信号禁止修改
    :param stable_time: 源输出到稳定的时间
    :param current_direction: 电能方向，仅在mV信号时使用，非mV信号禁止修改
    :param mv_flag: 是否使用mv信号的标志，True:需要将源的电流接入方式为间接接入式，False:源的电流接入方式为直接接入式
    :return:
    """
    if mv_flag is True:
        vol = voltage / 1000 * 100
        cur = current * 18 / 650 / shunt_rate * 100
    else:
        vol = voltage / 1000 * 100
        cur = current / 600 * 100
    data = '''<源输出>
    电压检定点:{}%;
    电流检定点:{}%;
    电压纹波比例:0%;
    电流纹波比例:0%;
    电压纹波相位:0度;
    电流纹波相位:0度;
    纹波频率:300Hz;
    电能方向:{};
    <End>'''.format(str(vol), str(cur), str(current_direction))
    re = SourCon()
    re.send(data)
    recive_ret = re.recv()
    re.close()
    expect_recive = '<源输出应答>'
    if expect_recive not in recive_ret:
        raise SourceControlError('Source control fail,Please check Environment.')
    logging.info('Source control success, voltage is:{}, current is:{}'.format(voltage, current))
    time.sleep(stable_time)


def sour_stop():
    data = '''<源停止>
    <End>'''
    re = SourCon()
    re.send(data)
    re.close()


def get_verification_error(times=5, timeout=15):
    data = '''<误差读取>    
        统计次数:{};
        <End>'''.format(times)
    re = SourCon(timeout=timeout)
    re.send(data)
    recive_ret = re.recv()
    re.close()
    return float(recive_ret.split(':')[1].split(';')[0])


print(get_verification_error())


class Cl3021SourCon:
    def __init__(self):
        self.udp_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.udp_socket.settimeout(3)
        logging.info(modbus_config['local']['ip'], modbus_config['local']['port'])
        self.udp_socket.bind((modbus_config['local']['ip'], modbus_config['local']['port']))
        self.dest_addr = (modbus_config['source']['ip'], modbus_config['source']['port'])

    def send(self, hex_data):
        ret = self.udp_socket.sendto(hex_data, self.dest_addr)
        recv_data = self.udp_socket.recvfrom(1024)
        return ret, recv_data

    def recv(self):
        try:
            recv_data = self.udp_socket.recvfrom(1024)
        except TimeoutError:
            raise SourceControlError(
                'Source control timeout. Check whether the software of the control source is turned on.')
        return recv_data[0]

    def close(self):
        self.udp_socket.close()


def bin_to_hex(binary):
    # 将二进制补足为4的倍数
    binary = binary.zfill((len(binary) + 3) // 4 * 4)

    # 将二进制数按4位一组分组
    groups = [binary[i:i + 4] for i in range(0, len(binary), 4)]

    # 将每个分组转换为十六进制数
    hex_str = ''
    for group in groups:
        hex_digit = hex(int(group, 2))[2:]  # 将二进制数转换为十六进制数
        hex_str += hex_digit

    return hex_str


def xor_sum(numbers):
    result = 0
    for number in numbers:
        result ^= number
    return result


def online():
    online_cmd = [0x81, 0x01, 0x25, 0x06, 0xc9, 0xeb]
    pdu = bytearray(online_cmd)
    source_control = Cl3021SourCon()
    source_control.send(pdu)
    source_control.close()


def set_wire(way: str):
    """
    设置接线方式
    :param way:
        0x08---->00001000
        BIT7 0——自动量程; 1——手动量程,
        BIT6 0——三相四线;1——三相三线,
        BIT5 0——功率;1——A相小电压信号,
        BIT3 1——PQ;BIT2 1——Q33;
        BIT1 1——Q90;
        BIT0 1——Q60;
        其中BIT0~BIT3只能有一位为1，并与BIT6一起使用
    :return:
    """
    try:
        int(way, 2)  # 使用int()函数，第二个参数2表示二进制
    except ValueError:
        logging.info(f"{way}不是一个有效的二进制字符串")
    set_cmd = [0x81, 0x01, 0x25, 0x0a, 0xa3, 0x00, 0x01, 0x20, int(bin_to_hex(way))]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()
    return ret


def set_ac(quc: float, qub: float, qua: float, qic: float, qib: float, qia: float, uc: float, ub: float, ua: float,
           ic: float, ib: float, ia: float, f: float):
    """
    设置AC，相位，幅值，频率值
    :param quc: C相电压相位
    :param qub: B相电压相位
    :param qua: A相电压相位
    :param qic: C相电流相位
    :param qib: B相电流相位
    :param qia: A相电流相位
    :param uc: C相电压
    :param ub: B相电压
    :param ua: A相电压
    :param ic: C相电流
    :param ib: B相电流
    :param ia: A相电流
    :param f: 频率
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x49, 0xa3, 0x05, 0x46, 0x3f]
    pdu = str(hex(int(quc * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)).replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(qub * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)).replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(qua * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)).replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(qic * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)).replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(qib * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)).replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(qia * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)).replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    set_cmd.append(0xFF)
    pdu = str(hex(int(uc * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'
    pdu = pdu.replace('0x', '').zfill(10)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
    set_cmd += pdu
    pdu = str(hex(int(ub * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'
    pdu = pdu.replace('0x', '').zfill(10)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
    set_cmd += pdu
    pdu = str(hex(int(ua * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'
    pdu = pdu.replace('0x', '').zfill(10)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
    set_cmd += pdu
    pdu = str(hex(int(ic * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'
    pdu = pdu.replace('0x', '').zfill(10)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
    set_cmd += pdu
    pdu = str(hex(int(ib * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'
    pdu = pdu.replace('0x', '').zfill(10)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
    set_cmd += pdu
    pdu = str(hex(int(ia * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'
    pdu = pdu.replace('0x', '').zfill(10)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
    set_cmd += pdu
    pdu = str(hex(int(f * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16))
    pdu = pdu.replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    set_cmd += [0x07, 0x07, 0x3F, 0x3F, 0x00]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(hex(xor).replace('0x', ''), 16))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()
    time.sleep(5)
    return ret


def set_gear_switching_mode(mode: str = '00000000'):
    """
    设置档位切换模式
    :param mode:
            控源档位模式：
            BIT7=0 自动档,
            BIT7=1 手动档
            手动档模式下，各通道档
            位更新标志：
            BIT0=1,Uc 更新档位
            BIT1=1,Ub 更新档位
            BIT2=1,Ua 更新档位
            BIT3=1,Ic 更新档位
            BIT4=1,Ib 更新档位
            BIT5=1,Ia 更新档位
    :return:
    """
    try:
        int(mode, 2)  # 使用int()函数，第二个参数2表示二进制
    except ValueError:
        logging.error(f"{mode}不是一个有效的二进制字符串")
    set_cmd = [0x81, 0x01, 0x25, 0x0a, 0xa3, 0x05, 0x40, 0x04, int(bin_to_hex(mode))]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()
    return ret


def set_harmonic_content(harmonic_content: list):
    """
    设置谐波含量
    :param harmonic_content:长度限制为21,每一个元素为谐波百分比值
    :return:
    """
    if len(harmonic_content) != 21:
        raise '谐波次数最大为21次，请确保为21个谐波'
    set_cmd = [0x81, 0x01, 0x07, 0x74, 0xa6, 0x05, 0x02, 0x00, 0x00, 0x69]
    for index, element in enumerate(harmonic_content):
        if index == 0:
            pdu = str(hex(int(element))).replace('0x', '').zfill(8)
            pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + '00'
            pdu = pdu.replace('0x', '').zfill(10)
            pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
            set_cmd += pdu
            continue
        pdu = str(hex(int(element))).replace('0x', '').zfill(8)
        pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fe'
        pdu = pdu.replace('0x', '').zfill(10)
        pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
        set_cmd += pdu

    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()
    return ret


def set_harmonic_phase(harmonic_phase: list):
    """
    设置谐波相位
    :param harmonic_phase:
    :return:
    """
    if len(harmonic_phase) != 21:
        raise '谐波次数最大为21次，请确保为21个谐波'
    set_cmd = [0x81, 0x01, 0x07, 0x5f, 0xa6, 0x05, 0x0a, 0x00, 0x00, 0x54]
    for index, element in enumerate(harmonic_phase):
        pdu = str(hex(int(element * 10000))).replace('0x', '').zfill(8)
        pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16))
        pdu = pdu.replace('0x', '').zfill(10)
        pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
        set_cmd += pdu
        logging.info(set_cmd)

    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()
    return ret


def set_harmonic_switch(uc_hc: str, ub_hc: str, ua_hc: str, ic_hc: str, ib_hc: str, ia_hc: str, total_switch: str):
    """
    设置谐波开关
    :param uc_hc:
        Uc 开关，每一 bit 为 1 代表当前次
        谐波开启
        Bit 0 = 基波(必须是 1)
        Bit 1 = 2 次谐波
        Bit 2 = 3 次谐波
        Bit 3 = 4 次谐波
        .
        .
        .
        Bit 20 = 21 次谐波
        开启 3 次谐波
        Bin:
        00000000000000000000000
        000000101
        转成 hex：0x00000005
        小端模式：05000000
    :param ub_hc:
    :param ua_hc:
    :param ic_hc:
    :param ib_hc:
    :param ia_hc:
    :param total_switch:
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x22, 0xa3, 0x05, 0x20, 0x7f]
    pdu = str(hex(int(bin_to_hex(uc_hc)))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16))
    pdu = pdu.replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(bin_to_hex(ub_hc)))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16))
    pdu = pdu.replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(bin_to_hex(ua_hc)))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16))
    pdu = pdu.replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(bin_to_hex(ic_hc)))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16))
    pdu = pdu.replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(bin_to_hex(ib_hc)))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16))
    pdu = pdu.replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    pdu = str(hex(int(bin_to_hex(ia_hc)))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16))
    pdu = pdu.replace('0x', '').zfill(8)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16)]
    set_cmd += pdu
    set_cmd.append(int(bin_to_hex(total_switch), 16))
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()
    return ret


def clear_overload_lock(overload_flag: str):
    """
    清除过载锁定
    :param overload_flag:
            BIT0=0,清除 UC
            BIT1=0,清除 UB
            BIT2=0,清除 UA
            BIT3=0,清除 IC
            BIT4=0,清除 IB
            BIT5=0,清除 IA
            其他 BIT 无效忽略
    :return:
    """
    try:
        int(overload_flag, 2)  # 使用int()函数，第二个参数2表示二进制
    except ValueError:
        logging.error(f"{overload_flag}不是一个有效的二进制字符串")
    set_cmd = [0x81, 0x01, 0x25, 0x0a, 0xa3, 0x00, 0x01, 0x80, int(bin_to_hex(overload_flag))]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()
    return ret


def switch_device_screen_interface(inter: int):
    """
    切换设备屏幕界面
    :param inter: 0x00  ARM版显示主界面;0x01  交流表界面;0x02  直流表界面;0x03 电能表误差检定界面
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x0a, 0xa3, 0x00, 0x10, 0x80, inter]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    source_control.send(pdu)
    source_control.close()


def change__underly_communicate(port):
    """
    切换底层通讯端口
    :param port:0-交流源;1-交流表;2-直流源;3-其他串口0;4-其他串口1
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x0a, 0xa3, 0x00, 0x10, 0x80, port]
    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    source_control.send(pdu)
    source_control.close()


def phase_amplitude_update():
    """
    相位更新 幅值更新
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x29, 0xa3, 0x05, 0x44, 0x3f, 0xe8, 0xcd, 0x08, 0x00, 0xfc, 0xe8, 0xcd, 0x08, 0x00,
               0xfc, 0xe8, 0xcd, 0x08, 0x00, 0xfc, 0x40, 0x4b, 0x4c, 0x00, 0xfa, 0x40, 0x4b, 0x4c, 0x00, 0xfa, 0x40,
               0x4b, 0x4c, 0x00, 0xfa, 0x02, 0x3f, 0x81]
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    source_control.send(pdu)
    source_control.close()


def frequency_renewal():
    """
    频率更新
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x0e, 0xa3, 0x05, 0x04, 0xc0, 0x20, 0xa1, 0x07, 0x00, 0x07, 0xc9]
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    source_control.send(pdu)
    source_control.close()


def harmonic_settings_and_switches():
    """
    谐波设置和谐波开关
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x23, 0xa3, 0x05, 0x42, 0x3f, 0x80, 0x4f, 0x12, 0x00, 0x00, 0x9f, 0x24, 0x00, 0x00,
               0x00, 0x00, 0x00, 0x80, 0x4f, 0x12, 0x00, 0x00, 0x9f, 0x24, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x3f,
               0xe2]
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    source_control.send(pdu)
    source_control.close()


def read_ac():
    """

    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x0d, 0xa0, 0x02, 0x3d, 0xff, 0x3f, 0xff, 0xff, 0x0f]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    logging.info(ret)
    logging.info(ret[1][0].hex())
    source_control.close()


def up_source_ac():
    """
    关源
    :return: 返回发送字节数
    """
    ret = set_ac(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    return ret


def update_voltage_gear(uc_gear, ub_gear, ua_gear):
    """
    下发电压档位
    :param uc_gear:uc电压档位
    :param ub_gear:ub电压档位
    :param ua_gear:ua电压档位
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x0c, 0xa3, 0x02, 0x02, 0x07, uc_gear, ub_gear, ua_gear]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()


def get_voltage_gear(voltage_value):
    """
    通过电压值，获取电压档位
    :param voltage_value: 电压值
    :return: 电压档位
    """
    voltage_gear = 0
    if voltage_value <= 30:
        voltage_gear = 5
    elif voltage_value <= 60:
        voltage_gear = 4
    elif voltage_value <= 120:
        voltage_gear = 3
    elif voltage_value <= 240:
        voltage_gear = 2
    elif voltage_value <= 480:
        voltage_gear = 1
    else:
        voltage_gear = 0
    return voltage_gear


def set_voltage_gear(uc_value, ub_value, ua_value):
    """
    设置电压档位
    :param uc_value:uc电压值
    :param ub_value:ub电压值
    :param ua_value:ua电压值
    :return:
    """
    uc_gear = get_voltage_gear(uc_value)
    ub_gear = get_voltage_gear(ub_value)
    ua_gear = get_voltage_gear(ua_value)
    update_voltage_gear(uc_gear, ub_gear, ua_gear)


def update_current_gear(ic_gear, ib_gear, ia_gear):
    """
    下发电流档位
    :param ic_gear:uc电流档位
    :param ib_gear:ub电流档位
    :param ia_gear:ua电流档位
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x0c, 0xa3, 0x02, 0x02, 0x38, ic_gear, ib_gear, ia_gear]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu)
    source_control.close()


def get_current_gear(current_value):
    """
    通过电流值，获取电流档位
    :param current_value: 电流值
    :return: 电流档位
    """
    current_gear = 0
    if current_value <= 0.01:
        current_gear = 12
    elif current_value <= 0.02:
        current_gear = 11
    elif current_value <= 0.05:
        current_gear = 10
    elif current_value <= 0.1:
        current_gear = 9
    elif current_value <= 0.2:
        current_gear = 8
    elif current_value <= 0.5:
        current_gear = 7
    elif current_value <= 1:
        current_gear = 6
    elif current_value <= 2:
        current_gear = 5
    elif current_value <= 5:
        current_gear = 4
    elif current_value <= 10:
        current_gear = 3
    elif current_value <= 20:
        current_gear = 2
    elif current_value <= 50:
        current_gear = 1
    else:
        current_gear = 0
    return current_gear


def set_current_gear(ic_value, ib_value, ia_value):
    """
    设置电流档位
    :param ic_value: ic电流值
    :param ib_value: ib电流值
    :param ia_value: ia电流值
    :return:
    """
    ic_gear = get_current_gear(ic_value)
    ib_gear = get_current_gear(ib_value)
    ia_gear = get_current_gear(ia_value)
    update_current_gear(ic_gear, ib_gear, ia_gear)


Log(str(__file__).split("\\")[-1])


def global_exception_handler(exc_type, exc_value, exc_traceback):
    """
    全局异常处理函数。适用于科陆源，当出现异常时关源

    全局异常钩子接收三个参数：
    :exc_type: 异常类 (如 ValueError, TypeError)
    :exc_value: 异常实例 (包含错误消息等)
    :exc_traceback: 跟踪对象 (包含调用栈信息)

    """
    # 记录错误信息
    logging.error("未捕获的异常:", exc_info=(exc_type, exc_value, exc_traceback))

    logging.info("系统将在5秒后关源...")
    sleep(5)

    # 执行关源命令
    ret = set_ac(120, 240, 0, 120, 240, 0, 0, 0, 0, 0, 0, 0, 50)
    time.sleep(5)
    switch_device_screen_interface(0x00)

    # 调用原始异常处理器
    sys.__excepthook__(exc_type, exc_value, exc_traceback)
