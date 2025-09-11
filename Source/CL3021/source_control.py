import socket
import struct
import time
import logging
from Common.modbus_config import modbus_config
from decimal import Decimal, ROUND_HALF_UP


class SourceControlError(Exception):
    def __init__(self, msg):
        self.msg = msg


class SourCon:
    def __init__(self):
        self.udp_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)  # udp方式socket配置：ipv4，udp
        self.udp_socket.settimeout(2)  # 2s超时
        self.udp_socket.bind((modbus_config['local']['ip'], modbus_config['local']['port']))  # 绑定本地地址
        self.dest_addr = (modbus_config['source']['ip'], modbus_config['source']['port'])  # 目标设备地址

    def send(self, hex_data, wait_response=True):
        send_data = hex_data.encode('gbk')  # 将中文字符通过gbk编码转化为字节流发送
        logging.info(send_data.decode('gbk'))  # 通过gbk解码并记录日志
        ret = self.udp_socket.sendto(send_data, self.dest_addr)  # 发送，不需要管理连接（connect等），返回发送的字节数
        if wait_response:
            try:
                self.udp_socket.recvfrom(1024)
            except TimeoutError:
                pass
        return ret

    def recv(self):
        try:
            recv_data = self.udp_socket.recvfrom(1024)  # 接受设备响应，最大1024字节
        except TimeoutError:
            raise SourceControlError(
                'Source control timeout. Check whether the software of the control source is turned on.')
        return recv_data[0].decode('gbk')  # 返回解码后的GBK字符串（中文）的第一段

    def close(self):
        self.udp_socket.close()


def sour_para_conf(input_method='直接'):
    """
    配置电源基本参数
    :param input_method:
    :return:
    """
    data = ''
    if input_method == '直接':
        data = '''<参数配置>
        电流接入方式:直接接入式;
        供电方式:电源供电;
        额定电压:1000;
        标定电流:500;
        分流器额定:18mV;
        被检表阻抗:0.0000277Ω;
        脉冲常数:21600;
        校验圈数:自动;
        校验秒数:1;
        <End>'''
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
        logging.error('电流输入方式错误，请重新配置')
    re = SourCon()
    re.send(data, wait_response=False)  # 发送数据
    re.close()


def sour_output(voltage: float, current: float, stable_time=10):
    """
    控制电压电流输出
    :param voltage:
    :param current:
    :param stable_time:
    :return:
    """
    vol = voltage / 1000 * 100  # 输出电压，转化为百分比，适应量程变化
    cur = current / 500 * 100  # 输出电流，转化为百分比，适应量程变化
    data = f'''<源输出>
    电压检定点:{vol}%;
    电流检定点:{cur}%;
    电压纹波比例:0%;
    电流纹波比例:0%;
    电压纹波相位:0度;
    电流纹波相位:0度;
    纹波频率:300Hz;
    电能方向:正向;
    <End>'''
    re = SourCon()
    re.send(data, wait_response=False)
    re.close()
    time.sleep(stable_time)  # 输出稳定等待时间10s


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
    re.send(data, wait_response=False)
    receive_ret = re.recv()  # 获取响应内容（中文应答）
    re.close()
    expect_receive = '<源输出应答>'
    if expect_receive not in receive_ret:
        raise SourceControlError('Source control fail,Please check Environment.')
    logging.info('Source control success, voltage is:{}, current is:{}'.format(voltage, current))
    time.sleep(stable_time)


def sour_stop():
    data = '''<源停止>
    <End>'''
    re = SourCon()
    re.send(data, wait_response=False)
    re.close()


class Cl3021SourCon:
    def __init__(self):
        self.udp_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)  # ipv4，udp
        self.udp_socket.settimeout(3)  # 3s超时
        self.udp_socket.bind((modbus_config['local']['ip'], modbus_config['local']['port']))
        self.dest_addr = (modbus_config['source']['ip'], modbus_config['source']['port'])

    def send(self, hex_data, wait_response=True):
        ret = self.udp_socket.sendto(hex_data, self.dest_addr)  # 返回发送的字节数
        if wait_response:
            recv_data = self.udp_socket.recvfrom(1024)
            return ret, recv_data
        return ret, None

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
    """
    将二进制字节流转化为16进制，（tcp或udp都是将数据转化为二进制字节流）
    :param binary:
    :return:
    """
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
    """
    计算异或校验和
    :param numbers:
    :return:
    """
    result = 0
    for number in numbers:
        result ^= number
    return result


def online():
    """
    设备上线
    :return:
    """
    online_cmd = [0x81, 0x01, 0x25, 0x06, 0xc9, 0xeb]
    pdu = bytearray(online_cmd)
    source_control = Cl3021SourCon()
    source_control.send(pdu, wait_response=False)
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
        print(f"{way}不是一个有效的二进制字符串")
    a = int(bin_to_hex(way), 16)
    set_cmd = [0x81, 0x01, 0x25, 0x0a, 0xa3, 0x00, 0x01, 0x20, a]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)  # 转化为2进制字节
    source_control = Cl3021SourCon()
    bytes_sent, _ = source_control.send(pdu, wait_response=False)
    source_control.close()
    return bytes_sent


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

    #  value = int(f * 10000)
    #  pdu = list(value.to_bytes(4, 'little'))
    # 放大10000倍，浮点转整数，整数转8位16进制，补0到8位
    pdu = str(hex(int(quc * 10000))).replace('0x', '').zfill(8)
    # 大端序改为小端序，可使用value.to_bytes(4, 'little') ？
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)).replace('0x', '').zfill(8)
    # 将8位16进制转化为4个10进制数字列表
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

    pdu = str(hex(int(uc * 10000))).replace('0x', '').zfill(8)  # 电压幅值转化为16进制
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'  # 转为小端序
    pdu = pdu.replace('0x', '').zfill(10)  # 去掉0x，补足10位
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]  # 转化为5个数的列表
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
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]  # 转化为5个数的列表
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
    set_cmd.append(int(hex(xor).replace('0x', ''), 16))  # 添加校验码
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu, wait_response=False)
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
    ret = source_control.send(pdu, wait_response=False)
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
    ret = source_control.send(pdu, wait_response=False)
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
        print(set_cmd)

    xor = xor_sum(set_cmd[1:-1])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu, wait_response=False)
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
    ret = source_control.send(pdu, wait_response=False)
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
    ret = source_control.send(pdu, wait_response=False)
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
    source_control.send(pdu, wait_response=False)
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
    source_control.send(pdu, wait_response=False)
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
    source_control.send(pdu, wait_response=False)
    source_control.close()


def frequency_renewal():
    """
    频率更新
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x0e, 0xa3, 0x05, 0x04, 0xc0, 0x20, 0xa1, 0x07, 0x00, 0x07, 0xc9]
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    source_control.send(pdu, wait_response=False)
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
    source_control.send(pdu, wait_response=False)
    source_control.close()


def read_ac():
    """
    读取交流参数
    :return:
    """
    set_cmd = [0x81, 0x01, 0x25, 0x0d, 0xa0, 0x02, 0x3d, 0xff, 0x3f, 0xff, 0xff, 0x0f]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu, wait_response=False)
    print(ret)
    print(ret[1][0].hex())
    source_control.close()


def voltage_gear_update(gear):
    """
    更新电压档位
    :param gear: 值 0：600V 档位，值 1：480V 档位，值 2：240V 档位，值 3：120V 档位，值 4：60V 档位，值 5：30V 档位，
    :return:
    """
    # gear=hex(gear)
    # print(gear)
    set_cmd = [0x81, 0x01, 0x25, 0x0c, 0xa3, 0x02, 0x02, 0x07, gear, gear, gear]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu, wait_response=False)
    source_control.close()


def current_gear_update(gear):
    """
     更新电流档位
    :param gear:
    0：100A 档位，
    1：50A 档位，
    2：20A 档位，
    3：10A 档位，
    4：5A 档位，
    5：2A 档位，
    6：1A 档位，
    7：0.5A 档位，
    8：0.2A 档位，
    9：0.1A 档位，
    10：0.05A 档位，
    11：0.02A 档位，
    12：0.01A 档位，
    :return:
    """
    # gear=hex(gear)
    # print(gear)
    set_cmd = [0x81, 0x01, 0x25, 0x0c, 0xa3, 0x02, 0x02, 0x38, gear, gear, gear]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu, wait_response=False)
    source_control.close()


def set_current_gear(gear):
    """
    根据电流值自动选择合适档位
    :param gear:
    :return:
    """
    if gear <= 0.01:
        current_gear_update(12)
    elif 0.01 < gear <= 0.02:
        current_gear_update(11)
    elif 0.02 < gear <= 0.05:
        current_gear_update(10)
    elif 0.05 < gear <= 0.1:
        current_gear_update(9)
    elif 0.1 < gear <= 0.2:
        current_gear_update(8)
    elif 0.2 < gear <= 0.5:
        current_gear_update(7)
    elif 0.5 < gear <= 1:
        current_gear_update(6)
    elif 1 < gear <= 2:
        current_gear_update(5)
    elif 2 < gear <= 5:
        current_gear_update(4)
    elif 5 < gear <= 10:
        current_gear_update(3)
    elif 10 < gear <= 20:
        current_gear_update(2)
    elif 20 < gear <= 50:
        current_gear_update(1)
    else:
        current_gear_update(0)


def set_voltage_gear(gear):
    """
    根据电压值自动选择合适档位
    :param gear:
    :return:
    """

    if gear <= 30:
        voltage_gear_update(5)
    elif 30 < gear <= 60:
        voltage_gear_update(4)
    elif 60 < gear <= 120:
        voltage_gear_update(3)
    elif 120 < gear <= 240:
        voltage_gear_update(2)
    elif 240 < gear <= 480:
        voltage_gear_update(1)
    else:
        voltage_gear_update(0)


def set_dc(u: float, i: float):
    """
    设置直流源输出
    :param u: 单位V
    :param i: 单位mA
    :return:
    """
    set_cmd = [0x81, 0x01, 0x26, 0x11, 0x31, 0x03]
    pdu = str(hex(int(u * 10000))).replace('0x', '').zfill(8)  # 电压幅值转化为16进制
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'  # 转为小端序
    pdu = pdu.replace('0x', '').zfill(10)  # 去掉0x，补足10位
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]  # 转化为5个数的列表
    set_cmd += pdu
    i = i/1000
    pdu = str(hex(int(i * 10000))).replace('0x', '').zfill(8)
    pdu = hex(int(pdu[6:8] + pdu[4:6] + pdu[2:4] + pdu[0:2], 16)) + 'fc'
    pdu = pdu.replace('0x', '').zfill(10)
    pdu = [int(pdu[0:2], 16), int(pdu[2:4], 16), int(pdu[4:6], 16), int(pdu[6:8], 16), int(pdu[8:10], 16)]
    set_cmd += pdu
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(hex(xor).replace('0x', ''), 16))  # 添加校验码
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu, wait_response=False)
    source_control.close()


def bytes_to_float(hex_list, scale=1e6):
    integer_value = int.from_bytes(hex_list, byteorder='little', signed=False)
    return integer_value/scale


def read_dc(u_or_ma=3):
    """
    读取直流测量值
    :return:
    """
    set_cmd = [0x81, 0x01, 0x26, 0x06, 0xA3]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu, wait_response=True)
    source_control.close()
    # 分析返回值
    byte_data = ret[1][0]
    # 字节转16进制整数列表
    measurement_data = list(byte_data)
    # [129, 38, 1, 32, 83, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 254, 255, 255, 255, 250, 191, 255, 255, 255, 250, 0, 0, 0, 0, 0, 21]
    u_data = bytes_to_float(measurement_data[16:20])
    i_data = bytes_to_float(measurement_data[21:25])
    if u_or_ma == 0:
        return u_data
    elif u_or_ma == 1:
        return i_data
    else:
        return u_data, i_data


def set_dc_read_mode():
    """
    配置直流表测量模式
        0x00,同时测量电压和电流
        0x01,测量直流电压
        0x02,测量直流电流`
    :return:
    """
    set_cmd = [0x81, 0x01, 0x26, 0x07, 0x3C, 0x00]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    print(set_cmd)
    pdu = bytearray(set_cmd)
    source_control = Cl3021SourCon()
    ret = source_control.send(pdu, wait_response=False)
    source_control.close()


def close_dc(gear):
    """
    关闭输出，电流或电压
    :param gear: 1：电压，2：电流
    :return:
    """
    source_control = Cl3021SourCon()
    set_cmd = [0x81, 0x01, 0x26, 0x07, 0x38, gear]
    xor = xor_sum(set_cmd[1:])
    set_cmd.append(int(xor))
    pdu = bytearray(set_cmd)
    source_control.send(pdu, wait_response=False)
    source_control.close()


def close_dc_all():
    """
    关闭直流源输出
    """
    source_control = Cl3021SourCon()
    set_cmd_clear_overload = [0x81, 0x01, 0x26, 0x07, 0x39, 0x00, 0x19]
    set_cmd_u_close = [0x81, 0x01, 0x26, 0x07, 0x38, 0x01, 0x19]
    set_cmd_i_close = [0x81, 0x01, 0x26, 0x07, 0x38, 0x02, 0x19]
    set_cmd_1 = [0x81, 0x01, 0x25, 0x0a, 0xa3, 0x05, 0x01, 0x40, 0x00, 0xc9]
    set_cmd_2 = [0x81, 0x01, 0x25, 0x0a, 0xa3, 0x00, 0x10, 0x80, 0x00, 0x1d]
    pdu = bytearray(set_cmd_clear_overload)
    source_control.send(pdu, wait_response=False)
    time.sleep(0.5)
    pdu = bytearray(set_cmd_u_close)
    source_control.send(pdu, wait_response=False)
    time.sleep(0.5)
    pdu = bytearray(set_cmd_i_close)
    source_control.send(pdu, wait_response=False)
    time.sleep(0.5)
    pdu = bytearray(set_cmd_1)
    source_control.send(pdu, wait_response=False)
    time.sleep(0.5)
    pdu = bytearray(set_cmd_2)
    source_control.send(pdu, wait_response=False)
    source_control.close()


if __name__ == "__main__":

    close_dc_all()
    # set_dc(0, 12)

    # set_dc(0, 12)


