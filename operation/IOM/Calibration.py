import time
import serial
import struct
from api.modbus_connet import ModbusRtuOrTcp
from config.modbus_config import modbus_config
from common.Source.CL3021.source_control import set_dc, close_dc_all, close_dc, read_dc
from operation.IOM.IOM_get_attr import get_single_ai_y_measurement
from operation.IOM.IOM_set_attr import float_to_uint32t_4bytes, set_all_ai_param, set_all_ao_param


# CRC16
def crc16(data):
    crc = 0xFFFF
    for ch in data:
        crc ^= ch
        for _ in range(8):
            if crc & 1:
                crc = (crc >> 1) ^ 0xA001
            else:
                crc >>= 1
    return struct.pack('<H', crc)


def float_to_hex_bytes(value):
    """
    将浮点数转换为IEEE 754单精度浮点数格式的十六进制字节表示
    例如: 2.000456 -> '40 00 07 79'
    :param value: 要转换的浮点数
    :return: 空格分隔的十六进制字节字符串
    """
    bytes_data = struct.pack('>f', value)
    # 每个字节转换为两位十六进制，并用空格分隔
    hex_str = ' '.join(f'{b:02X}' for b in bytes_data)
    return hex_str


def send_and_wait(ser, frame, expected_frame_code=None, timeout=40):
    """
    发送一帧请求并等待正确的响应
    :param ser: 已打开的串口对象
    :param frame: 要发送的bytes数据帧
    :param expected_frame_code: 期望返回的帧码，None表示不校验
    :param timeout: 超时时间（秒）
    :return: 返回正确响应（bytes），或 None 表示失败
    """
    start_time = time.time()
    frame += crc16(frame)
    while True:
        # 每次发送前清空接收缓冲区
        ser.reset_input_buffer()
        ser.write(frame)
        ser.flush()

        res = b""
        while True:
            if ser.in_waiting:
                res += ser.read(ser.in_waiting)
            if time.time() - start_time > timeout:
                print("等待响应超时！")
                return None
            if len(res) >= 7:
                if expected_frame_code is not None:
                    if res[4] != expected_frame_code:
                        time.sleep(2)
                        print("期望帧码:", expected_frame_code, "实际帧码:", res[4])
                        break
                    print("期望帧码:", expected_frame_code, "实际帧码:", res[4])
                return res


def ai_c_calibration(ser, model_type, slave_id=modbus_config['rtu']['slaveid']):
    """
    AI电流校准
    :param ser: 已打开的串口对象
    :param model_type: function model type，1或2
    :param slave_id:
    :return: None
    """
    hex_id = f'{slave_id:02X}'
    ai_numbers = 17 if model_type == 2 else 9
    f_ai_c_1 = bytes.fromhex(f'{hex_id} 6A 50 00 00 02 04 00 FA 01 01')
    f_ai_c_2 = bytes.fromhex(f'{hex_id} 6A 50 00 00 02 04 00 FC 01 01')
    f_ai_c_0 = bytes.fromhex(f'{hex_id} 03 50 02 00 01')

    numbers = [x for x in range(1, ai_numbers)]
    all_f_1 = []
    all_f_2 = []

    for n in numbers:
        frame = bytearray(f_ai_c_1)
        frame2 = bytearray(f_ai_c_2)
        frame[9] = n
        frame2[9] = n
        all_f_1.append(bytes(frame))
        all_f_2.append(bytes(frame2))

    time.sleep(3)
    print(f"*************开始校准AI1电流*************")
    for i in range(8, ai_numbers):
        print("DC源输出2mA")
        set_dc(0, 2)
        time.sleep(6.8)
        print(f"发送：{all_f_1[i - 1].hex(' ')}")
        send_and_wait(ser, all_f_1[i - 1])

        print(f"查询AI状态：{f_ai_c_0.hex(' ')}")
        send_and_wait(ser, f_ai_c_0, 0x01)

        print("DC源输出12mA")
        set_dc(0, 12)
        time.sleep(6.8)
        print(f"发送：{all_f_2[i - 1].hex(' ')}")
        send_and_wait(ser, all_f_2[i - 1])

        print(f"查询AI状态：{f_ai_c_0.hex(' ')}")
        send_and_wait(ser, f_ai_c_0, 0x03)

        close_dc(2)
        print(f"*************AI{i}电流校准完成，你有8s时间切换到AI{i + 1}*************")
        time.sleep(6)
    close_dc_all()


def ai_v_calibration(ser, model_type, slave_id=modbus_config['rtu']['slaveid']):
    """
    AI电压校准
    :param ser: 已打开的串口对象
    :param model_type: function model type，1或2
    :param slave_id:
    :return: None
    """
    hex_id = f'{slave_id:02X}'
    if model_type == 2:
        f_ai_v_1 = bytes.fromhex(f'{hex_id} 6A 50 00 00 02 04 00 AA 01 10')
        f_ai_v_2 = bytes.fromhex(f'{hex_id} 6A 50 00 00 02 04 00 BC 01 10')
    else:
        f_ai_v_1 = bytes.fromhex(f'{hex_id} 6A 50 00 00 02 04 00 AA 01 08')
        f_ai_v_2 = bytes.fromhex(f'{hex_id} 6A 50 00 00 02 04 00 BC 01 08')
    f_ai_v_0 = bytes.fromhex(f'{hex_id} 03 50 02 00 01')
    print(f"*************开始校准AI电压*************")
    print("DC源输出输出1V")
    set_dc(1, 0)
    time.sleep(7)
    print(f"发送：{f_ai_v_1.hex(' ')}")
    send_and_wait(ser, f_ai_v_1)
    print(f"查询AI状态：{f_ai_v_0.hex(' ')}")
    send_and_wait(ser, f_ai_v_0, 0x01)
    print("DC源输出输出6V")
    set_dc(6, 0)
    time.sleep(6)
    print(f"发送：{f_ai_v_2.hex(' ')}")
    send_and_wait(ser, f_ai_v_2)
    print(f"查询AI状态：{f_ai_v_0.hex(' ')}")
    send_and_wait(ser, f_ai_v_0, 0x03)
    print(f"*************AI电压校准完成*************")
    close_dc_all()


def ao_v_calibration(ser, model_type, slave_id=modbus_config['rtu']['slaveid']):
    """
    AO电压校准
    :param ser: 已打开的串口对象
    :param model_type: function model type，1或2
    :param slave_id:
    :return: None
    """
    hex_id = f'{slave_id:02X}'
    ao_numbers = 5 if model_type == 2 else 3
    if model_type == 2:
        f_ao_v_1 = bytes.fromhex(f'{hex_id} 6A 51 00 00 02 04 00 AA 01 04')
        f_ao_v_2 = bytes.fromhex(f'{hex_id} 6A 51 00 00 02 04 00 BC 01 04')
        f_ao_v_3 = bytes.fromhex(f'{hex_id} 6A 51 00 00 02 04 00 CC 01 04')
    elif model_type == 1:
        f_ao_v_1 = bytes.fromhex(f'{hex_id} 6A 51 00 00 02 04 00 AA 01 02')
        f_ao_v_2 = bytes.fromhex(f'{hex_id} 6A 51 00 00 02 04 00 BC 01 02')
        f_ao_v_3 = bytes.fromhex(f'{hex_id} 6A 51 00 00 02 04 00 CC 01 02')
    f_ao1_v_1 = bytes.fromhex(f'{hex_id} 6A 51 03 00 02 04')
    f_ao1_v_2 = bytes.fromhex(f'{hex_id} 6A 51 05 00 02 04')
    f_ao2_v_1 = bytes.fromhex(f'{hex_id} 6A 51 07 00 02 04')
    f_ao2_v_2 = bytes.fromhex(f'{hex_id} 6A 51 09 00 02 04')
    f_ao3_v_1 = bytes.fromhex(f'{hex_id} 6A 51 0B 00 02 04')
    f_ao3_v_2 = bytes.fromhex(f'{hex_id} 6A 51 0D 00 02 04')
    f_ao4_v_1 = bytes.fromhex(f'{hex_id} 6A 51 0F 00 02 04')
    f_ao4_v_2 = bytes.fromhex(f'{hex_id} 6A 51 11 00 02 04')
    f_ao_v_0 = bytes.fromhex(f'{hex_id} 03 51 02 00 01')
    list_f_ao = [f_ao1_v_1, f_ao2_v_1, f_ao3_v_1, f_ao4_v_1, f_ao1_v_2, f_ao2_v_2, f_ao3_v_2, f_ao4_v_2]

    print(f"**************************AO校准第一步**************************")
    send_and_wait(ser, f_ao_v_1)
    send_and_wait(ser, f_ao_v_0, 0x01)

    for n in range(1, ao_numbers):
        if n != 1:
            time.sleep(8)
        else:
            time.sleep(3)
        print(f"**************************开始记录AI{n}的第一个点**************************")
        value_ao_1 = read_dc(0)
        frame_ao = bytearray(list_f_ao[n - 1])
        list_f_ao[n - 1] = bytes(frame_ao) + bytes.fromhex(float_to_hex_bytes(value_ao_1))
        if n == ao_numbers-1:
            print(
                f"AO{n}电压记录完成{value_ao_1},写入为{float_to_hex_bytes(value_ao_1)}，记录结束")
        else:
            print(f"AO{n}电压记录完成{value_ao_1},写入为{float_to_hex_bytes(value_ao_1)}，你有8s时间切换到AO{n + 1}")

    print(f"**************************AO校准第二步**************************")
    send_and_wait(ser, f_ao_v_2)
    send_and_wait(ser, f_ao_v_0, 0x02)

    for n in range(1, ao_numbers):
        if n != 1:
            time.sleep(8)
        else:
            time.sleep(3)
        print(f"**************************开始记录AI{n}的第二个点**************************")
        value_ao_2 = read_dc(0)
        frame_ao = bytearray(list_f_ao[n - 1 + 4]) + bytes.fromhex(float_to_hex_bytes(value_ao_2))
        list_f_ao[n - 1 + 4] = bytes(frame_ao)
        if n== ao_numbers-1:
            print(
                f"AO{n}电压记录完成{value_ao_2}，写入为{float_to_hex_bytes(value_ao_2)},记录结束")
        else:
            print(f"AO{n}电压记录完成{value_ao_2}，写入为{float_to_hex_bytes(value_ao_2)}，你有7s时间切换到AO{n + 1}")

    print(f"**************************开始写入**************************")
    if model_type == 2:
        for n in range(8):
            send_and_wait(ser, list_f_ao[n])
    else:
        for n in range(2):
            send_and_wait(ser, list_f_ao[n])
        for m in range(2):
            send_and_wait(ser, list_f_ao[m + 4])

    send_and_wait(ser, f_ao_v_3)
    send_and_wait(ser, f_ao_v_0, 0x04)
    print(f"*************AO校准完成*************")


if __name__ == "__main__":

    client = ModbusRtuOrTcp()
    target_port = 'COM39'
    mode_type = 1
    print(f"进入校准程序,修改所有ai_type为电压档")
    if mode_type == 1:
        set_all_ao_param(client, [0,1], [10,0,1,1,2,2,3,3,4,4], [1,8])
    elif mode_type ==2:
        set_all_ao_param(client, [0,1], [10,0,1,1,2,2,3,3,4,4])
    ser = serial.Serial(target_port, 19200, timeout=1)
    # print("开始校准AI电流")
    # ai_c_calibration(ser, mode_type=mode_type)
    # print("开始校准AI电压")
    # ai_v_calibration(ser, slave_id=slave_id, mode_type=mode_type)
    print("开始校准AO电压")
    ao_v_calibration(ser, model_type=mode_type)

