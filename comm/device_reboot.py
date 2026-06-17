import logging
import time
from tools.log import Log
import serial
from modbus_config import modbus_config
from comm.modbus_get_attr import read_voltage_measurement, read_current_measurement
from comm.source_control import sour_output, sour_stop

Log(str(__file__).split("\\")[-1])


def pow_on_device():
    ser = serial.Serial(port=modbus_config['reboot']['port'], baudrate=modbus_config['reboot']['baudrate'],
                        parity=modbus_config['reboot']['parity'], timeout=2)
    data = [0xA0, 0x01, 0x01, 0xA2]
    pdu = bytearray(data)
    ser.write(pdu)
    ser.close()


def pow_off_device():
    ser = serial.Serial(port=modbus_config['reboot']['port'], baudrate=modbus_config['reboot']['baudrate'],
                        parity=modbus_config['reboot']['parity'], timeout=2)
    data = [0xA0, 0x01, 0x00, 0xA1]
    pdu = bytearray(data)
    ser.write(pdu)
    ser.close()
    time.sleep(5)


def device_start_fail():
    # 上电后，单板未启动
    interval = 30
    i = 0
    while True:
        i += 1
        print(i)
        start_time = time.time()
        pow_on_device()
        time.sleep(interval - 10 - 10)
        vol = read_voltage_measurement()
        cur = read_current_measurement()
        logging.info('vol ret is:{}'.format(vol))
        logging.info('cur ret is:{}'.format(cur))
        print('vol ret is:{}'.format(vol))
        print('cur ret is:{}'.format(cur))
        if vol != 0 or cur != 0:
            logging.info('vol ret is:{}'.format(vol))
            logging.info('cur ret is:{}'.format(cur))
            print('vol ret is:{}'.format(vol))
            print('cur ret is:{}'.format(cur))
            break
        pow_off_device()
        print(time.time() - start_time)


def source_stop_measure_vol_cur():
    # 停止源输出时，电压电流仍旧存在
    interval = 30
    while True:
        start_time = time.time()
        sour_output(voltage=1000, current=130)
        time.sleep(interval - 10 - 12)
        sour_stop()
        vol = read_voltage_measurement()
        cur = read_current_measurement()
        logging.info('vol ret is:{}'.format(vol))
        logging.info('cur ret is:{}'.format(cur))
        print('vol ret is:{}'.format(vol))
        print('cur ret is:{}'.format(cur))
        if vol != 0 or cur > 0.5:
            logging.info('vol ret is:{}'.format(vol))
            logging.info('cur ret is:{}'.format(cur))
            print('vol ret is:{}'.format(vol))
            print('cur ret is:{}'.format(cur))
            break
        print(time.time() - start_time)



