import json
import os

root_path = r"C:\Users\ZihanGao\PycharmProjects\pythonProject\Datas"


def read_json():
    with open(root_path + '/config.json', "r", encoding="utf-8") as f:
        ret = json.load(f)
    return ret


modbus_config = read_json()


def write_json(key, value):
    if key == 'baudrate':
        modbus_config['rtu'][key] = value
    if key == 'parity':
        modbus_config['rtu'][key] = value
    if key == 'slaveid':
        modbus_config['rtu'][key] = value
    if key == 'ip':
        modbus_config['tcp'][key] = value
    if key == 'port':
        modbus_config['tcp'][key] = value
    with open(root_path + '/config.json', "w", encoding="utf-8") as f:
        json.dump(modbus_config, f, indent=4, ensure_ascii=False)
