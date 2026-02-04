import serial
import time
import logging
import datetime
import pandas as pd
import struct
import os
import sys

# =================================================###########
#   通过串口定时读取电表的电流电压读数并记录当前时间戳保存到excl表
#                      用户配置区 (在此修改参数)
# =================================================###########
CONFIG = {
    "port": "COM3",               # 串口号 (例如 "COM3" )
    "baudrate": 9600,             # 波特率
    "read_interval": 2,           # 读取间隔 (秒)
    "output_file": "采集数据.xlsx", # 保存的 Excel 文件名
    # 电压指令: 01 03 30 00 00 02 CB 0B
    "voltage_cmd": "01 03 30 00 00 02 CB 0B",
    # 电流指令: 01 03 30 02 00 02 6A CB
    "current_cmd": "01 03 30 02 00 02 6A CB"
}
# =================================================###########

# --- 路径处理逻辑：确保文件保存在 EXE 同级目录 ---
if getattr(sys, 'frozen', False):
    # 如果是打包后的 EXE 运行
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # 如果是普通 Python 脚本运行
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 设置日志和输出文件路径
LOG_FILE = os.path.join(BASE_DIR, f"log_{datetime.datetime.now().strftime('%Y%m%d')}.txt")
OUTPUT_FILE = os.path.join(BASE_DIR, CONFIG["output_file"])

# 设置日志格式
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# 转换十六进制指令为字节流
VOLT_CMD = bytes.fromhex(CONFIG['voltage_cmd'])
CURR_CMD = bytes.fromhex(CONFIG['current_cmd'])

data_list = []

def decode_modbus_float(raw_bytes):
    """
    解析 Modbus 返回的 4 字节数据（通常为 32位浮点数）。
    """
    # 标准响应长度：地址(1)+功能(1)+字节数(1)+数据(4)+CRC(2) = 9 字节
    if len(raw_bytes) < 9:
        return None

    # 提取第 3 到第 6 字节（Modbus 数据区）
    data_part = raw_bytes[3:7]
    try:
        # '>f' 代表大端模式浮点数。如果数值不对，可能需要尝试其他字节序
        val = struct.unpack('>f', data_part)[0]
        return round(val, 3)
    except Exception as e:
        logging.error(f"解析数据失败: {e}")
        return None

def save_to_excel():
    """将采集到的数据保存到 Excel"""
    global data_list
    if not data_list:
        return
    try:
        df = pd.DataFrame(data_list)
        df.to_excel(OUTPUT_FILE, index=False)
        logging.info(f"数据已保存至 Excel (共 {len(data_list)} 条)")
    except Exception as e:
        logging.error(f"Excel 保存失败: {e}")
        print(f"!!! Excel 写入失败 (请检查文件是否被 Excel 打开): {e}")

def main():
    global data_list
    print(f"========================================")
    print(f" 串口号: {CONFIG['port']}")
    print(f" 波特率: {CONFIG['baudrate']}")
    print(f" 间  隔: {CONFIG['read_interval']} 秒")
    print(f" 文  件: {OUTPUT_FILE}")
    print(f"========================================")
    print(f"提示: 按 Ctrl+C 可停止并保存数据\n")

    try:
        with serial.Serial(CONFIG['port'], CONFIG['baudrate'], timeout=1) as ser:
            logging.info("串口已打开，开始采集...")

            while True:
                now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                # 1. 发送电压读取指令
                ser.write(VOLT_CMD)
                time.sleep(0.1) # 稍作停顿等待响应
                res_v = ser.read(ser.in_waiting or 9)
                v_val = decode_modbus_float(res_v)

                # 2. 发送电流读取指令
                time.sleep(0.1)
                ser.write(CURR_CMD)
                time.sleep(0.1)
                res_i = ser.read(ser.in_waiting or 9)
                i_val = decode_modbus_float(res_i)

                if v_val is not None and i_val is not None:
                    record = {
                        "时间戳": now_str,
                        "电压(V)": v_val,
                        "电流(A)": i_val
                    }
                    data_list.append(record)
                    print(f"[{now_str}] 电压: {v_val:>7} V | 电流: {i_val:>7} A")

                    # 每 10 条记录自动执行一次磁盘写入，防止断电丢失
                    if len(data_list) % 10 == 0:
                        save_to_excel()
                else:
                    msg = "读取失败 (设备响应超时或数据格式错误)"
                    print(f"[{now_str}] {msg}")
                    logging.warning(msg)

                time.sleep(CONFIG['read_interval'])

    except KeyboardInterrupt:
        print("\n\n>>> 检测到手动停止，正在执行最终保存...")
    except serial.SerialException as e:
        print(f"\n!!! 串口错误: {e}")
        logging.critical(f"串口错误: {e}")
    except Exception as e:
        print(f"\n!!! 发生意外错误: {e}")
        logging.critical(f"崩溃信息: {e}")
    finally:
        save_to_excel()
        print(f"完成！数据已保存到: {OUTPUT_FILE}")
        print("程序将在 3 秒后退出...")
        time.sleep(3)

if __name__ == "__main__":
    main()