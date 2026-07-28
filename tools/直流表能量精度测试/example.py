# -*- coding: utf-8 -*-
"""
XL-9600 控源示例：走一遍完整校准流程。

运行：
    python example.py
按协议默认连接 192.168.1.105:24433，可在下方 IP/PORT 处修改。
"""

import time

from xl9600 import XL9600, SourceParams, OutputPoint, XL9600Error, XL9600Timeout

IP = "192.168.1.105"
PORT = 24433


def main() -> None:
    with XL9600(IP, PORT, timeout=5.0) as dev:
        # 1) 参数配置 ----------------------------------------------------
        print("[1] 参数配置 ...")
        dev.config(SourceParams(
            电流接入方式="间接接入式",
            供电方式="电源供电",
            额定电压="100V",
            标定电流="100A",
            分流器额定="75mV",
            被检表阻抗="1000Ω",
            脉冲常数=1000,        # imp/kwh
            校验圈数="自动",
            校验秒数=1,
        ))
        print("    配置完成")

        # 2) (可选) 开启误差自动上报 ------------------------------------
        # dev.set_error_report(True)

        # 3) 源输出 ------------------------------------------------------
        print("[2] 源输出 100% 点 ...")
        resp = dev.source_output(OutputPoint(
            电压检定点="100%",
            电流检定点="100%",
            电能方向="正向",
        ))
        print(f"    电压总值={resp.get('电压总值')} "
              f"电流总值={resp.get('电流总值')} "
              f"功率总值={resp.get('功率总值')}")

        # 源稳定一会儿再读误差
        time.sleep(2)

        # 4) 误差读取 ----------------------------------------------------
        print("[3] 误差读取（统计 5 次）...")
        result = dev.read_error(统计次数=5)
        print(f"    均值 = {result.均值}")
        print(f"    原始值 = {result.原始值}")

        # 5) 源停止 ------------------------------------------------------
        print("[4] 源停止 ...")
        dev.source_stop()

        # 6) 供电关闭 ----------------------------------------------------
        print("[5] 供电关闭 ...")
        dev.power_off()

        print("流程结束。")


if __name__ == "__main__":
    try:
        main()
    except XL9600Timeout as e:
        print(f"[超时] {e}")
    except XL9600Error as e:
        print(f"[设备错误] {e}")
    except OSError as e:
        print(f"[网络错误] {e}  —— 请检查 IP/端口与网络连通性")
