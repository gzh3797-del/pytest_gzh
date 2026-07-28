"""列出可用 COM 串口与 USB(VISA) 资源，供填写 config.yaml 使用。

双击 `列出设备.bat` 运行；把打印出来的 COM 口 / USB 资源填进 config.yaml。
"""


def main():
    print("=== 可用 COM 串口 ===")
    try:
        from serial.tools import list_ports

        ports = list(list_ports.comports())
        if ports:
            for p in ports:
                print(f"  {p.device}  ({p.description})")
        else:
            print("  (未发现 COM 口)")
    except Exception as e:  # noqa: BLE001
        print("  读取 COM 口失败（需安装 pyserial）:", e)

    print()
    print("=== 可用 USB (VISA) 资源 ===")
    try:
        from src.transport import list_usb_resources

        res = list_usb_resources()
        if res:
            for r in res:
                print(f"  {r}")
        else:
            print("  (未发现 USB 设备；确认已装 NI-VISA 且频率计已用 USB 连接)")
    except Exception as e:  # noqa: BLE001
        print("  读取 USB 资源失败（需安装 NI-VISA + pyvisa）:", e)

    print()
    print("提示：源/电表的串口填对应 COM 口；频率计 USB 可在 config.yaml 用 resource: auto 自动识别。")
    input("按回车键退出…")


if __name__ == "__main__":
    main()
