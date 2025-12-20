import serial
import time

try:
    # 尝试以最简方式打开
    print("正在尝试打开 COM39...")
    ser = serial.Serial('COM39', 19200, timeout=1)

    if ser.is_open:
        print("成功打开串口！")
        print(f"当前设置: {ser.name}, {ser.baudrate}")
        # 做一点简单的操作
        time.sleep(0.5)
        ser.close()
        print("串口已正常关闭。")

except serial.SerialException as e:
    print(f"依然报错: {e}")
    print("建议执行：拔掉USB线，等待5秒重新插入，再次运行。")
except Exception as e:
    print(f"其他错误: {e}")