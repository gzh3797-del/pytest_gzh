# -*- coding: utf-8 -*-
"""
config.py — AcuHMI-1-7 项目统一配置

所有需要因环境而调整的参数集中在此，各子模块从这里导入，不再各自散落定义。

子模块导入示例：
    from test_case.AcuHMI_1_7.config import HMI_IP, HMI_USERNAME, HMI_PASSWORD
    from test_case.AcuHMI_1_7 import config as cfg; cfg.METER_TCP_IP
"""

# ══════════════════════════════════════════════════════════════════════════════
# HMI 网关 — Web UI 连接
# ══════════════════════════════════════════════════════════════════════════════
HMI_IP       = "192.168.2.8"
HMI_URL      = f"https://{HMI_IP}"   # Playwright goto 用
HMI_USERNAME = "q"
HMI_PASSWORD = "1"

# HMI 页面 Device 下拉中显示的被测 4100 设备名（如 'Acu4100'）；留空则选 All
HMI_DEVICE_NAME = "Acu4100"

# ══════════════════════════════════════════════════════════════════════════════
# BACnet/IP 客户端（BAC0）
# ══════════════════════════════════════════════════════════════════════════════
BACNET_PORT  = 47808        # HMI BACnet/IP 服务端口（ASHRAE 默认值）

LOCAL_IP     = "192.168.2.45"   # 本机 BAC0 监听地址（与 WEB2 测试机相同）
LOCAL_PORT   = 47810            # 本机 BAC0 监听端口（不与 HMI 服务端口冲突）

BACNET_CONNECT_WAIT  = 4.0    # s，BAC0 启动后等待网关响应
BACNET_READ_TIMEOUT  = 20.0   # s，单次 BACnet 属性读取超时
BACNET_RESTART_WAIT  = 6.0    # s，UI 保存配置后等待 BACnet 服务重启

# ══════════════════════════════════════════════════════════════════════════════
# 接线检查 — AcuRev4100 Modbus TCP
# ══════════════════════════════════════════════════════════════════════════════
METER_TCP_IP   = "192.168.2.29"   # 4100 电表 IP
METER_TCP_PORT = 502
MODBUS_SLAVE   = 102

# ══════════════════════════════════════════════════════════════════════════════
# 接线检查 — 测试参数
# ══════════════════════════════════════════════════════════════════════════════
NOMINAL_VOLTAGE = 100    # V，写入设备的额定电压
NORMAL_CURRENT  = 1.0    # A，电流侧测试时的正常电流幅值
CHECK_TIMEOUT   = 30     # s，等待检查完成的超时时间
