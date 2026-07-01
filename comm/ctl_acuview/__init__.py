"""Acuview 2 上位机自动化控制框架。

模块:
    config        - 配置加载
    spec_loader   - Excel 寄存器表 + JSON 控件模型 → 统一模型
    modbus_client - Modbus TCP/RTU 直读直写(底层能力 + 校验真值源)
    uia_probe     - 探测上位机控件树，决定 GUI 后端
    gui_driver    - 驱动 Acuview 2 GUI(模拟测试人员)
    verify        - 跨传输闭环断言 + 报告
"""

__version__ = "0.1.0"
