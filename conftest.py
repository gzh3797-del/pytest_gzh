import atexit
import sys
import yaml
import pytest
import json
import os
from typing import Dict, Union
import logging
from _pytest.fixtures import FixtureRequest
from pymodbus.client import ModbusSerialClient
from Config.IOM.modbus_connet import ModbusRtuOrTcp
from Config.IOM.modbus_set_attr import set_all_ai_top_bot
from Source.CL3021.source_control import close_dc_all

data_file_path = r"C:\Users\ZihanGao\PycharmProjects\pythonProject\Datas\IOM"

# @pytest.fixture(scope="session", autouse=True)
# def first_step():
#     logging.info("开始执行前置步骤======》关闭直流电源")
#     close_dc_all()


# ================= Modbus Client Fixture ================= #
@pytest.fixture(scope="function")
def modbus_client():
    """
    提供一个全局唯一的 Modbus 连接实例（会话级别）
    """
    import time
    client = None
    try:
        client = ModbusRtuOrTcp(conn_mode="rtu")  # 也可以 "tcp"，根据你的配置文件
        logging.info("Modbus 客户端初始化完成")
        
        if not client.client.is_connected():
            raise ConnectionError("Modbus 客户端连接失败")
        
        yield client
    finally:
        if client:
            client.close()
            time.sleep(0.5)  # 延长等待时间确保Windows释放串口资源
            logging.info("Modbus 客户端已关闭")


# ================= 数据驱动 Fixture ================= #
@pytest.fixture(scope="session")  # 每个模块加载一次 YAML
def yaml_data(request):
    # 获取当前测试模块的文件路径
    module_path = request.module.__file__
    # 提取模块名（不带扩展名）
    module_name = os.path.splitext(os.path.basename(module_path))[0]
    # 构建 YAML 文件路径
    yaml_file = f"{module_name}.yaml"
    file_path = os.path.join(data_file_path, yaml_file)

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            try:
                return yaml.safe_load(f)
            except yaml.YAMLError as e:
                pytest.fail(f"YAML 格式错误: {file_path}\n{str(e)}")
    except FileNotFoundError:
        pytest.skip(f"YAML 文件未找到: {file_path}")


@pytest.fixture()
def test_data(request: FixtureRequest, yaml_data):
    """
    自动匹配测试函数名的数据
    逻辑：
    1. 去除参数化后缀（如 [1]）
    2. 优先匹配完整测试名，再匹配基础函数名
    3. 支持去掉 test_ 前缀
    """
    raw_test_name = request.node.name.split('[')[0]  # 去掉参数化标记
    possible_keys = [
        request.node.name,                 # 完整名称
        raw_test_name,                     # 基础名称
        raw_test_name.replace('test_', '', 1)  # 去掉 test_ 前缀
    ]

    for key in possible_keys:
        if key in yaml_data:
            return yaml_data[key]

    pytest.skip(f"测试数据未找到，尝试的键名: {possible_keys}")
# def pytest_configure(config):
#     """配置 pytest 日志"""
#     log_dir = "logs"
#     if not os.path.exists(log_dir):
#         os.makedirs(log_dir)
#
#     logging.basicConfig(
#         level=logging.INFO,
#         format="%(asctime)s [%(levelname)s] %(message)s",
#         handlers=[
#             logging.FileHandler(f"{log_dir}/test.log"),  # 输出到文件
#             logging.StreamHandler()  # 输出到控制台
#         ]
#     )


# def pytest_configure(config):
#     # 清除所有现有的日志处理器
#     for handler in logging.root.handlers[:]:
#         logging.root.removeHandler(handler)
#         handler.close()
#
#     # 然后重新配置日志
#     log_dir = "logs"
#     log_filename = "test.log"
#     log_file = os.path.join(log_dir, log_filename)
#
#     os.makedirs(log_dir, exist_ok=True)
#     file_handler = logging.FileHandler(log_file, encoding='GBK')
#
#     logging.basicConfig(
#         level=logging.INFO,
#         format="%(asctime)s [%(levelname)s] %(message)s",
#         handlers=[file_handler, logging.StreamHandler()]
#     )