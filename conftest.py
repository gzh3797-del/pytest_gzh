import sys

import webdriver_manager.chrome
import yaml
import pytest
import os
import time
import logging
from _pytest.fixtures import FixtureRequest
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from api.modbus_connet import ModbusRtuOrTcp
from common.Source.CL3021.source_control import close_dc_all
from operation.WEB2.login import LoginPage

current_file_path = os.path.abspath(__file__)
project_root = os.path.dirname(current_file_path)
sys.path.append(project_root)
# 使用相对路径构建数据文件路径
data_file_path = os.path.join(project_root, 'datas', 'IOM')


# ================= 计时功能 Fixture ================= #
@pytest.fixture(scope="function")
def timer(request):
    """
    计时功能fixture：在函数执行前开始计时，执行完结束计时
    """
    # 获取当前测试函数名称
    test_name = request.node.name
    start_time = time.time()
    logging.info(f"测试函数 [{test_name}] 开始执行")
    # 执行测试函数
    yield
    end_time = time.time()
    execution_time = end_time - start_time
    close_dc_all()
    logging.info("关闭所有直流源输出！")
    logging.info(f"测试函数 [{test_name}] 执行完成，耗时: {execution_time:.4f} 秒")


# ================= 异常处理 Fixture ================= #
@pytest.fixture(scope="function")
def test_error_handler(request):
    """
    统一处理测试函数中的异常
    包括：KeyboardInterrupt、Exception 等
    执行完成后进行必要的清理工作
    """
    try:
        # 执行测试函数
        yield
    except KeyboardInterrupt:
        logging.info("程序被用户中断，关闭电源输出")
    except Exception as e:
        logging.error("程序异常终止，关闭电源输出")
        logging.error(str(e))
    finally:
        # 无论测试成功或失败，都执行清理工作
        logging.info("程序执行完毕，关闭电源输出")


# ================= 统一连接对象 Fixture ================= #
@pytest.fixture(scope="function")
def modbus_client():
    """
    提供一个全局唯一的 Modbus 连接实例（会话级别）
    """
    client = None
    try:
        client = ModbusRtuOrTcp()
        yield client
    except Exception as e:
        logging.error(f"Modbus 客户端初始化失败: {str(e)}")
        raise
    finally:
        if client:
            try:
                client.close()
                time.sleep(0.5)  # 延长等待时间确保Windows释放串口资源
                logging.info("Modbus 客户端已关闭")
            except Exception as e:
                logging.error(f"关闭 Modbus 客户端时出错: {str(e)}")


# ================= 自动打开浏览器 Fixture ================= #
@pytest.fixture(scope="function")
def web2_driver():
    configured_options = Options()
    # 🚨 关键一步：忽略证书错误（成功配置）
    configured_options.add_argument('--ignore-certificate-errors')
    print("\n启动 Chrome 浏览器...")
    # 使用 webdriver-manager 自动获取驱动
    service = Service(webdriver_manager.chrome.ChromeDriverManager().install())
    # 2. 启动驱动，并传入 配置好的 configured_options
    driver = webdriver.Chrome(service=service, options=configured_options)

    web2_driver = LoginPage(driver)
    web2_driver.driver.maximize_window()
    # 3. 统一登录，确保每个测试用例都从登录状态开始
    web2_driver.driver.get("https://192.168.2.176/#/login")
    web2_driver.login("admin", "admin")

    yield web2_driver  # 测试用例执行期间，driver 对象保持存活

    # Teardown: 关闭浏览器
    time.sleep(1)
    print("\n关闭浏览器...")
    web2_driver.driver.quit()


# ================= 数据驱动 Fixture ================= #
@pytest.fixture(scope="module")
def yaml_data(request):
    """
    根据测试模块的文件名自动加载对应的 YAML 数据文件
    :param request:
    :return:
    """
    # 获取当前测试模块的文件路径
    module_path = request.module.__file__
    # 提取测试文件名
    module_name = os.path.splitext(os.path.basename(module_path))[0]
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
    根据测试函数名自动匹配 YAML 文件中的测试数据
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


# ================= pytest_generate_tests钩子自动识别案例数 ================= #
def pytest_generate_tests(metafunc):
    """
    自动识别测试数据中的案例数并生成参数化测试
    统一处理所有测试用例格式
    """
    # 检查测试函数是否接受index参数
    if "index" in metafunc.fixturenames:
        # 获取测试函数名称
        func_name = metafunc.function.__name__
        try:
            # 获取当前测试模块
            module = metafunc.module
            # 尝试直接加载yaml_data（复用现有的逻辑）
            module_path = module.__file__
            module_name = os.path.splitext(os.path.basename(module_path))[0]
            yaml_file = f"{module_name}.yaml"
            file_path = os.path.join(data_file_path, yaml_file)
            
            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8") as f:
                    yaml_data_content = yaml.safe_load(f)
                    
                    # 尝试获取测试函数对应的数据
                    raw_test_name = func_name.split('[')[0]  # 去掉可能的参数化标记
                    possible_keys = [
                        func_name,                 # 完整名称
                        raw_test_name,             # 基础名称
                        raw_test_name.replace('test_', '', 1)  # 去掉 test_ 前缀
                    ]
                    # 查找匹配的测试数据
                    test_data_content = None
                    for key in possible_keys:
                        if key in yaml_data_content:
                            test_data_content = yaml_data_content[key]
                            break
                    # 如果找到测试数据
                    if test_data_content:
                        # 确定最大案例数
                        max_cases = 0
                        # 处理字典类型数据
                        if isinstance(test_data_content, dict):
                            # 检查voltage或current列表的长度，适应不同的测试用例函数
                            if 'voltage' in test_data_content and isinstance(test_data_content['voltage'], list):
                                max_cases = len(test_data_content['voltage'])
                            elif 'current' in test_data_content and isinstance(test_data_content['current'], list):
                                max_cases = len(test_data_content['current'])
                        # 处理列表类型数据
                        elif isinstance(test_data_content, list):
                            max_cases = len(test_data_content)
                        
                        # 如果有案例，生成索引参数
                        if max_cases > 0:
                            indices = list(range(max_cases))
                            metafunc.parametrize("index", indices, ids=[f"case_{i}" for i in indices])
        except Exception as e:
            # 如果发生异常，记录但不中断测试
            logging.warning(f"自动参数化测试时出错: {str(e)}")