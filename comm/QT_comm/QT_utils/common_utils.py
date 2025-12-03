from datetime import datetime

import pyautogui
import pyperclip
import time
import logging
import pytest
from comm.QT_comm.QT_utils.ModbusClient import ModbusProtocol, ModbusClient


class CommonUtils:
    """测试工具类，包含所有配置方法"""

    def __init__(self, helper, app_path, device_image_path):
        self.helper = helper
        self.app_path = app_path
        self.device_image_path = device_image_path

    def connect_device(self):
        """连接设备"""
        self.helper.connect_device(self.device_image_path)

    def configure_identification_status(self, value):
        """
        配置Identification Status信息

        Args:
            value (bool): 'True'或'False'，表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 配置Identification Status信息
        # IS位置: (595, 422)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IS', offset_x=258)

        # 根据传入的值选择不同的坐标
        if value:
            # 选择True的坐标
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IS',
                                    offset_x=169, offset_y=37)
        else:
            # 选择False的坐标
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IS',
                                    offset_x=169, offset_y=21)

        # 更新配置
        self._update_configuration()

    def configure_identification_level(self, value):
        """
        配置Identification Level信息

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 配置Identification Status信息
        # II位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IL', offset_x=340)

        # 根据传入的值选择不同的坐标
        dic = {
            'NONE': 20,
            'HEARSAY': 35,
            'TRUSTED': 50,
            'VERIFIED': 67,
            'CERTIFIED': 97,
            'SECURE': 112,
            'MISMATCH': 127,
            'INVALID': 142,
            'OUTDATED': 157,
            'UNKNOWN': 170
        }

        # 选择值的坐标
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IL',
                                offset_x=187, offset_y=dic[value])

        # 更新配置
        self._update_configuration()

    def configure_identification_flag1(self, value):
        """
        配置Identification flag1信息

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 配置Identification Status信息
        # IF位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IF1', offset_x=340)

        # 根据传入的值选择不同的坐标
        dic = {
            'RFID_NONE': 20,
            'RFID_PLAIN': 35,
            'RFID_RELATED': 50,
            'RFID_PSK': 67
        }

        # 选择值的坐标
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IF1',
                                offset_x=187, offset_y=dic[value])

        # 更新配置
        self._update_configuration()

    def configure_identification_flag2(self, value):
        """
        配置Identification flag2信息

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 配置Identification Status信息
        # IF位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IF1', offset_x=340, index=1)

        # 根据传入的值选择不同的坐标
        dic = {
            'OCPP_NONE': 20,
            'OCPP_RS': 35,
            'OCPP_AUTH': 50,
            'OCPP_RS_TLS': 67,
            'OCPP_AUTH_TLS': 97,
            'OCPP_CACHE': 112,
            'OCPP_WHITELIST': 127,
            'OCPP_CERTIFIED': 142,
        }

        # 选择值的坐标
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IF1',
                                offset_x=187, offset_y=dic[value], index=1)

        # 更新配置
        self._update_configuration()

    def configure_identification_flag3(self, value):
        """
        配置Identification flag2信息

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 配置Identification Status信息
        # IF位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IF1', offset_x=340, index=2)

        # 根据传入的值选择不同的坐标
        dic = {
            'ISO15118_NONE': 20,
            'ISO15118_PNC': 35,
        }

        # 选择值的坐标
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IF1',
                                offset_x=187, offset_y=dic[value], index=2)

        # 更新配置
        self._update_configuration()

    def configure_identification_flag4(self, value):
        """
        配置Identification flag2信息

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 配置Identification Status信息
        # IF位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IF1', offset_x=340, index=3)

        # 根据传入的值选择不同的坐标
        dic = {
            'PLMN_NONE': 20,
            'PLMN_RING': 35,
            'PLMN_SMS': 50,
        }

        # 选择值的坐标
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IF1',
                                offset_x=187, offset_y=dic[value], index=3)

        # 更新配置
        self._update_configuration()

    def configure_identification_type(self, value):
        """
        配置Identification type

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 配置Identification Status信息
        # IF位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IT', offset_x=365)

        # 根据传入的值选择不同的坐标
        dic = {
            'NONE': 20,
            'DENIED': 35,
            'UNDEFINED': 50,
            'ISO14443': 67,
            'ISO15693': 97,
            'EMAID': 112,
            'EVCCID': 127,
            'EVCOID': 142,
            'ISO7812': 157,
            'CARD_TXN_NR': 170,
            'CENTRAL': 50,
            'CENTRAL_1': 67,
            'CENTRAL_2': 97,
            'LOCAL': 112,
            'LOCAL_1': 127,
            'LOCAL_2': 142,
            'PHONE_NUMBER': 157,
            'KEY_CODE': 170
        }

        keys = list(dic.keys())
        if keys.index(value) > 9:
            pyautogui.moveTo(1352, 502)
            time.sleep(0.5)
            pyautogui.scroll(-500)  # 向下滚动
            time.sleep(1)
        else:
            pyautogui.moveTo(1352, 502)
            time.sleep(0.5)
            pyautogui.scroll(500)  # 向下滚动
            time.sleep(1)

        # 选择值的坐标
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\IT',
                                offset_x=221, offset_y=dic[value])

        # 更新配置
        self._update_configuration()

    def configure_identification_data(self, value):
        """
        配置Identification data

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 配置Identification Status信息
        # IF位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\ID', offset_x=340)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(value)

        # 更新配置
        self._update_configuration()

    def configure_Tariff_Text(self, value):
        """
        配置Tariff_Text

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 配置Identification Status信息
        # IF位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\TT', offset_x=340)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(value)

        # 更新配置
        self._update_configuration()

    def configure_CI(self, value):
        """
        配置CI

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Transaction')

        # 配置Identification Status信息
        # IF位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Transaction\CI', offset_x=200)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(value)

        # 更新配置
        self._update_configuration()

    def perform_charging_cycle(self, wait=2):
        """执行完整的充电循环"""
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging'):
            # 点击Setting
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

            # 点击交易
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 点击Start_Charging
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        if self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\operate_failed'):
            pytest.fail('充电操作执行失败')
        # 记录开始充电时间
        actual_start_time = time.time()
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        time.sleep(wait)
        # 点击End_Charging
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\End_Charging')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')

        # 记录结束充电时间
        actual_end_time = time.time()
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        return actual_start_time, actual_end_time

    def read_and_parse_transaction_log(self):
        """读取并解析交易日志"""
        # 点击Reading
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading')

        # 点击Transaction_Log
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log')

        # 点击Read_Log
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Read_Log')
        time.sleep(2)
        # 点击Stop
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Stop')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')

        # 点击第一行日志
        self.helper.double_click_pos((1079, 364))

        # 复制日志
        self.helper.hotkey('ctrl', 'a')
        time.sleep(2)
        self.helper.hotkey('ctrl', 'c')

        # 解析日志
        result = pyperclip.paste()
        logging.info(f"读取的交易日志: {result}")

        return result

    def _update_configuration(self):
        """通用的配置更新流程"""
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Update')
        if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Nothing_updata'):
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        else:
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Confirm'):
                self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Confirm')
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\update_failed'):
                pytest.fail('更新操作失败')
            else:
                if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Yes'):
                    self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
                    time.sleep(4)
                if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Yes'):
                    self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')

    def restart_application(self):
        """重启应用程序"""
        self.helper.kill_acuview_apps()
        self.helper.wait(2)
        self.helper.hotkey('win', 'd')
        self.helper.wait(1)
        self.helper.launch_app(self.app_path)

    def configure_time_zone_shift(self, value):
        """
            配置充电交易时区信息

            Args:
                value (int): '-1440~1440'，表示要配置的值单位min
            """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击transaction配置
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Transaction')

        # 配置time_zone_shift信息
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Transaction\time_zone_shift', offset_x=189)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(value)

        # 更新配置
        self._update_configuration()

    def configure_CT(self, value):
        """
        配置configure_CT信息

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击交易
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Transaction')

        # 配置Identification Status信息
        # IF位置: (594, 461)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Transaction\CT', offset_x=340)

        # 根据传入的值选择不同的坐标
        dic = {
            'EVSEID': 20,
            'CBIDC': 35,
        }

        # 选择值的坐标
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Transaction\CT',
                                offset_x=250, offset_y=dic[value])

        # 更新配置
        self._update_configuration()

    def configure_Transaction_Timeout(self, value):
        """
            配置Transaction_Timeout

            Args:
                value (int): 表示要配置的值单位s
            """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击transaction配置
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Transaction')

        # 配置time_zone_shift信息
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Transaction\Transaction_Timeout',
                                offset_x=291, offset_y=10)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(value)

    def configure_time_Sync_status(self):
        """
            配置time Sync status
            """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击 Test_Charging
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 点击 Set_Time
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\Set_Time')

        # 更新配置
        self._update_configuration()

    def reboot_device(self):
        """重启电表"""
        self.restart_application()
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.validate_register_value('Reset for update', '¡XX¡', True)
        time.sleep(15)
        # 重新连接设备
        self.helper.connect_device(self.device_image_path)
        time.sleep(2)

    def construct_transaction_logs(self, value, clear=True, connect=True):
        """
        构造交易日志

        Args:
            value (int): 表示构造的数量
        """
        with ModbusClient(ModbusProtocol.TCP) as client:
            current_log_num = client.parse_data(client.validate_register_value('max Transaction Id'))
            logging.info(f'目前交易日志数量{current_log_num}')

            def send_with_retry(message, max_retries=10):
                """发送消息并重试，最多重试max_retries次"""
                for attempt in range(max_retries):
                    res = client.send_custom_message(message)
                    time.sleep(0.03)
                    if res and res[21:23] == '10':
                        return res, True
                    logging.warning(f'第{attempt + 1}次发送失败，响应: {res}')
                return None, False

            def reset_device_and_retry(max_reset_retries=3):
                """重置设备并重试发送消息"""
                for reset_attempt in range(max_reset_retries):
                    try:
                        logging.info(f'第{reset_attempt + 1}次尝试重置设备...')

                        # 重置设备
                        client.validate_register_value('Serial Number', 'DE55061235', True)
                        client.validate_register_value('Reset for update', '¡XX¡', True)

                        # 等待设备重启
                        logging.info('等待设备重启...')
                        time.sleep(20)

                        # 重新连接设备
                        with ModbusClient(ModbusProtocol.TCP) as new_client:
                            # 重新获取当前日志数量
                            new_current_log_num = new_client.parse_data(
                                new_client.validate_register_value('max Transaction Id')
                            )
                            logging.info(f'设备重启后交易日志数量: {new_current_log_num}')

                            # 尝试再次发送消息
                            def retry_after_reset(message):
                                for attempt in range(5):  # 重置后减少重试次数
                                    res = new_client.send_custom_message(message)
                                    time.sleep(0.05)
                                    if res and res[21:23] == '10':
                                        return res, True
                                    logging.warning(f'重置后第{attempt + 1}次发送失败')
                                return None, False

                            return new_client, retry_after_reset

                    except Exception as e:
                        logging.error(f'设备重置第{reset_attempt + 1}次失败: {str(e)}')
                        if reset_attempt < max_reset_retries - 1:
                            time.sleep(5)  # 等待5秒后重试
                        continue

                return None, None

            if current_log_num < value:
                construct_num = value - current_log_num
                logging.info(f'需要写入的交易日志数量{construct_num}')

                for i in range(construct_num):
                    logging.info(f'正在写入第{i + 1}/{construct_num}条交易日志')

                    # 发送第一条消息
                    res1, success1 = send_with_retry('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')
                    if not success1:
                        logging.error('第一条消息发送失败，尝试重置设备...')
                        new_client, retry_func = reset_device_and_retry()
                        if new_client and retry_func:
                            client = new_client  # 更新客户端
                            res1, success1 = retry_func('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')

                        if not success1:
                            pytest.fail('写入交易日志失败，请检查电表状态')

                    # 发送第二条消息
                    res2, success2 = send_with_retry('00 01 00 00 00 09 01 10 52 00 00 01 02 00 45')
                    if not success2:
                        logging.error('第二条消息发送失败，尝试重置设备...')
                        new_client, retry_func = reset_device_and_retry()
                        if new_client and retry_func:
                            client = new_client  # 更新客户端
                            res2, success2 = retry_func('00 01 00 00 00 09 01 10 52 00 00 01 02 00 45')

                        if not success2:
                            pytest.fail('写入交易日志失败，请检查电表状态')

            elif current_log_num > value:
                if clear:
                    self.clear_transaction_logs()
                    logging.info(f'目前交易数量大于测试数量，需要清空交易数量，重新写入')

                    for i in range(value):
                        logging.info(f'正在写入第{i + 1}/{value}条交易日志')

                        # 发送第一条消息
                        res1, success1 = send_with_retry('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')
                        if not success1:
                            logging.error('第一条消息发送失败，尝试重置设备...')
                            new_client, retry_func = reset_device_and_retry()
                            if new_client and retry_func:
                                client = new_client  # 更新客户端
                                res1, success1 = retry_func('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')

                            if not success1:
                                pytest.fail('写入交易日志失败，请检查电表状态')

                        # 发送第二条消息
                        res2, success2 = send_with_retry('00 01 00 00 00 09 01 10 52 00 00 01 02 00 45')
                        if not success2:
                            logging.error('第二条消息发送失败，尝试重置设备...')
                            new_client, retry_func = reset_device_and_retry()
                            if new_client and retry_func:
                                client = new_client  # 更新客户端
                                res2, success2 = retry_func('00 01 00 00 00 09 01 10 52 00 00 01 02 00 45')

                            if not success2:
                                pytest.fail('写入交易日志失败，请检查电表状态')

            current_log_num2 = client.parse_data(client.validate_register_value('max Transaction Id'))

        if connect:
            self.connect_device()
            time.sleep(2)
            # 点击Reading
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading')
            # 点击Transaction_Log
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log')

        return current_log_num2

    def clear_transaction_logs(self):
        with ModbusClient(ModbusProtocol.TCP) as rtu_client:
            rtu_client.validate_register_value('Clear Transaction Log', value=1)

    def read_logs(self):
        """
        读取交易日志和ec日志通用步骤
        """
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Read_Log')
        while True:
            time.sleep(20)
            self.helper.click_pos((1353, 282))
            if self.helper.check_image_exists(
                    r'page_elements\Acuview_public\Reading_page\Transaction_Log\Read_Failed'):
                pytest.fail('读取日志失败')
            elif self.helper.check_image_exists(
                    r'page_elements\Acuview_public\Reading_page\Transaction_Log\Reading_Successful'):
                self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
                break

    def configure_Cable(self, value):
        """
        配置电缆损失补偿配置

        Args:
            value (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击General
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General')

        # 配置电缆损失补偿信息
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\cable', offset_x=180)
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(value)

        # 更新配置
        self._update_configuration()

    def configure_Cable_status(self, value):
        """
        配置电缆损失补偿状态

        Args:
            状态 (str): 表示要配置的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

        # 点击General
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General')

        # 配置电缆损失补偿状态
        if value == '0':
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\cable', offset_x=-66)

        if value == '1':
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\cable_disable', offset_x=-66)

        # 更新配置
        self._update_configuration()

    def read_echilog_info(self):
        """
        读取echilog最新日志内容

        Args:
            数量 (str): 表示要第几条的日志
        """
        # 点击Reading
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading')
        # 点击Echi_Log
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Echilog')
        # 点击Read_Log
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Read_Log')
        # 点击Stop
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Stop')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')

        # 点击第一行日志
        time_value, type_value, old_value, new_value = self.helper.copy_echilog_info((660, 367), (889, 367),
                                                                                     (1184, 367), (1586, 367))
        return time_value, type_value, old_value, new_value

    def start_charging(self):
        """开始充电操作"""
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging'):
            # 点击Setting
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

            # 点击交易
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 点击Start_Charging
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        popup_exists = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Setting_page\Test_Charging\Charging_Started_Popup')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        return popup_exists

    def end_charging(self):
        """结束充电操作"""
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging'):
            # 点击Setting
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

            # 点击交易
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 点击End_Charging
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\End_Charging')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        popup_exists = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Setting_page\Test_Charging\Charging_End_Popup')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        return popup_exists

    def abort_charging(self):
        """终止充电操作"""
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging'):
            # 点击Setting
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

            # 点击交易
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')

        # 点击End_Charging
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\Abort_Charging')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        popup_exists = self.helper.check_image_exists(
            r'page_elements\Acuview_public\Setting_page\Test_Charging\Charging_End_Popup')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        return popup_exists

    def long_time_charging(self, charging_time):
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging'):
            # 点击Setting
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)
            # 点击交易
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')
        # 点击Start_Charging
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        self.restart_application()
        time.sleep(charging_time - 13)
        # 点击End_Charging
        self.connect_device()
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\Start_Charging'):
            # 点击Setting
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)

            # 点击交易
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')
        # 点击End_Charging
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\End_Charging')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')

    def construct_EClig_logs(self, value, clear=True, connect=True):
        """
        构造交易日志

        Args:
            value (int): 表示构造的数量
        """
        with ModbusClient(ModbusProtocol.TCP) as client:
            current_log_num = client.parse_data(client.validate_register_value('Record Number'))
            logging.info(f'目前EClig_logs数量{current_log_num}')

            def send_with_retry(message, max_retries=10):
                """发送消息并重试，最多重试max_retries次"""
                for attempt in range(max_retries):
                    res = client.send_custom_message(message)
                    time.sleep(0.1)
                    if res and res[21:23] == '10':
                        return res, True
                    logging.warning(f'第{attempt + 1}次发送失败，响应: {res}')
                return None, False

            def reset_device_and_retry(max_reset_retries=3):
                """重置设备并重试发送消息"""
                for reset_attempt in range(max_reset_retries):
                    try:
                        logging.info(f'第{reset_attempt + 1}次尝试重置设备...')

                        # 重置设备
                        client.validate_register_value('Serial Number', 'DE55061235', True)
                        client.validate_register_value('Reset for update', '¡XX¡', True)

                        # 等待设备重启
                        logging.info('等待设备重启...')
                        time.sleep(20)

                        # 重新连接设备
                        with ModbusClient(ModbusProtocol.TCP) as new_client:
                            # 重新获取当前日志数量
                            new_current_log_num = new_client.parse_data(
                                new_client.validate_register_value('Record Number')
                            )
                            logging.info(f'设备重启后EClig_logs数量: {new_current_log_num}')

                            # 获取重启后的状态
                            new_get_loss_status = new_client.parse_data(
                                new_client.send_custom_message('00 01 00 00 00 09 01 03 10 22 00 01')
                            )

                            if new_get_loss_status == 1:
                                new_status_code = '00'
                                new_next_status_code = '01'
                            else:
                                new_status_code = '01'
                                new_next_status_code = '00'

                            logging.info(f'设备重启后状态码: status={new_status_code}, next={new_next_status_code}')

                            # 尝试再次发送消息
                            def retry_after_reset(message):
                                for attempt in range(5):  # 重置后减少重试次数
                                    res = new_client.send_custom_message(message)
                                    time.sleep(0.05)
                                    if res and res[21:23] == '10':
                                        return res, True
                                    logging.warning(f'重置后第{attempt + 1}次发送失败')
                                return None, False

                            return new_client, retry_after_reset, new_status_code, new_next_status_code

                    except Exception as e:
                        logging.error(f'设备重置第{reset_attempt + 1}次失败: {str(e)}')
                        if reset_attempt < max_reset_retries - 1:
                            time.sleep(5)  # 等待5秒后重试
                        continue

                return None, None, None, None

            get_loss_status = client.parse_data(client.send_custom_message('00 01 00 00 00 09 01 03 10 22 00 01'))
            if get_loss_status == 1:
                status_code = '00'
                next_status_code = '01'
            else:
                status_code = '01'
                next_status_code = '00'

            logging.info(f'初始状态码: status={status_code}, next={next_status_code}')

            if current_log_num < value:
                construct_num = value - current_log_num
                logging.info(f'需要写入的Record Number数量{construct_num}')

                for i in range(construct_num):
                    logging.info(f'正在写入第{i + 1}/{construct_num}条EClig日志')

                    if i % 2 == 0:
                        res1, success1 = send_with_retry(f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {status_code}')
                        if not success1:
                            logging.error('第一条消息发送失败，尝试重置设备...')
                            new_client, retry_func, new_status, new_next = reset_device_and_retry()
                            if new_client and retry_func:
                                client = new_client  # 更新客户端
                                status_code = new_status  # 更新状态码
                                next_status_code = new_next
                                res1, success1 = retry_func(f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {status_code}')

                            if not success1:
                                pytest.fail('写入EClig_logs失败，请检查电表状态')
                    else:
                        res2, success2 = send_with_retry(
                            f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {next_status_code}')
                        if not success2:
                            logging.error('第二条消息发送失败，尝试重置设备...')
                            new_client, retry_func, new_status, new_next = reset_device_and_retry()
                            if new_client and retry_func:
                                client = new_client  # 更新客户端
                                status_code = new_status  # 更新状态码
                                next_status_code = new_next
                                res2, success2 = retry_func(
                                    f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {next_status_code}')

                            if not success2:
                                pytest.fail('写入EClig_logs失败，请检查电表状态')

            elif current_log_num > value:
                if clear:
                    self.clear_transaction_logs()
                    logging.info(f'目前EClig_logs数量大于测试数量，需要清空交易数量，重新写入')

                    # 清空操作
                    clear_res, clear_success = send_with_retry('00 01 00 00 00 09 01 10 20 09 00 01 02 00 01')
                    if not clear_success:
                        logging.error('清空操作失败，尝试重置设备...')
                        new_client, retry_func, new_status, new_next = reset_device_and_retry()
                        if new_client and retry_func:
                            client = new_client
                            status_code = new_status
                            next_status_code = new_next
                            clear_res, clear_success = retry_func('00 01 00 00 00 09 01 10 20 09 00 01 02 00 01')

                        if not clear_success:
                            pytest.fail('清空EClig_logs失败，请检查电表状态')

                    for i in range(value):
                        logging.info(f'正在写入第{i + 1}/{value}条EClig日志')

                        if i % 2 == 0:
                            res1, success1 = send_with_retry(f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {status_code}')
                            if not success1:
                                logging.error('第一条消息发送失败，尝试重置设备...')
                                new_client, retry_func, new_status, new_next = reset_device_and_retry()
                                if new_client and retry_func:
                                    client = new_client
                                    status_code = new_status
                                    next_status_code = new_next
                                    res1, success1 = retry_func(
                                        f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {status_code}')

                                if not success1:
                                    pytest.fail('写入EClig_logs失败，请检查电表状态')
                        else:
                            res2, success2 = send_with_retry(
                                f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {next_status_code}')
                            if not success2:
                                logging.error('第二条消息发送失败，尝试重置设备...')
                                new_client, retry_func, new_status, new_next = reset_device_and_retry()
                                if new_client and retry_func:
                                    client = new_client
                                    status_code = new_status
                                    next_status_code = new_next
                                    res2, success2 = retry_func(
                                        f'00 01 00 00 00 09 01 10 10 22 00 01 02 00 {next_status_code}')

                                if not success2:
                                    pytest.fail('写入EClig_logs失败，请检查电表状态')

            current_log_num2 = client.parse_data(client.validate_register_value('Record Number'))
            logging.info(f'最终EClig_logs数量: {current_log_num2}')

        if connect:
            self.connect_device()
            time.sleep(2)
            # 点击Reading
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading')
            # 点击Transaction_Log
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Echilog')

        return current_log_num2

    def configure_Pulse_LED_Energy(self, parameter=None, constant=None):
        """配置LED能量脉冲"""
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General')
        # parameter(1303, 689)
        parameter_dict = {'None': 20,
                          'Import_Energy': 39,
                          'Export_Energy': 52,
                          'NET_Energy': 72,
                          'TOTAL_Energy': 89, }
        if parameter:
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\parameter', offset_x=211)
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\parameter', offset_x=127,
                                    offset_y=parameter_dict[parameter])
        if constant:
            # constant(1301, 731)
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\constant', offset_x=169)
            self.helper.hotkey('ctrl', 'a')
            self.helper.paste_text(constant)

        # 更新配置
        self._update_configuration()

    def read_echilog_multi_line(self, line_num):
        """
        读取echilog最新日志内容

        Args:
            数量 (str): 表示要第几条的日志
        """
        # 点击Reading
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading')
        # 点击Echi_Log
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Echilog')
        # 点击Read_Log
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Read_Log')
        # 点击Stop
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Stop')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        result = []
        y = 367
        # 点击第一行日志
        for i in range(line_num):
            result.append(self.helper.copy_echilog_info((660, y), (889, y), (1184, y), (1586, y)))
            y += 30
        return result

    def configure_cable_loss(self, status=None, resistance=None):
        """
        配置电缆损失补偿配置

        Args:
            status(str):'off' or 'on'
            resistance:具体补偿的值
        """
        # 点击Setting
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)
        # 点击General
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General')
        if status == 'off':
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\General\cable'):
                self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\cable', offset_x=-66)
        if status == 'on':
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\General\cable_disable'):
                self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\cable_disable',
                                        offset_x=-66)
        if resistance:
            # 配置电缆损失补偿信息
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\General\cable', offset_x=180)
            self.helper.hotkey('ctrl', 'a')
            self.helper.paste_text(resistance)

        # 更新配置
        self._update_configuration()

    def set_time(self, offset_second, way):
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\System_Status')
        if way == 'Reading':
            # Set_Time(1065, 299)
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\System_Status\Set_Time')
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Confirm'):
                self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Confirm')
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\update_failed'):
                pytest.fail('更新操作失败')
            else:
                if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Yes'):
                    self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\System_Status\Set_Time', offset_y=45)
            self.helper.double_click_image(r'page_elements\Acuview_public\Reading_page\System_Status\Set_Time',
                                         offset_x=-148, offset_y=235)
            new_time = self.helper.get_future_time_str(offset_second)
            self.helper.enter_text(new_time)
            old_time = self.helper.get_future_time_str(1)
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\System_Status\Set_Time')
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\update_failed'):
                pytest.fail('更新操作失败')
            else:
                if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Yes'):
                    self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            return self.time_format(old_time), self.time_format(new_time)
        if way == 'Setting':
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\System_Status\Set_Time', offset_y=45)
            self.helper.double_click_image(r'page_elements\Acuview_public\Reading_page\System_Status\Set_Time',
                                         offset_x=-148, offset_y=235)
            set_time = self.helper.get_future_time_str(offset_second)
            old_time = self.helper.get_future_time_str(offset_second+8)
            self.helper.enter_text(set_time)
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\System_Status\Set_Time')
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Confirm'):
                self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Confirm')
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\update_failed'):
                pytest.fail('更新操作失败')
            else:
                if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Yes'):
                    self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Setting', 1)
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging')
            new_time = self.helper.get_future_time_str(1)
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Test_Charging\Set_Time')
            if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\update_failed'):
                pytest.fail('更新操作失败')
            else:
                if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Yes'):
                    self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            return self.time_format(old_time), self.time_format(new_time)



    def time_format(self, time_str: str) -> str:
        """
        将 "YYYYMMDDHHMMSS" 格式转换为 "YYYY-MM-DD HH:MM:SS" 格式

        Args:
            time_str: 格式为 "20251203152809" 的时间字符串

        Returns:
            str: 格式为 "2025-12-03 15:28:09" 的时间字符串
        """
        # 解析紧凑格式的时间字符串
        dt = datetime.strptime(time_str, "%Y%m%d%H%M%S")

        # 格式化为易读格式
        return dt.strftime("%Y-%m-%d %H:%M:%S")

    def is_time_within_2_seconds(self,expected_str: str, actual_str: str,
                                 format_str: str = "%Y-%m-%d %H:%M:%S") -> bool:
        """
        判断两个时间字符串相差是否不超过2秒

        Args:
            expected_str: 期望时间字符串
            actual_str: 实际时间字符串
            format_str: 时间格式

        Returns:
            bool: 如果相差不超过2秒返回True，否则返回False
        """
        # 解析时间字符串为datetime对象
        expected_time = datetime.strptime(expected_str, format_str)
        actual_time = datetime.strptime(actual_str, format_str)

        # 计算时间差的绝对值（秒数）
        time_difference = abs((actual_time - expected_time).total_seconds())

        # 判断是否不超过2秒
        return time_difference <= 2

