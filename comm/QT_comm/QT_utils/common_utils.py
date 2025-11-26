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
        with ModbusClient(ModbusProtocol.RTU) as rtu_client:
            rtu_client.validate_register_value('Reset for update', value='¡XX¡', use_6a=True)
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
            if current_log_num < value:
                construct_num = value - current_log_num
                logging.info(f'需要写入的交易日志数量{construct_num}')
                for i in range(construct_num):
                    res1 = client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')
                    time.sleep(0.02)
                    while res1[21:23] != '10':
                        res1 = client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')
                        time.sleep(0.02)
                    res2 = client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 45')
                    time.sleep(0.02)
                    while res2[21:23] != '10':
                        res2 = client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')
                        time.sleep(0.02)
                    if i % 50 == 0:
                        self.helper.click_pos((1269, 500))
            elif current_log_num > value:
                if clear:
                    self.clear_transaction_logs()
                    logging.info(f'目前交易数量大于测试数量，需要清空交易数量，重新写入')
                    for i in range(value):
                        res1 = client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')
                        time.sleep(0.02)
                        while res1[21:23] != '10':
                            res1 = client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')
                            time.sleep(0.02)
                        res2 = client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 45')
                        time.sleep(0.02)
                        while res2[21:23] != '10':
                            res2 = client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')
                            time.sleep(0.02)
                        if i % 50 == 0:
                            self.helper.click_pos((1269, 500))
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

    def read_transaction_logs(self):
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
        time_value, type_value, old_value, new_value = self.helper.copy_echilog_info((660, 367), (889, 368),
                                                                                     (1184, 369), (1586, 367))
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
