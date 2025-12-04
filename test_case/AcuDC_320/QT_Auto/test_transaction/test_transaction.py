import time
import pyperclip
import pytest
import sys
import os
from comm.QT_comm.QT_utils.ModbusClient import ModbusProtocol, ModbusClient
from comm.QT_comm.QT_utils.QT_auto_utils import AutoHelper
from comm.QT_comm.QT_utils.common_utils import CommonUtils
from modbus_config import modbus_config

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class TestTransaction:
    """交易相关测试用例"""

    @pytest.fixture(autouse=True)
    def setup(self, request):
        """每个测试用例前的准备工作"""
        self.app_path = modbus_config["QT_path"]
        self.app_root_path = os.path.dirname(self.app_path)
        self.device_image_path = modbus_config["device_image_path"]
        self.helper = AutoHelper(confidence=0.8)
        self.test_name = request.node.name
        # 初始化工具类
        self.utils = CommonUtils(self.helper, self.app_path, self.device_image_path)
        self.helper.kill_acuview_apps()
        self.helper.hotkey('win', 'd')
        self.helper.launch_app(self.app_path)
        yield
        self.helper.kill_acuview_apps()

    def test_Function_AcuDC320_Sprint2_003_01_case1(self):
        """上位机可配置identification_status信息 - 参数化版本"""
        with ModbusClient(ModbusProtocol.TCP) as client:
            res1 = client.send_custom_message(
                '00 01 00 00 00 09 01 6A F0 6E 00 21 42 04 1D F8 39 E6 22 DB 8D 3F BB 8E F8 E8 27 C1 A3 79 BF 93 AE AB EE B9 AA 32 16 65 DD E1 15 B9 6B A0 04 74 A0 8F 7F 4F 2D CF 6B E6 DD C1 1D 1D 5A F9 79 D1 BA 51 56 1E DE A2 1A A8 EF E8 D2 44 A6 79 00 40 32')
            res2 = client.send_custom_message(
                '00 01 00 00 00 09 01 6A F0 8F 00 10 20 ED AE 6B 7E 1F CD 1D 51 59 FA EF 60 F7 6B 8F 9B 1D D5 70 55 A3 80 83 6B AE 71 21 F3 EE FA 84 50 A9 4E')
            res3 = client.send_custom_message(
                '00 01 00 00 00 09 01 03 F0 6E 00 21')
            res4 = client.send_custom_message(
                '00 01 00 00 00 09 01 03 F0 8F 00 10')
        pytest.assume(res1[21:23] == '6A', f"公钥写入失败，预期6A，实际{res1[21:23]}")
        pytest.assume(res2[21:23] == '6A', f"私钥写入失败，预期6A，实际{res2[21:23]}")
        pytest.assume(res3[21:23] == '03', f"公钥读取失败，预期03，实际{res3[21:23]}")
        pytest.assume(res4[21:23] != '03', f"私钥读取成功检查失败，预期不是03，实际{res4[21:23]}")

    @pytest.mark.parametrize("config_value,expected_value", [
        (True, True),
        (False, False)
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case3(self, config_value, expected_value):
        """上位机可配置identification_status信息 - 参数化版本"""
        self.utils.construct_transaction_logs(30, clear=False, connect=False)
        self.utils.connect_device()
        self.utils.configure_identification_status(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['IS']
        assert actual_result == expected_value, f"IS配置失败: 期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        ('NONE', 'NONE'),
        ('HEARSAY', 'HEARSAY'),
        ('TRUSTED', 'TRUSTED'),
        ('VERIFIED', 'VERIFIED'),
        ('CERTIFIED', 'CERTIFIED'),
        ('SECURE', 'SECURE'),
        ('MISMATCH', 'MISMATCH'),
        ('INVALID', 'INVALID'),
        ('OUTDATED', 'OUTDATED'),
        ('UNKNOWN', 'UNKNOWN')
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case4(self, config_value, expected_value):
        """上位机可配置Identification Level信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_identification_level(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['IL']
        assert actual_result == expected_value, f"IL配置失败: 期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        ('RFID_NONE', 'RFID_NONE'),
        ('RFID_PLAIN', 'RFID_PLAIN'),
        ('RFID_RELATED', 'RFID_RELATED'),
        ('RFID_PSK', 'RFID_PSK')
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case5(self, config_value, expected_value):
        """上位机可配置Identification Flag1信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_identification_flag1(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['IF'][0]
        assert actual_result == expected_value, f"IF1配置失败: 期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        ('OCPP_NONE', 'OCPP_NONE'),
        ('OCPP_RS', 'OCPP_RS'),
        ('OCPP_AUTH', 'OCPP_AUTH'),
        ('OCPP_RS_TLS', 'OCPP_RS_TLS'),
        ('OCPP_AUTH_TLS', 'OCPP_AUTH_TLS'),
        ('OCPP_CACHE', 'OCPP_CACHE'),
        ('OCPP_WHITELIST', 'OCPP_WHITELIST'),
        ('OCPP_CERTIFIED', 'OCPP_CERTIFIED'),
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case6(self, config_value, expected_value):
        """上位机可配置Identification Flag2信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_identification_flag2(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['IF'][1]
        assert actual_result == expected_value, f"IF2配置失败: 期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        ('ISO15118_NONE', 'ISO15118_NONE'),
        ('ISO15118_PNC', 'ISO15118_PNC')
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case7(self, config_value, expected_value):
        """上位机可配置Identification Flag3信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_identification_flag3(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['IF'][2]
        assert actual_result == expected_value, f"IF3配置失败: 期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        ('PLMN_NONE', 'PLMN_NONE'),
        ('PLMN_RING', 'PLMN_RING'),
        ('PLMN_SMS', 'PLMN_SMS')
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case8(self, config_value, expected_value):
        """上位机可配置Identification Flag4信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_identification_flag4(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['IF'][3]
        assert actual_result == expected_value, f"IF4配置失败: 期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        ('NONE', 'NONE'),
        ('DENIED', 'DENIED'),
        ('UNDEFINED', 'UNDEFINED'),
        ('ISO14443', 'ISO14443'),
        ('ISO15693', 'ISO15693'),
        ('EMAID', 'EMAID'),
        ('EVCCID', 'EVCCID'),
        ('EVCOID', 'EVCOID'),
        ('ISO7812', 'ISO7812'),
        ('CARD_TXN_NR', 'CARD_TXN_NR'),
        ('CENTRAL', 'CENTRAL'),
        ('CENTRAL_1', 'CENTRAL_1'),
        ('CENTRAL_2', 'CENTRAL_2'),
        ('LOCAL', 'LOCAL'),
        ('LOCAL_1', 'LOCAL_1'),
        ('LOCAL_2', 'LOCAL_2'),
        ('PHONE_NUMBER', 'PHONE_NUMBER'),
        ('KEY_CODE', 'KEY_CODE')
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case9(self, config_value, expected_value):
        """上位机可配置Identification Type信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_identification_type(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['IT']
        assert actual_result == expected_value, f"IT配置失败: 期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        ('aaaabbbbbbbbbccccccccccc', 'aaaabbbbbbbbbccccccccccc'),
        ('111111111122222222333333', '111111111122222222333333'),
        ('@@@@!!!!!%%%%%***^&%%%%$', '@@@@!!!!!%%%%%***^&%%%%$'),
        ('1122222222222222cccafsQWQEREQR$%^@$#@#$#@$@$%EGFHHGJ~ADV',
         '1122222222222222cccafsQWQEREQR$%^@$#@#$#@$@$%EGFHHGJ~ADV')
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case10(self, config_value, expected_value):
        """上位机可配置Identification Data信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_identification_data(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['ID']

        if config_value != '1122222222222222cccafsQWQEREQR$%^@$#@#$#@$@$%EGFHHGJ~ADV':
            assert actual_result == expected_value, f"ID配置失败: 期望{expected_value}, 实际{actual_result}"
        else:
            assert actual_result != expected_value, f"ID配置失败: 不期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        ('aaaabbbbbbbbbccccccccccc', 'aaaabbbbbbbbbccccccccccc'),
        ('111111111122222222333333', '111111111122222222333333'),
        ('@@@@!!!!!%%%%%***^&%%%%$', '@@@@!!!!!%%%%%***^&%%%%$'),
        ('1122222222222222cccafsQWQEREQR$%^@$#@#$#@$@$%EGFHHGJ~ADV',
         '1122222222222222cccafsQWQEREQR$%^@$#@#$#@$@$%EGFHHGJ~ADV')
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case11(self, config_value, expected_value):
        """上位机可配置Tariff Text信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_Tariff_Text(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['TT']

        if config_value != '1122222222222222cccafsQWQEREQR$%^@$#@#$#@$@$%EGFHHGJ~ADV':
            assert actual_result == expected_value, f"TT配置失败: 期望{expected_value}, 实际{actual_result}"
        else:
            assert actual_result != expected_value, f"TT配置失败: 不期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        (1, 1), (1800, 1800), (3599, 3599)
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case14_1(self, config_value, expected_value):
        """上位机可配置CT信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_Transaction_Timeout(value=config_value)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Update')
        if self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\Confirm'):
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Confirm')
        if not self.helper.check_image_exists(r'page_elements\Acuview_public\Setting_page\invalid_value'):
            assert False, f'Transaction_Timeout非法值{config_value}设置成功'

    @pytest.mark.parametrize("config_value,expected_value", [
        ('EVSEID', 'EVSEID'),
        ('CBIDC', 'CBIDC')
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case15(self, config_value, expected_value):
        """上位机可配置CT信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_CT(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['CT']
        assert actual_result == expected_value, f"CT配置失败: 期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        ('aaaabbbbbbbbbccccccccccc', 'aaaabbbbbbbbbccccccccccc'),
        ('111111111122222222333333', '111111111122222222333333'),
        ('@@@@!!!!!%%%%%***^&%%%%$', '@@@@!!!!!%%%%%***^&%%%%$'),
        ('1122222222222222cccafsQWQEREQR$%^@$#@#$#@$@$%EGFHHGJ~ADV',
         '1122222222222222cccafsQWQEREQR$%^@$#@#$#@$@$%EGFHHGJ~ADV')
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case16(self, config_value, expected_value):
        """上位机可配置CI信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_CI(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['CI']

        if config_value != '1122222222222222cccafsQWQEREQR$%^@$#@#$#@$@$%EGFHHGJ~ADV':
            assert actual_result == expected_value, f"CI配置失败: 期望{expected_value}, 实际{actual_result}"
        else:
            assert actual_result != expected_value, f"CI配置失败: 不期望{expected_value}, 实际{actual_result}"

    @pytest.mark.parametrize("config_value,expected_value", [
        (480, 480),
        (-1440, -1440),
        (1440, 1440)
    ])
    def test_Function_AcuDC320_Sprint2_003_01_case17(self, config_value, expected_value):
        """上位机可配置time_zone_shift信息 - 参数化版本"""
        self.utils.connect_device()
        self.utils.configure_time_zone_shift(value=config_value)
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)['RD']
        for i in actual_result:
            time_zone = i['TM']
            if config_value > 0:
                timezone_str = time_zone.split('+')[1].split()[0]
                hours = int(timezone_str[:2])
                minutes = int(timezone_str[2:4])
                total_minutes = hours * 60 + minutes
            else:
                timezone_str = time_zone.split('-')[-1][:-2].split()[0]
                hours = int(timezone_str[:2])
                minutes = int(timezone_str[2:4])
                total_minutes = -hours * 60 - minutes
            assert total_minutes == expected_value, f"时区配置失败: 期望{expected_value}, 实际{total_minutes}"

    def test_Function_AcuDC320_Sprint2_003_01_case18(self):
        """上位机可发送时钟同步命令，进行时间同步成功"""
        self.utils.connect_device()
        self.utils.configure_time_Sync_status()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading')
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_in_progress')
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Reading_page\Transaction_in_progress\sync'):
            assert False, f'时间同步失败'

    def test_Function_AcuDC320_Sprint2_003_01_case19_1(self):
        """复位、上电后，电表time Sync status变为未同步"""
        self.utils.connect_device()
        self.utils.configure_time_Sync_status()
        self.utils.reboot_device()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading')
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_in_progress')
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Reading_page\Transaction_in_progress\not_sync'):
            assert False, f'时间同步失败'

    def test_Function_AcuDC320_Sprint2_003_01_case20(self):
        """上下电，交易配置未丢失 """
        self.utils.connect_device()
        self.utils.configure_identification_status(True)
        self.utils.configure_identification_level('TRUSTED')
        self.utils.configure_identification_flag1('RFID_PLAIN')
        self.utils.configure_identification_flag2('OCPP_RS')
        self.utils.configure_identification_flag3('ISO15118_NONE')
        self.utils.configure_identification_flag4('PLMN_RING')
        self.utils.configure_identification_type('UNDEFINED')
        self.utils.configure_identification_data('4525asddf$%@!')
        self.utils.configure_Tariff_Text('4525asddf$%@!')
        self.utils.reboot_device()
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        result = self.helper.parse_ocmf(log_data)
        pytest.assume(result['IS'] == True, f"重启后配置发生变化,预期True,实际{result['IS']}")
        pytest.assume(result['IL'] == 'TRUSTED', f"重启后配置发生变化,预期TRUSTED,实际{result['IL']}")
        pytest.assume(result['IF'][0] == 'RFID_PLAIN', f"重启后配置发生变化,预期RFID_PLAIN,实际{result['IF'][0]}")
        pytest.assume(result['IF'][1] == 'OCPP_RS', f"重启后配置发生变化,预期OCPP_RS,实际{result['IF'][1]}")
        pytest.assume(result['IF'][2] == 'ISO15118_NONE', f"重启后配置发生变化,预期ISO15118_NONE,实际{result['IF'][2]}")
        pytest.assume(result['IF'][3] == 'PLMN_RING', f"重启后配置发生变化,预期PLMN_RING,实际{result['IF'][3]}")
        pytest.assume(result['IT'] == 'UNDEFINED', f"重启后配置发生变化,预期UNDEFINED,实际{result['IT']}")
        pytest.assume(result['ID'] == '4525asddf$%@!', f"重启后配置发生变化,预期4525asddf$%@!,实际{result['ID']}")
        pytest.assume(result['TT'] == '4525asddf$%@!', f"重启后配置发生变化,预期4525asddf$%@!,实际{result['TT']}")

    def test_Function_AcuDC320_Sprint2_003_03_case1(self):
        """交易日志完整性。每条交易日志包含：交易ID、交易开始时：时间戳（Timestamp）
        进线电能（Import energy）、出线电能（Export energy），交易完成时：时间戳（Timestamp）、
        进线电能（Import energy）、出线电能（Export energy）、充电桩配置及交易信息"""
        self.utils.connect_device()
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)
        self.helper.validate_json(actual_result)

    def test_Function_AcuDC320_Sprint2_003_03_case11(self):
        """铅封打开，清除交易日志成功"""
        self.utils.construct_transaction_logs(40, clear=False)
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\clear_log')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Confirm')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        time.sleep(1)
        if self.helper.check_image_exists(
                r'page_elements\Acuview_public\Reading_page\Transaction_Log\clear_successful'):
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        else:
            pytest.fail('交易日志清理失败')
        current_log_num = int(self.helper.quick_ocr_by_config('T Used Records'))
        assert current_log_num == 0, '交易日志提示清理失败，Used Records值不为0'

    @pytest.mark.parametrize("config_value,expected_value", [
        (49, 49),
        (50, 50),
        (51, 50),
    ])
    def test_Function_AcuDC320_Sprint2_003_03_case2(self, config_value, expected_value):
        """读取交易日志，上位机选择 Read latest 50 Records 读取正常 """
        # 写入并读取日志
        self.utils.construct_transaction_logs(config_value)
        self.utils.read_logs()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\save_to_file')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        self.helper.hotkey('ctrl', 'c')
        # 获取文件名
        file_name = pyperclip.paste()
        self.helper.hotkey('enter')
        time.sleep(5)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        # 点击第一行日志
        self.helper.double_click_pos((1079, 364))
        # 复制日志
        self.helper.hotkey('ctrl', 'a')
        time.sleep(2)
        self.helper.hotkey('ctrl', 'c')
        result = pyperclip.paste()
        file_path = rf'{self.app_root_path}\Export\{file_name}'
        self.helper.check_csv_file(file_path, expected_value, result, 'T')

    @pytest.mark.parametrize("config_value,expected_value", [
        (999, 999),
        (1000, 1000),
        (1001, 1000),
    ])
    def test_Function_AcuDC320_Sprint2_003_03_case3(self, config_value, expected_value):
        """读取交易日志，上位机选择 Read latest 1000 Records 读取正常 """
        self.utils.construct_transaction_logs(config_value)
        self.helper.click_pos((1005, 208))
        self.helper.click_pos((928, 245))
        self.utils.read_logs()
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\save_to_file')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        self.helper.hotkey('ctrl', 'c')
        # 获取文件名
        file_name = pyperclip.paste()
        self.helper.hotkey('enter')
        time.sleep(5)
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        # 点击第一行日志
        self.helper.double_click_pos((1079, 364))
        # 复制日志
        self.helper.hotkey('ctrl', 'a')
        time.sleep(2)
        self.helper.hotkey('ctrl', 'c')
        result = pyperclip.paste()
        file_path = rf'{self.app_root_path}\Export\{file_name}'
        self.helper.check_csv_file(file_path, expected_value, result, 'T')

    @pytest.mark.parametrize("config_value,expected_value", [
        (50, 50),
        (1000, 1000),
    ])
    def test_Function_AcuDC320_Sprint2_003_03_case6(self, config_value, expected_value):
        """上位机读取日志，循环执行读取50条、1000条日志20次，读取正常"""
        self.utils.construct_transaction_logs(config_value, clear=False)
        for i in range(20):
            if config_value == 50:
                self.utils.read_logs()
            else:
                self.helper.click_pos((1005, 208))
                self.helper.click_pos((928, 245))
                self.utils.read_logs()

    def test_Function_AcuDC320_Sprint2_003_01_case12(self):
        self.utils.connect_device()
        self.utils.configure_Transaction_Timeout(value=3600)
        self.utils._update_configuration()
        self.utils.long_time_charging(3600)
        if not self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\no_charging'):
            pytest.fail('充电时间超过3600秒仍未结束充电')

    def test_Function_AcuDC320_Sprint2_003_01_case14(self):
        self.utils.connect_device()
        self.utils.configure_Transaction_Timeout(value=0)
        self.utils._update_configuration()
        self.utils.long_time_charging(1800)
        if self.helper.check_image_exists(
                r'page_elements\Acuview_public\Setting_page\Test_Charging\operate_failed'):
            pytest.fail('充电时间超过1800秒已经结束充电，设置0不生效')

    def test_Function_AcuDC320_Sprint2_003_03_case7(self):
        """上位机读取日志时，点击停止按钮，可停止日志读取"""
        self.utils.construct_transaction_logs(50, clear=False)
        # 点击Read_Log
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Read_Log')
        time.sleep(2)
        # 点击Stop
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\Stop')
        self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
        if self.helper.check_image_exists(
                r'page_elements\Acuview_public\Reading_page\Transaction_Log\Read_Failed'):
            pytest.fail('读取日志失败')
        elif self.helper.check_image_exists(
                r'page_elements\Acuview_public\Reading_page\Transaction_Log\Reading_Successful'):
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')

    def test_Function_AcuDC320_Sprint2_003_03_case8(self):
        """验证掉电重启，交易日志非易失性"""
        current_log_num = self.utils.construct_transaction_logs(50, clear=False, connect=False)
        self.utils.reboot_device()
        # 点击Reading
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Reading')
        # 点击Transaction_Log
        self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log')
        self.utils.read_logs()
        used = int(self.helper.quick_ocr_by_config('T Used Records'))
        assert current_log_num == used, f'重启前数量{current_log_num}，重启后数量{used}'

    def test_Function_AcuDC320_Sprint2_003_03_case12_1(self):
        """使用默认交易配置触发交易日志，检查交易日志配置报文是否正确"""
        self.utils.connect_device()
        self.utils.perform_charging_cycle()
        log_data = self.utils.read_and_parse_transaction_log()
        actual_result = self.helper.parse_ocmf(log_data)
        pytest.assume(actual_result['GI'] == 'ACCUENERGY AcuDC-320',
                      f"GI值错误,预期ACCUENERGY AcuDC-320,实际{actual_result['GI']}")
        pytest.assume(actual_result['MM'] == 'Acu-DC261-1000-A2-P2',
                      f"GI值错误,预期Acu-DC261-1000-A2-P2,实际{actual_result['MM']}")
        pytest.assume(actual_result['MF'] == '1.02', f"GI值错误,预期1.02,实际{actual_result['MF']}")

    def test_Function_AcuDC320_Sprint2_003_03_case1_1(self):
        """频繁触发交易和结束交易，used Records小于等于Max Transactio"""
        self.utils.construct_transaction_logs(5000, clear=False)
        max = int(self.helper.quick_ocr_by_config('Max Transaction Id'))
        used = int(self.helper.quick_ocr_by_config('T Used Records'))
        assert used <= max, 'Used Records大于Max Transaction'

    @pytest.mark.parametrize("config_value", [
        '1',
        '10',
        '61440',
        '61441'
    ])
    def test_Function_AcuDC320_Sprint2_003_03_case4(self, config_value):
        """读取交易日志，上位机选择Read 1000 Records (from selected Record)读取正常"""
        self.utils.construct_transaction_logs(61440, clear=False)
        self.helper.click_pos((1005, 208))
        self.helper.click_pos((916, 265))
        self.helper.click_pos((1198, 204))
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(config_value)
        if config_value == '61441':
            self.helper.click_pos((1120, 272))
            Start_Id = int(self.helper.quick_ocr_by_config('Start Id'))
            if Start_Id == int(config_value):
                pytest.fail('Start_Id配置61441成功')
        else:
            self.utils.read_logs()
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\save_to_file')
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            self.helper.hotkey('ctrl', 'c')
            # 获取文件名
            file_name = pyperclip.paste()
            self.helper.hotkey('enter')
            time.sleep(5)
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            # 点击第一行日志
            self.helper.double_click_pos((1079, 364))
            # 复制日志
            self.helper.hotkey('ctrl', 'a')
            time.sleep(2)
            self.helper.hotkey('ctrl', 'c')
            result = pyperclip.paste()
            file_path = rf'{self.app_root_path}\Export\{file_name}'
            used = int(self.helper.quick_ocr_by_config('T Used Records'))
            read_num = used - int(config_value) + 1
            if read_num > 1000:
                expected_data_len = 1000
            else:
                expected_data_len = read_num

            self.helper.check_csv_file(file_path, expected_data_len, result, 'T')

    @pytest.mark.parametrize("config_value", [
        '1',
        '10',
        '61440',
        '61441'
    ])
    def test_Function_AcuDC320_Sprint2_003_03_case5(self, config_value):
        """读取交易日志，上位机选择Read 64000 Records (from selected Record)读取正常"""
        self.utils.construct_transaction_logs(300, clear=False)
        self.helper.click_pos((1005, 208))
        self.helper.click_pos((925, 280))
        self.helper.click_pos((1198, 204))
        self.helper.hotkey('ctrl', 'a')
        self.helper.paste_text(config_value)
        if config_value == '61441':
            self.helper.click_pos((1120, 272))
            Start_Id = int(self.helper.quick_ocr_by_config('Start Id'))
            if Start_Id == int(config_value):
                pytest.fail('Start_Id配置61441成功')
        else:
            self.utils.read_logs()
            self.helper.click_image(r'page_elements\Acuview_public\Reading_page\Transaction_Log\save_to_file')
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            self.helper.hotkey('ctrl', 'c')
            # 获取文件名
            file_name = pyperclip.paste()
            self.helper.hotkey('enter')
            time.sleep(5)
            self.helper.click_image(r'page_elements\Acuview_public\Setting_page\Yes')
            # 点击第一行日志
            self.helper.double_click_pos((1079, 364))
            # 复制日志
            self.helper.hotkey('ctrl', 'a')
            time.sleep(2)
            self.helper.hotkey('ctrl', 'c')
            result = pyperclip.paste()
            file_path = rf'{self.app_root_path}\Export\{file_name}'
            used = int(self.helper.quick_ocr_by_config('T Used Records'))
            read_num = used - int(config_value) + 1
            if read_num > 6400:
                expected_data_len = 6400
            else:
                expected_data_len = read_num
            self.helper.check_csv_file(file_path, expected_data_len, result, 'T')

    def test_Function_AcuDC320_Sprint2_003_03_case10(self, ):
        self.utils.construct_transaction_logs(61400, clear=False)
        with ModbusClient(ModbusProtocol.TCP) as client:
            client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 42')
            client.send_custom_message('00 01 00 00 00 09 01 10 52 00 00 01 02 00 45')
            current_log_num = client.parse_data(client.validate_register_value('max Transaction Id'))
        assert current_log_num== '61400', "可以写入第61401条交易日志"

    def test_Function_AcuDC320_Sprint2_002_01_case9(self):
        """构建Transaction log满，系统状态：Fatal Error，生成一条Echilog，记录旧新值、时间戳、ID"""
        self.utils.construct_transaction_logs(61400, clear=False, connect=False)
        self.utils.connect_device()
        result = self.utils.read_echilog_multi_line(1)
        time_value, type_value, old_value, new_value = result[0]
        pytest.assume(type_value == "Echilog Full",
                      f"类型不正确，期望: Echilog Full, 实际: {type_value}")


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--html=report.html", "--self-contained-html"])
