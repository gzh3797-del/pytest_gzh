import os
import time
import logging
import pytest
from Config.IOM.modbus_get_attr import excel_append_ai_measurement, \
    get_all_ai_y_measurements, get_single_ai_y_measurement
from Config.IOM.modbus_set_attr import set_all_ai_param, set_all_ai_top_bot, set_ai_param, iom_test
from Source.CL3021.source_control import set_dc, close_dc, close_dc_all
from Config.IOM.modbus_connet import ModbusRtuOrTcp

case_number = 10


class TestAi:
    # @pytest.mark.critical
    # def test_all_ai_voltage(self,test_data):
    #     """
    #     测试所有AI口的电压测量值，并输出到xlsx表格
    #     : voltage: 输入一组电压值
    #     : expected: 预期测量范围
    #     :return:
    #     """
    #     logging.info("开始执行测试用例：test_all_ai_voltage")
    #     # for i in range(len(test_data["setting"])):
    #     i = 2
    #     logging.info(f"使用测试数据中的第{i+1}条数据")
    #     # setting = test_data["setting"][i]
    #     voltage = test_data["voltage"][i]
    #     expected = test_data["expected"][i]
    #     # set_all_ai_param(setting)
    #     for t in range(len(voltage)):  # 循环输入的电流值
    #         set_dc(voltage[t], 0)
    #         logging.info(
    #           f"配置输入电压为: {voltage[t]}V时：=====================================================================")
    #         time.sleep(7)
    #         measurement_data = get_all_ai_y_measurements()
    #         close_dc(1)
    #         for n in range(16):  # 循环16个ai口
    #             measurement = measurement_data[f"AI{n+1}"]
    #             logging.info(f"AI{n+1}口输入电压{voltage[t]}V，实际测量值为{measurement}，预期范围在{expected[t]}, "
    #                          f"判定结果为：{excel_append_ai_measurement(n+1, 0, voltage[t], measurement, expected[t])}")
    #         logging.info(f"{voltage[t]}V测试结束=====================================================================")
    #     close_dc_all()

    def test_single_ai_current(self):
        modbus_client = ModbusRtuOrTcp()
        register = modbus_client.read_measurement(address=0x3000, count=1, slave=1)
        print(register)

    # @pytest.mark.critical
    # @pytest.mark.parametrize("ai_number", [x for x in range(2, 8)])
    # def test_single_ai_current(self, test_data, ai_number):
    #     """
    #     测试所有AI口的电流测量值
    #     根据电流值列表，控源输出电流（A）；再从板子获取单个个AI口的测量数据（ma），并输出到xlsx表格
    #     每个AI口结束需要在7s内手动切换到下一个AI口
    #     : ai_number: ai口号
    #     : current: 输入一组电流值
    #     : expected: 预期测量范围
    #     :return:
    #     """
    #     time.sleep(7)
    #     logging.info(f"AI{ai_number}测试开始")
    #     current = test_data["current"][case_number-1]
    #     expected = test_data["expected"][case_number-1]
    #     setting_top_bot = test_data["setting_top_bot"][case_number-1]
    #     if ai_number == 1:
    #         set_all_ai_top_bot(setting_top_bot)
    #         logging.info(f"top和bot={setting_top_bot}配置成功")
    #     for n in range(len(current)):
    #         set_dc(0, current[n])
    #         time.sleep(7)
    #         print(ai_number)
    #         measurement_data = get_single_ai_y_measurement(ai_number)
    #         close_dc(2)
    #         logging.info(f"现在执行AI {ai_number}，输入电流为{current[n]*1000}mA，实际测量值为{measurement_data}，预期范围在{expected[n]}, "
    #                      f"判定结果为：{excel_append_ai_measurement(ai_number, 2, current[n], measurement_data, expected[n])}")
    #     close_dc_all()

    #     if ai_number == 16:
    #         logging.info(f"所有AI口测试结束！********************************************************")
    #     else:
    #         logging.info(
    #             f"AI {ai_number} 测试结束，你有7s时间切换到AI {ai_number + 1}！**********************************************")