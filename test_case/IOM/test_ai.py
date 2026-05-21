"""
终端运行：
| **AI 电压精度** | `pytest -m ai_v`
| **AI 电流精度** | `pytest -m ai_c`
"""
import pytest
from test_case.IOM.operation.IOM_set_attr import iom_test, set_all_ai_param


class TestAi:
    @pytest.mark.ai_v
    def test_ai_voltage(self, timer, modbus_client, test_data, index):
        """
        AI电压用例
        :param timer: fixture 前置计时以及异常处理
        :param modbus_client: fixture modbus客户端
        :param test_data: fixture 测试数据
        :param index: fixture 钩子函数，测试执行前自动识别用例数
        :return:
        """
        ai_number = test_data['ai_number'][index]
        type_line_value = test_data['type_line'][index]
        parameter_value = test_data['parameter'][index]
        voltage = test_data['voltage'][index]
        expected = test_data['expected'][index]
        set_all_ai_param(modbus_client, type_line_value, parameter_value, ai_ao_number=ai_number)
        res = iom_test(modbus_client, ai_number=ai_number, ai_voltage=voltage, expected=expected)
        assert res, "AI电压精度不符合预期，用例失败"

    @pytest.mark.ai_c
    def test_ai_current(self, timer, modbus_client, test_data, index):
        ai_number = test_data['ai_number'][index]
        type_line_value = test_data['type_line'][index]
        parameter_value = test_data['parameter'][index]
        current = test_data['current'][index]
        expected = test_data['expected'][index]
        set_all_ai_param(modbus_client, type_line_value, parameter_value, ai_ao_number=ai_number)
        res = iom_test(modbus_client, ai_number=ai_number, ai_current=current, expected=expected)
        assert res, "AI电流精度不符合预期，用例失败"


