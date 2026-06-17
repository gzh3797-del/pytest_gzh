"""
终端运行：
| **AO 电压精度** | `pytest -m ao_v`
| **AO 电流精度** | `pytest -m ao_c`
"""
import pytest
from projects.IOM.operation.IOM_set_attr import set_all_ao_param, iom_test


class TestAO:
    @pytest.mark.ao_v
    def test_ao_voltage(self, timer, modbus_client, test_data, index):
        """
        AO电压用例
        :param timer: fixture 前置计时以及异常处理
        :param modbus_client: fixture modbus客户端
        :param test_data: fixture 测试数据
        :param index: fixture 钩子函数，测试执行前自动识别用例数
        :return:
        """
        ao_number = test_data['ao_number'][index]
        type_line_value = test_data['type_line'][index]
        parameter_value = test_data['parameter'][index]
        voltage = test_data['voltage'][index]
        expected = test_data['expected'][index]
        set_all_ao_param(modbus_client, type_line_value, parameter_value, ai_ao_number=ao_number)
        res = iom_test(modbus_client, ao_number=ao_number, ao_voltage=voltage, expected=expected)
        assert res, "AO电压精度不符合预期，用例失败"

    @pytest.mark.ao_c
    def test_ao_current(self, timer, modbus_client, test_data, index):
        ao_number = test_data['ao_number'][index]
        type_line_value = test_data['type_line'][index]
        parameter_value = test_data['parameter'][index]
        current = test_data['current'][index]
        expected = test_data['expected'][index]
        set_all_ao_param(modbus_client, type_line_value, parameter_value, ai_ao_number=ao_number)
        res = iom_test(modbus_client, ao_number=ao_number, ao_current=current, expected=expected)
        assert res, "AO电流精度不符合预期，用例失败"
