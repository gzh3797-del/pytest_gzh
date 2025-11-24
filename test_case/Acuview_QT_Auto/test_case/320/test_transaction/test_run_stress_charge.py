#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from comm.modbus_rtu_tcp import ModbusRtuOrTcp
from datetime import datetime
import struct
import json
from comm.source_control import *
import logging
import pandas as pd
import os
import time

def run_stress_charge(charge_count, charge_time, charge_voltage=0.0, charge_current=0.0, command='E'):
    """
    自动运行充电，并获取交易日志，检查电表电压电流，充电开始结束时间，充电能量与实际是否一致
    """

    def write_transaction_command(modbus, command, slave=1, channel=1):
        """
        开始充电功能

        Args:
            modbus: ModbusRtuOrTcp实例
            command: 充电命令字符 ('B', 'E', 'A', 或其他字符)
            slave: 从站地址，默认为1
            channel: 充电通道，默认为通道1
        Returns:
            Modbus响应对象或异常信息
        """
        valid_commands = {'B': '开始充电', 'E': '结束充电', 'A': '中止充电'}
        if channel == "2":
            Modbus_address = 0x5478
        else:
            Modbus_address = 0x5200
        try:
            # 将字符转换为寄存器值（ASCII码）
            transaction_value = ord(command)
            resp = modbus.write_registers(address=Modbus_address, values=[transaction_value], slave=slave)

            command_description = valid_commands.get(command, f"未知命令 '{command}'")

            if hasattr(resp, 'isError') and resp.isError():
                logging.error(f"{command_description}失败 - 从站{slave} - 错误: {resp}")
                return f"{command_description}失败: {resp}"
            else:
                logging.info(
                    f"{command_description}成功 - 从站{slave} - 寄存器0x5200写入值: {transaction_value} ('{command}')")
                return resp

        except Exception as e:
            logging.error(f"{command_description}异常 - 从站{slave} - 错误: {e}")
            return e

    def read_power_data(modbus, slave=1):
        """
        带详细日志记录的电参数读取 - 32位数据转为浮点数
        """
        try:
            # 读取电压 (32位浮点数)
            voltage_data = modbus.read_measurement(address=0x3000, count=2, slave=slave)
            if voltage_data == "resp is error" or isinstance(voltage_data, Exception):
                return "读取电压失败"

            # 读取电流 (32位浮点数)
            current_data = modbus.read_measurement(address=0x3002, count=2, slave=slave)
            if current_data == "resp is error" or isinstance(current_data, Exception):
                return "读取电流失败"

            # 读取功率 (32位浮点数)
            power_data = modbus.read_measurement(address=0x3004, count=2, slave=slave)
            if power_data == "resp is error" or isinstance(power_data, Exception):
                return "读取功率失败"

            # 将32位数据转换为浮点数
            def words_to_float(high_word, low_word):
                """将两个16位寄存器值转换为32位浮点数"""
                int32_value = (high_word << 16) | low_word
                float_bytes = struct.pack('>I', int32_value)  # 大端字节序
                float_value = struct.unpack('>f', float_bytes)[0]
                return float_value

            # 处理数据 - 转换为浮点数
            voltage = words_to_float(voltage_data[0], voltage_data[1])
            current = words_to_float(current_data[0], current_data[1])
            power = words_to_float(power_data[0], power_data[1])

            # 格式化显示，保留适当小数位
            result = {
                'voltage': round(voltage, 4),
                'current': round(current, 4),
                'power': round(power, 4)
            }

            logging.info(
                f"从站{slave} 电参数读取成功: 电压={result['voltage']}V, 电流={result['current']}A, 功率={result['power']}W")
            print(f"电参数读取成功: 电压={result['voltage']}V, 电流={result['current']}A, 功率={result['power']}KW")

            return result

        except Exception as e:
            error_msg = f"从站{slave} 读取电参数异常: {e}"
            logging.error(error_msg)
            print(error_msg)
            return error_msg

    def read_OCMF_num(modbus, slave=1):
        """
        获取交易日志总数量
        """
        try:
            # 读取2个寄存器
            registers = modbus.read_measurement(0x5502, 2, slave)
            # 组合成32位数：第一个寄存器是高16位，第二个是低16位
            total_count = (registers[0] << 16) | registers[1]

            return int(total_count)

        except Exception as e:
            print(f"读取失败: {e}")
            return 0

    def validate_voltage_current(charge_voltage, charge_current, measured_data):
        """
        验证测量的电压电流是否在预期范围内

        Args:
            charge_voltage: 预期电压
            charge_current: 预期电流
            measured_data: 测量的数据字典

        Returns:
            tuple: (是否匹配, 错误信息)
        """
        try:
            # 计算百分比误差（0.2%）
            voltage_tolerance = charge_voltage * 0.001
            current_tolerance = charge_current * 0.005

            # 判断输入的电压电流是否等于读取出来的，考虑0.2%的百分比误差
            voltage_match = abs(charge_voltage - measured_data['voltage']) <= voltage_tolerance
            current_match = abs(charge_current - measured_data['current']) <= current_tolerance

            if voltage_match and current_match:
                logging.info("电压电流匹配成功")
                return True, "电压电流匹配成功"
            else:
                error_msg = f"电压电流不匹配 - 预期: {charge_voltage}V/{charge_current}A, 实际: {measured_data['voltage']}V/{measured_data['current']}A"
                logging.warning(error_msg)
                return False, error_msg

        except KeyError as e:
            error_msg = f"测量数据缺少必要字段: {e}"
            logging.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"验证电压电流时发生错误: {e}"
            logging.error(error_msg)
            return False, error_msg

    def write_OCMF(modbus, id, address=0x5218, slave=1):
        """
        写入OCMF ID并读取相关数据

        Args:
            modbus: ModbusRtuOrTcp实例
            id: OCMF ID值
            slave: 从站地址，默认为1
            address: transactionLog id地址

        Returns:
            包含操作结果的字典
        """
        try:
            # 写入OCMF ID到寄存器
            if id > 0xFFFF:
                id_high = (id >> 16) & 0xFFFF
                id_low = id & 0xFFFF
                values = [id_high, id_low]
            else:
                values = [0, id]

            resp = modbus.write_registers(address=address, values=values, slave=slave)
            time.sleep(5)

            # 检查写入是否成功
            if not resp.isError():
                # 定义要读取的四个寄存器区域
                read_commands = [
                    {'address': 0x521B, 'count': 0x78},
                    {'address': 0x5293, 'count': 0x78},
                    {'address': 0x530B, 'count': 0x78},
                    {'address': 0x5383, 'count': 0x41}
                ]

                combined_data = []

                # 依次读取每个寄存器区域并收集响应数据
                for i, cmd in enumerate(read_commands):
                    try:
                        response_data = modbus.read_measurement(
                            address=cmd['address'],
                            count=cmd['count'],
                            slave=slave
                        )

                        if isinstance(response_data, list):
                            combined_data.extend(response_data)
                        else:
                            combined_data.extend([0] * cmd['count'])

                    except Exception as e:
                        combined_data.extend([0] * cmd['count'])

                # 将combined_data中的ASCII码转换为字符
                ascii_string = registers_to_ascii(combined_data)
                parts = ascii_string.split('|')

                # 第二个部分就是您需要的 JSON 字典
                json_string = parts[1] if len(parts) > 1 else ""

                return {
                    'success': True,
                    'id': id,
                    'OCMF': json_string,
                    'message': 'OCMF写入和数据读取成功'
                }
            else:
                return {
                    'success': False,
                    'error': 'OCMF ID写入失败',
                    'message': resp
                }

        except Exception as e:
            logging.error(f"OCMF操作发生异常: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'message': 'OCMF操作失败'
            }

    def registers_to_ascii(registers):
        """
        将寄存器数据转换为ASCII字符串

        Args:
            registers: 寄存器数据列表

        Returns:
            str: 转换后的ASCII字符串
        """
        ascii_chars = []
        for register in registers:
            high_byte = (register >> 8) & 0xFF
            low_byte = register & 0xFF

            if 32 <= high_byte <= 126:
                ascii_chars.append(chr(high_byte))
            elif high_byte == 0:
                pass
            else:
                ascii_chars.append('.')

            if 32 <= low_byte <= 126:
                ascii_chars.append(chr(low_byte))
            elif low_byte == 0:
                pass
            else:
                ascii_chars.append('.')

        result = ''.join(ascii_chars)
        return result.strip() if result.strip() else result

    def read_json(ocmf_data):
        # 将数据加载为 Python 字典
        if isinstance(ocmf_data, str):
            data_dict = json.loads(ocmf_data.replace("'", '"'))
        else:
            data_dict = ocmf_data  # 如果已经是字典，直接使用

        records = data_dict["RD"]

        # 获取充电开始和结束时间
        start_time = None
        end_time = None

        # 存储能量值
        energy_input_values = []  # 充电进入能量
        energy_output_values = []  # 充电导出能量

        # 遍历记录
        for record in records:
            timestamp = record.get("TM", None)
            energy = record.get("RV", 0)
            ri = record.get("RI", "")

            if timestamp:
                timestamp_cleaned = timestamp.split(" ")[0].split("+")[0]
                timestamp_obj = datetime.strptime(timestamp_cleaned, "%Y-%m-%dT%H:%M:%S,%f")
                if start_time is None:
                    start_time = timestamp_obj
                end_time = timestamp_obj

            # 收集能量值
            if ri == "01-00:01.08.00*FF":
                energy_input_values.append(energy)
            elif ri == "01-00:02.08.00*FF":
                energy_output_values.append(energy)

        # 计算充电时长
        duration = end_time - start_time if start_time and end_time else None
        duration_hours = duration.total_seconds() / 3600 if duration else 0

        # 计算能量差值
        input_energy_diff = energy_input_values[-1] - energy_input_values[0] if len(energy_input_values) >= 2 else 0
        output_energy_diff = energy_output_values[-1] - energy_output_values[0] if len(energy_output_values) >= 2 else 0

        # 返回结果字典
        result = {
            "start_time": start_time,  # datetime对象
            "end_time": end_time,  # datetime对象
            "duration_hours": round(duration_hours, 6),
            "input_energy_kwh": round(input_energy_diff, 9),
            "output_energy_kwh": round(output_energy_diff, 9)
        }
        print(f"OCMF信息：{result}")
        return result

    def validate_energy(recorded_energy, expected_energy, energy_type="能量"):
        """
        验证能量是否符合精度要求
        能量精度: 0.5%

        Args:
            recorded_energy: OCMF记录的能量值(kWh)
            expected_energy: 计算得到的期望能量值(kWh)
            energy_type: 能量类型描述，如"输入"、"输出"

        Returns:
            tuple: (验证结果, 详细消息)
        """
        try:
            # 计算容差 (0.5%)
            energy_tolerance = abs(expected_energy) * 0.005

            # 计算实际误差
            energy_error = abs(recorded_energy - expected_energy)

            # 验证能量
            energy_ok = energy_error <= energy_tolerance

            # 计算相对误差百分比
            if expected_energy != 0:
                relative_error = (energy_error / abs(expected_energy)) * 100
            else:
                relative_error = 0

            # 生成详细消息
            message = (f"{energy_type}能量验证: 记录值={recorded_energy:.6f}kWh, "
                       f"期望值={expected_energy:.6f}kWh, "
                       f"误差={energy_error:.6f}kWh ({relative_error:.3f}%), "
                       f"容差={energy_tolerance:.6f}kWh, 结果={'通过' if energy_ok else '失败'}")

            return energy_ok, message

        except Exception as e:
            error_msg = f"验证{energy_type}能量时发生错误: {str(e)}"
            return False, error_msg

    modbus = ModbusRtuOrTcp()
    results = []

    for i in range(charge_count):
        round_results = []
        round_num = i + 1
        round_errors = []

        try:
            print(f"第{round_num}轮开始")
            sour_output(charge_voltage, charge_current)
            # 开始充电
            try:
                start_charge_time = time.time()
                write_transaction_command(modbus, 'B', slave=1)
                round_results.append(True)
            except Exception as e:
                error_msg = f"开始充电命令失败: {str(e)}"
                round_errors.append(error_msg)
                results.append(False)
                continue

            # 充电时长（最小等待时间）
            time.sleep(max(charge_time, 1))

            # 结束充电
            try:
                end_charge_time = time.time()
                result_end = write_transaction_command(modbus, command, slave=1)
                result2 = not isinstance(result_end, str)
                if not result2:
                    write_transaction_command(modbus, command, slave=1)
                round_results.append(True)
            except Exception as e:
                error_msg = f"结束充电命令失败: {str(e)}"
                round_errors.append(error_msg)
                round_results.append(False)
            # 并行读取功率数据和OCMF操作
            power_result = False
            ocmf_info = None
            end_power_data = None

            try:
                # 读取功率数据
                end_power_data = read_power_data(modbus, slave=1)
                sour_stop()
                power_result, power_msg = validate_voltage_current(charge_voltage, charge_current, end_power_data)

                # OCMF操作
                ocmf_id = read_OCMF_num(modbus)
                print(f"交易日志数量：{ocmf_id}")
                res = write_OCMF(modbus, id=ocmf_id, slave=1)
                ocmf_info = read_json(res['OCMF'])

            except Exception as e:
                error_msg = f"数据读取失败: {str(e)}"
                round_errors.append(error_msg)

            # 输出所有实际值对比
            print(f"\n=== 第{round_num}轮实际值对比 ===")

            # 时间对比
            print("【时间对比】")
            print(f"  系统记录开始时间: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(start_charge_time))}")
            print(f"  系统记录结束时间: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end_charge_time))}")

            if ocmf_info:
                print(f"  OCMF开始时间: {ocmf_info['start_time']}")
                print(f"  OCMF结束时间: {ocmf_info['end_time']}")
                print(f"  开始时间误差: {abs(start_charge_time - ocmf_info['start_time'].timestamp()):.3f}秒")
                print(f"  结束时间误差: {abs(end_charge_time - ocmf_info['end_time'].timestamp()):.3f}秒")

            # 能量对比
            print("【能量对比】")
            if ocmf_info:
                actual_energy = ocmf_info['duration_hours'] * charge_voltage * abs(charge_current) * 0.001

                if charge_current >= 0:
                    recorded_energy = ocmf_info['input_energy_kwh']
                    energy_type = "输入"
                    output_energy = ocmf_info.get('output_energy_kwh', 0)
                else:
                    recorded_energy = ocmf_info['output_energy_kwh']
                    energy_type = "输出"
                    input_energy = ocmf_info.get('input_energy_kwh', 0)

                print(f"  计算期望{energy_type}能量: {actual_energy:.6f} kWh")
                print(f"  OCMF记录{energy_type}能量: {recorded_energy:.6f} kWh")
                print(f"  OCMF记录充电时长: {ocmf_info['duration_hours']:.6f} 小时")

                if charge_current >= 0:
                    print(f"  OCMF记录输出能量: {output_energy:.6f} kWh")
                else:
                    print(f"  OCMF记录输入能量: {input_energy:.6f} kWh")

                energy_diff = abs(recorded_energy - actual_energy)
                energy_ratio = abs((recorded_energy - actual_energy) / actual_energy * 100) if actual_energy != 0 else 0
                print(f"  能量差值: {energy_diff:.6f} kWh")
                print(f"  相对误差: {energy_ratio:.2f}%")

            # 电压电流功率对比
            print("【电压电流功率对比】")
            print(f"  期望电压: {charge_voltage} V")
            print(f"  期望电流: {charge_current} A")
            print(f"  期望功率: {charge_voltage * abs(charge_current):.1f} W")

            if end_power_data:
                actual_voltage = end_power_data.get('voltage', 'N/A')
                actual_current = end_power_data.get('current', 'N/A')
                actual_power = end_power_data.get('power', 'N/A')

                print(f"  实际电压: {actual_voltage} V")
                print(f"  实际电流: {actual_current} A")
                print(f"  实际功率: {actual_power} kW")

                if isinstance(actual_voltage, (int, float)) and isinstance(actual_current, (int, float)):
                    calculated_power = actual_voltage * abs(actual_current)
                    print(f"  计算功率: {calculated_power:.1f} W")

            # 验证结果
            if ocmf_info and end_power_data:
                # 时间验证
                start_time_diff = abs(start_charge_time - ocmf_info['start_time'].timestamp())
                end_time_diff = abs(end_charge_time - ocmf_info['end_time'].timestamp())
                time_ok = start_time_diff <= 2 and end_time_diff <= 2

                # 能量验证
                actual_energy = ocmf_info['duration_hours'] * charge_voltage * abs(charge_current) * 0.001
                recorded_energy = ocmf_info['input_energy_kwh'] if charge_current >= 0 else ocmf_info[
                    'output_energy_kwh']
                energy_ok, energy_msg = validate_energy(recorded_energy, actual_energy,
                                                        "输入" if charge_current >= 0 else "输出")

                round_results.extend([power_result, time_ok, energy_ok])

                # 输出验证结果
                print("\n【验证结果】")
                print(f"  功率数据: [{'PASS' if power_result else 'FAIL'}]")
                print(f"  时间验证: [{'PASS' if time_ok else 'FAIL'}] (允许误差≤2秒)")
                print(f"  能量验证: [{'PASS' if energy_ok else 'FAIL'}]")

            else:
                round_results.extend([False, False, False])
                print(f"第{round_num}轮: 数据获取失败")

            # 记录错误信息
            if round_errors:
                for error in round_errors:
                    logging.error(f"第{round_num}轮: {error}")

            # 本轮结果
            round_success = all(round_results)
            results.append(round_success)
            status = "成功" if round_success else "失败"
            print(f"\n第{round_num}轮总体结果: {status}")

            # 最小间隔等待
            if i < charge_count - 1:
                time.sleep(0.1)

        except Exception as e:
            logging.error(f"第{round_num}轮未知错误: {str(e)}")
            results.append(False)
            continue

    # 快速关闭连接
    try:
        modbus.close()
    except:
        pass

    # 最终统计
    success_count = sum(results)
    final_result = all(results)

    print(f"\n=== 压力测试最终结果 ===")
    print(f"总测试轮次: {charge_count}")
    print(f"成功轮次: {success_count}")
    print(f"失败轮次: {charge_count - success_count}")
    print(f"成功率: {success_count / charge_count * 100:.1f}%")
    print(f"总体结果: [{'PASS' if final_result else 'FAIL'}]")

    return final_result


def test_run_stress_charge_from_excel():
    """
    从stress_charging.xlsx文件读取测试参数，自动运行充电测试
    """

    def read_test_cases_from_excel():
        """
        从stress_charging.xlsx文件读取测试用例

        Returns:
            list: 测试用例列表，每个用例是一个字典
        """
        excel_file = "stress_charging.xlsx"

        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"未找到文件: {excel_file}")

        try:
            # 读取Excel文件
            df = pd.read_excel(excel_file)

            # 检查必要的列
            required_columns = ['charge_count', 'charge_time']
            optional_columns = ['charge_voltage', 'charge_current', 'command']

            # 验证必要列是否存在
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"Excel文件中缺少必要列: {col}")

            # 处理可选列，设置默认值
            test_cases = []
            for index, row in df.iterrows():
                test_case = {
                    'charge_count': int(row['charge_count']),
                    'charge_time': float(row['charge_time']),
                    'charge_voltage': float(row.get('charge_voltage', 0.0)),
                    'charge_current': float(row.get('charge_current', 0.0)),
                    'command': str(row.get('command', 'E')),
                    'case_id': index + 1  # 添加用例ID
                }
                test_cases.append(test_case)

            print(f"从Excel文件读取到 {len(test_cases)} 个测试用例")
            return test_cases

        except Exception as e:
            logging.error(f"读取Excel文件失败: {str(e)}")
            raise

    # 读取测试用例
    test_cases = read_test_cases_from_excel()

    # 执行所有测试用例
    all_results = []

    for test_case in test_cases:
        case_id = test_case['case_id']
        charge_count = test_case['charge_count']
        charge_time = test_case['charge_time']
        charge_voltage = test_case['charge_voltage']
        charge_current = test_case['charge_current']
        command = test_case['command']

        print(f"\n{'=' * 60}")
        print(f"执行测试用例 #{case_id}")
        print(f"{'=' * 60}")
        print(f"参数: charge_count={charge_count}, charge_time={charge_time}s")
        print(f"      charge_voltage={charge_voltage}V, charge_current={charge_current}A")
        print(f"      command='{command}'")
        print(f"{'=' * 60}")

        # 记录用例开始时间
        case_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        try:
            # 调用原有的测试函数
            result = run_stress_charge(
                charge_count=charge_count,
                charge_time=charge_time,
                charge_voltage=charge_voltage,
                charge_current=charge_current,
                command=command
            )

            # 记录用例结果
            case_result = {
                'case_id': case_id,
                'parameters': test_case,
                'result': result,
                'start_time': case_start_time,
                'end_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'status': 'PASS' if result else 'FAIL'
            }

            all_results.append(case_result)

            print(f"\n测试用例 #{case_id} 结果: {'PASS' if result else 'FAIL'}")

        except Exception as e:
            logging.error(f"测试用例 #{case_id} 执行失败: {str(e)}")
            case_result = {
                'case_id': case_id,
                'parameters': test_case,
                'result': False,
                'start_time': case_start_time,
                'end_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'status': 'ERROR',
                'error_message': str(e)
            }
            all_results.append(case_result)
            print(f"\n测试用例 #{case_id} 执行失败: {str(e)}")

        # 用例间等待时间
        if test_cases.index(test_case) < len(test_cases) - 1:
            print(f"\n等待3秒后执行下一个用例...")
            time.sleep(3)

    # 生成最终测试报告
    print(f"\n{'=' * 60}")
    print(f"测试执行完成")
    print(f"{'=' * 60}")

    total_cases = len(all_results)
    passed_cases = len([r for r in all_results if r.get('status') == 'PASS'])
    failed_cases = len([r for r in all_results if r.get('status') == 'FAIL'])
    error_cases = len([r for r in all_results if r.get('status') == 'ERROR'])

    print(f"总测试用例数: {total_cases}")
    print(f"通过: {passed_cases}")
    print(f"失败: {failed_cases}")
    print(f"错误: {error_cases}")
    print(f"成功率: {passed_cases / total_cases * 100:.1f}%")

    # 显示详细结果
    print(f"\n详细结果:")
    for result in all_results:
        status = result['status']
        case_id = result['case_id']
        if status == 'PASS':
            print(f"  用例 #{case_id}: ✅ PASS")
        elif status == 'FAIL':
            print(f"  用例 #{case_id}: ❌ FAIL")
        else:
            print(f"  用例 #{case_id}: 💥 ERROR - {result.get('error_message', '未知错误')}")

    overall_result = all(r.get('status') == 'PASS' for r in all_results)
    print(f"\n总体结果: {'✅ ALL PASS' if overall_result else '❌ SOME TESTS FAILED'}")

    return overall_result, all_results



if __name__ == "__main__":
    try:
        final_result, detailed_results = test_run_stress_charge_from_excel()
        exit(0 if final_result else 1)
    except Exception as e:
        logging.error(f"测试执行失败: {str(e)}")
        exit(1)
