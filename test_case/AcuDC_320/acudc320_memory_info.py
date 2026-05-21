#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:acudc320_memory_info.py
功能描述:获取寄存器地址
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""

import json
from acudc320_excel_operate import HandleExcel, json_to_dict, OUTPUT_JSON_PATH


class SlaveId:
    slave_id = 1


class MemoryReg:
    reg_uint32 = 2
    reg_int32 = 2
    reg_uint16 = 6
    reg_float32 = 2


class MemoryAddr:
    memory_data_dict = json_to_dict(json_path=OUTPUT_JSON_PATH)
    # memory_data_dict = HandleExcel().data_dict

    u_rms_addr = memory_data_dict["readings"].get("V(Measured or Compensated)", 0)
    i1_rms_addr = memory_data_dict["readings"].get("Current 1", 0)
    i2_rms_addr = memory_data_dict["readings"].get("Current 2", 0)
    i_sum_rms_addr = memory_data_dict["readings"].get("Current Sum", 0)
    p1_rms_addr = memory_data_dict["readings"].get("Power 1", 0)
    p2_rms_addr = memory_data_dict["readings"].get("Power 2", 0)
    p_sum_rms_addr = memory_data_dict["readings"].get("Power Sum", 0)
    u_rf_rms_addr = memory_data_dict["readings"].get("Voltage Ripple Factor", 0)
    i1_rf_rms_addr = memory_data_dict["readings"].get("Current Ripple Factor 1", 0)
    i2_rf_rms_addr = memory_data_dict["readings"].get("Current Ripple Factor 2", 0)
    i1_demand_rms_addr = memory_data_dict["readings"].get("Demand Current 1", 0)
    i2_demand_rms_addr = memory_data_dict["readings"].get("Demand Current 2", 0)
    i_sum_demand_rms_addr = memory_data_dict["readings"].get("Demand Current Sum", 0)
    p1_demand_rms_addr = memory_data_dict["readings"].get("Demand Power 1", 0)
    p2_demand_rms_addr = memory_data_dict["readings"].get("Demand Power 2", 0)
    p_sum_demand_rms_addr = memory_data_dict["readings"].get("Demand Power Sum", 0)
    v1_Comp_rms_addr = memory_data_dict["readings"].get("V(Measured)", 0)
    v2_Comp_rms_addr = memory_data_dict["readings"].get("V(Compensated)", 0)
