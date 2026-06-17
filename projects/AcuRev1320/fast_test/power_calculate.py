#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:power_calculate.py
功能描述:功率计算
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""

import math


class CalculatePower:
    def __init__(self):
        pass

    @classmethod
    def calculate_active_power(cls, voltage, current, voltage_angle, current_angle):
        """
        计算有功功率
        :param voltage:电压
        :param current:电流
        :param voltage_angle:电压相位角度
        :param current_angle:电流相位角度
        :return:
        """
        voltage_current_angle = voltage_angle - current_angle
        active_power = (voltage * current * math.cos(math.radians(voltage_current_angle)))
        return active_power / 1000

    @classmethod
    def calculate_reactive_power(cls, voltage, current, voltage_angle, current_angle):
        """
        计算无功功率
        :param voltage:电压
        :param current:电流
        :param voltage_angle:电压相位角度
        :param current_angle:电流相位角度
        :return:无功功率
        """
        voltage_current_angle = voltage_angle - current_angle
        reactive_power = (voltage * current * math.sin(math.radians(voltage_current_angle)))
        return reactive_power / 1000

    @classmethod
    def calculate_apparent_power(cls, voltage, current):
        """
        计算视在功率
        :param voltage:电压
        :param current:电流
        :return:视在功率
        """
        apparent_power = (voltage * current)
        return apparent_power / 1000

    @classmethod
    def calculate_power_factor(cls, voltage_angle, current_angle):
        """
        计算power factor
        :param voltage_angle:电压相位角度
        :param current_angle:电流相位角度
        :return:power factor
        """
        voltage_current_angle = voltage_angle - current_angle
        power_factor = math.cos(math.radians(voltage_current_angle))
        return power_factor
