#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""
文件名称:acudc320_test_report_table_heading.py
功能描述:定义测试结果列名
创建日期:2025-08-05
作者:王洋
版本:v1.0
修改记录:
"""


class TableTitle:
    REAL_TIME_MEASURE_COLUMNS_OF_DC320 = [
        '测试用例',
        'voltage_accuracy',
        'current_accuracy',
        'power_accuracy',
        '电压输入值',
        '电流输入值',

        '电压测量最小值',
        '电压精度最小值',
        '电压测量最大值',
        '电压精度最大值',
        '电压测量平均值',
        '电压精度平均值',
        '电压精度测试结果',

        '电流测量最小值',
        '电流精度最小值',
        '电流测量最大值',
        '电流精度最大值',
        '电流测量平均值',
        '电流精度平均值',
        '电流精度测试结果',

        '功率输入值',
        '功率测量最小值',
        '功率精度最小值',
        '功率测量最大值',
        '功率精度最大值',
        '功率测量平均值',
        '功率精度平均值',
        '功率精度测试结果',
    ]



if __name__ == '__main__':
    out_name = [i for i in TableTitle.REAL_TIME_MEASURE_COLUMNS_OF_DC260]
    print(out_name)
