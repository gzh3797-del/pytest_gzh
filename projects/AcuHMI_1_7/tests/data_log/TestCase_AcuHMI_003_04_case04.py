# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case04
类别：C — 正向推送验证（含完整三段比对）
功能：Rapid Logger 通过 FTP 推送 CSV，验证文件路径、格式、推送频率及数据准确性

【预置条件】
  1. FTP 服务器已启动，Post Channel 1=FTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 FTP 接收目录

【测试步骤】
  1. 配置 Rapid Logger：
       Enable=True  PostChannel=1(FTP)  Format=CSV  Length=1 minute
       TimestampFormat=Local Time String  NameFormat=UTC Timestamp
       Prefix=meter0_RapidLogger  Interval=1 minute
  2. 等待最长 120s，轮询 FTP 目录出现 .csv 文件
  3. 验证文件扩展名、文件名格式、时间戳格式、推送间隔、行数
  4. 执行完整三段验证（范围/单位/Modbus 数值）
  5. 将 Rapid Logger 设为 Disable

【预期结果】
  - FTP 目录收到 .csv 文件
  - 文件名以 meter0_RapidLogger 开头，符合 UTC Timestamp 格式
  - 时间戳列符合 Local Time String 格式
  - 相邻行间隔约 1 minute，数据行数与文件长度一致
  - 参数范围、单位、Modbus 数值均正确
"""
from helpers import run_rapid_push_case

CASE_ID     = "TestCase_AcuHMI_003_04_case04"
PROTOCOL    = "FTP"
FILE_FORMAT = "csv"
FILE_LENGTH = "1 minute"
TS_FMT      = "Local Time String"
NAME_FMT    = "UTC Timestamp"
PREFIX      = "meter0_RapidLogger"
INTERVAL    = "1 minute"


def test_case(pool, driver):
    run_rapid_push_case(
        CASE_ID, PROTOCOL, FILE_FORMAT,
        FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
        pool, driver, full_verify=True,
    )
