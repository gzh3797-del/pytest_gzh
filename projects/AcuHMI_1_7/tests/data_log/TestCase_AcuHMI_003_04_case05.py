# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case05
类别：C — 正向推送验证
功能：Rapid Logger 通过 SFTP 推送 CSV，验证文件路径、格式、推送频率及数据准确性

【预置条件】
  1. SFTP 服务器已启动，Post Channel 2=SFTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 SFTP 接收目录

【测试步骤】
  1. 配置 Rapid Logger：
       Enable=True  PostChannel=2(SFTP)  Format=CSV  Length=5 minute
       TimestampFormat=UTC Seconds  NameFormat=Time Interval Format
       Prefix=meter1_RapidLogger  Interval=1 minute
  2. 等待最长 450s，轮询 SFTP 目录出现 .csv 文件
  3. 验证文件名格式（Time Interval Format）、时间戳格式、推送间隔、行数、数据准确性
  4. 将 Rapid Logger 设为 Disable

【预期结果】
  - SFTP 目录收到 .csv 文件，格式、频率、数据均正确
"""
from helpers import run_rapid_push_case

CASE_ID     = "TestCase_AcuHMI_003_04_case05"
PROTOCOL    = "SFTP"
FILE_FORMAT = "csv"
FILE_LENGTH = "5 minute"
TS_FMT      = "UTC Seconds"
NAME_FMT    = "Time Interval Format"
PREFIX      = "meter1_RapidLogger"
INTERVAL    = "1 minute"


def test_case(pool, driver):
    run_rapid_push_case(
        CASE_ID, PROTOCOL, FILE_FORMAT,
        FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
        pool, driver, full_verify=True,
    )
