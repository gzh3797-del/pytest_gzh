# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case09
类别：C — 正向推送验证（快速间隔 30 s）
功能：Rapid Logger 以 30 秒间隔通过 FTP 推送 CSV，验证文件内容及推送频率

【预置条件】
  1. FTP 服务器已启动，Post Channel 1=FTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 FTP 接收目录

【测试步骤】
  1. 配置 Rapid Logger：
       Enable=True  PostChannel=1(FTP)  Format=CSV  Length=1 minute
       TimestampFormat=ISO8601 Format  NameFormat=Time Interval Format
       Prefix=meter0_RapidLogger  Interval=30 seconds
  2. 等待最长 120s，轮询 FTP 目录出现 .csv 文件
  3. 验证文件名格式、时间戳格式（±30%）、行数（约 2 行）
  4. 将 Rapid Logger 设为 Disable

【预期结果】
  - FTP 目录收到 .csv 文件，相邻行间隔约 30 s，数据行数约 2 行
"""
from helpers import run_rapid_push_case

CASE_ID     = "TestCase_AcuHMI_003_04_case09"
PROTOCOL    = "FTP"
FILE_FORMAT = "csv"
FILE_LENGTH = "1 minute"
TS_FMT      = "ISO8601 Format"
NAME_FMT    = "Time Interval Format"
PREFIX      = "meter0_RapidLogger"
INTERVAL    = "30 seconds"


def test_case(pool, driver):
    run_rapid_push_case(
        CASE_ID, PROTOCOL, FILE_FORMAT,
        FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
        pool, driver,
    )
