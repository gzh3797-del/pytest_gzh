# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case06
类别：C — 正向推送验证
功能：Rapid Logger 通过 HTTP/HTTPS 推送 CSV，验证文件路径、格式、推送频率及数据准确性

【预置条件】
  1. HTTP 服务器已启动，Post Channel 3=HTTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 HTTP 接收目录

【测试步骤】
  1. 配置 Rapid Logger：
       Enable=True  PostChannel=3(HTTP/HTTPS)  Format=CSV  Length=10 minute
       TimestampFormat=ISO8601 Format  NameFormat=Time Interval Format
       Prefix=meter2_RapidLogger  Interval=1 minute
  2. 等待最长 720s，轮询 HTTP 目录出现 .csv 文件
  3. 验证文件名格式、时间戳格式、推送间隔、行数、数据准确性
  4. 将 Rapid Logger 设为 Disable

【预期结果】
  - HTTP 目录收到 .csv 文件，格式、频率、数据均正确
"""
from helpers import run_rapid_push_case

CASE_ID     = "TestCase_AcuHMI_003_04_case06"
PROTOCOL    = "HTTP"
FILE_FORMAT = "csv"
FILE_LENGTH = "10 minute"
TS_FMT      = "ISO8601 Format"
NAME_FMT    = "Time Interval Format"
PREFIX      = "meter2_RapidLogger"
INTERVAL    = "1 minute"


def test_case(pool, driver):
    run_rapid_push_case(
        CASE_ID, PROTOCOL, FILE_FORMAT,
        FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
        pool, driver, full_verify=True,
    )
