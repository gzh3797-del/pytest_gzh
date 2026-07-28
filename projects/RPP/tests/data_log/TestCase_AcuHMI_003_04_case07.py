# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case07
类别：C — 正向推送验证（快速间隔 10 s）
功能：Rapid Logger 以 10 秒间隔通过 FTP 推送 CSV，验证文件内容及推送频率

【预置条件】
  1. FTP 服务器已启动，Post Channel 1=FTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 FTP 接收目录
  4. 所有 Physical Devices 的 Poll Interval 须 ≤ 10s（本用例自动设置并还原）

【测试步骤】
  1. 将所有 Physical Devices 的 Poll Interval 改为 10s
  2. 配置 Rapid Logger：
       Enable=True  PostChannel=1(FTP)  Format=CSV  Length=1 minute
       TimestampFormat=Local Time String  NameFormat=UTC Timestamp
       Prefix=meter0_RapidLogger  Interval=10 seconds
  3. 等待最长 120s，轮询 FTP 目录出现 .csv 文件
  4. 验证文件名格式、时间戳格式（±30%）、行数（约 6 行）
  5. 将 Rapid Logger 设为 Disable
  6. 恢复所有 Physical Devices 的 Poll Interval 为原始值

【预期结果】
  - FTP 目录收到 .csv 文件，相邻行间隔约 10 s，数据行数约 6 行
"""
from helpers import run_rapid_push_case
from datalog_page import PhysicalDevicePollHelper

CASE_ID     = "TestCase_AcuHMI_003_04_case07"
PROTOCOL    = "FTP"
FILE_FORMAT = "csv"
FILE_LENGTH = "1 minute"
TS_FMT      = "Local Time String"
NAME_FMT    = "UTC Timestamp"
PREFIX      = "meter0_RapidLogger"
INTERVAL    = "10 seconds"


def test_case(pool, driver):
    poll_helper = PhysicalDevicePollHelper(driver)
    originals = {}
    try:
        originals = poll_helper.set_all(10)
        # 等待设备以新 Poll Interval 生效
        driver.wait_for_timeout(5000)
        run_rapid_push_case(
            CASE_ID, PROTOCOL, FILE_FORMAT,
            FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
            pool, driver,
        )
    finally:
        if originals:
            poll_helper.restore_all(originals)
