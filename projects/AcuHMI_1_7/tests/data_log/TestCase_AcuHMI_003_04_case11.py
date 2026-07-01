# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case11
类别：C — 正向推送验证（JSON，1 second 间隔）
功能：Rapid Logger 以 1 秒间隔通过 FTP 推送 JSON，验证文件内容及推送频率

【预置条件】
  1. FTP 服务器已启动，Post Channel 1=FTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 FTP 接收目录
  4. 所有 Physical Devices 的 Poll Interval 须 ≤ 1s（本用例自动设置并还原）

【测试步骤】
  1. 将所有 Physical Devices 的 Poll Interval 改为 1s
  2. 配置 Rapid Logger：
       Enable=True  PostChannel=1(FTP)  Format=JSON  Length=1 minute
       NameFormat=UTC Timestamp  Prefix=meter0_RapidLogger  Interval=1 second
  3. 等待最长 120s，轮询 FTP 目录出现 .json 文件
  4. 验证文件名格式、时间戳格式、推送间隔（±30%）、行数（约 60 行）
  5. 将 Rapid Logger 设为 Disable
  6. 恢复所有 Physical Devices 的 Poll Interval 为原始值

【预期结果】
  - FTP 目录收到 .json 文件，相邻记录间隔约 1 s，行数约 60 行
"""
from helpers import run_rapid_push_case
from datalog_page import PhysicalDevicePollHelper

CASE_ID     = "TestCase_AcuHMI_003_04_case11"
PROTOCOL    = "FTP"
FILE_FORMAT = "json"
FILE_LENGTH = "1 minute"
TS_FMT      = "UTC Seconds"
NAME_FMT    = "UTC Timestamp"
PREFIX      = "meter0_RapidLogger"
INTERVAL    = "1 second"


def test_case(pool, driver):
    poll_helper = PhysicalDevicePollHelper(driver)
    originals = {}
    try:
        originals = poll_helper.set_all(1)
        driver.wait_for_timeout(5000)
        run_rapid_push_case(
            CASE_ID, PROTOCOL, FILE_FORMAT,
            FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
            pool, driver,
        )
    finally:
        if originals:
            poll_helper.restore_all(originals)
