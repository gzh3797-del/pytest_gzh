# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case10
类别：B — PostChannel=None 验证（JSON 格式，1 second 间隔）
功能：Rapid Logger Enable 但 Post Channel=None 时，所有远端目录均不应出现文件

【预置条件】
  1. setup_env.py 已在后台运行，FTP / SFTP / HTTP 服务器均已启动
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空所有协议数据目录
  4. 所有 Physical Devices 的 Poll Interval 须 ≤ 1s（本用例自动设置并还原）

【测试步骤】
  1. 将所有 Physical Devices 的 Poll Interval 改为 1s
  2. 配置 Rapid Logger：
       Enable=True  PostChannel=None  Format=JSON  Length=1 minute
       NameFormat=UTC Timestamp  Prefix=meter0_RapidLogger  Interval=1 second
  3. 等待 90 秒
  4. 检查所有远端目录
  5. 将 Rapid Logger 设为 Disable
  6. 恢复所有 Physical Devices 的 Poll Interval 为原始值

【预期结果】
  - FTP / SFTP / HTTP / HTTPS 所有目录下均不出现任何文件
  - 日志仅保存于设备本地（DataLogManagement）
"""
import time
from datalog_page import RapidLoggerPage, PhysicalDevicePollHelper
from helpers import collect_files, configure_rapid_logger_none_channel

CASE_ID = "TestCase_AcuHMI_003_04_case10"


def test_case(pool, driver):
    poll_helper = PhysicalDevicePollHelper(driver)
    originals = {}
    try:
        originals = poll_helper.set_all(1)
        driver.wait_for_timeout(5000)
        rl_page = RapidLoggerPage(driver)
        configure_rapid_logger_none_channel(
            rl_page,
            file_format="json",
            file_length="1 minute",
            timestamp_fmt="UTC Seconds",
            name_fmt="UTC Timestamp",
            prefix="meter0_RapidLogger",
            interval="1 second",
        )
        time.sleep(90)
        dirs  = [pool[p].data_dir for p in ["FTP", "SFTP", "HTTP", "HTTPS"] if p in pool]
        found = collect_files(dirs)
        rl_page.disable_rapid_logger()
        assert not found, f"[{CASE_ID}] PostChannel=None，但远端目录发现文件：{found}"
    finally:
        if originals:
            poll_helper.restore_all(originals)
