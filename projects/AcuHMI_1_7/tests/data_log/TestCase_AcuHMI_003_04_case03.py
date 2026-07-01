# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case03
类别：B — PostChannel=None 验证
功能：Rapid Logger Enable 但 Post Channel=None 时，所有远端目录均不应出现文件

【预置条件】
  1. setup_env.py 已在后台运行，FTP / SFTP / HTTP 服务器均已启动
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空所有协议数据目录

【测试步骤】
  1. 进入 Rapid Logger，配置如下并保存：
       Enable               = True
       Post Channel         = None
       Log File Format      = CSV
       Log File Length      = 1 minute
       Timestamp Format     = Local Time String
       Log File Name Format = UTC Timestamp
       Log File Name Prefix = meter0_RapidLogger
       Log Interval         = 1 minute
  2. 等待 90 秒
  3. 检查 FTP / SFTP / HTTP / HTTPS 全部远端目录
  4. 将 Rapid Logger 设为 Disable

【预期结果】
  - FTP / SFTP / HTTP / HTTPS 所有目录下均不出现任何文件
  - 日志仅保存于设备本地（DataLogManagement）
"""
import time
from datalog_page import RapidLoggerPage
from helpers import collect_files, configure_rapid_logger_none_channel

CASE_ID = "TestCase_AcuHMI_003_04_case03"


def test_case(pool, driver):
    rl_page = RapidLoggerPage(driver)
    configure_rapid_logger_none_channel(
        rl_page,
        file_format="csv",
        file_length="1 minute",
        timestamp_fmt="Local Time String",
        name_fmt="UTC Timestamp",
        prefix="meter0_RapidLogger",
        interval="1 minute",
    )
    time.sleep(90)
    dirs  = [pool[p].data_dir for p in ["FTP", "SFTP", "HTTP", "HTTPS"] if p in pool]
    found = collect_files(dirs)
    rl_page.disable_rapid_logger()
    assert not found, f"[{CASE_ID}] PostChannel=None，但远端目录发现文件：{found}"
