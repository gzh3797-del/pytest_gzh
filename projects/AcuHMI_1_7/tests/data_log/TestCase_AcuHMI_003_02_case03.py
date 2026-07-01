# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_02_case03
类别：B — PostChannel=None 验证
功能：Logger2 Enable 但 Post Channel=None 时，所有远端目录均不应出现文件

【预置条件】
  1. setup_env.py 已在后台运行，FTP / SFTP / HTTP 服务器均已启动
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已在本用例执行前清空所有协议数据目录

【测试步骤】
  1. 进入 Data Loggers 页面，对 Logger2 进行如下配置并保存：
       Enable               = True
       Post Channel         = None（不推送到任何远端服务器）
       Log File Format      = CSV
       Log File Length      = 1 minute
       Timestamp Format     = Local Time String
       Log File Name Format = UTC Timestamp
       Log File Name Prefix = meter0_Logger2
       Log Interval         = 1 minute
  2. 等待 90 秒，覆盖至少一个文件推送窗口
  3. 检查 FTP / SFTP / HTTP / HTTPS 全部协议接收目录
  4. 将 Logger2 设为 Disable

【预期结果】
  - FTP / SFTP / HTTP / HTTPS 所有目录下均不出现任何 .csv / .json 文件
"""
import time
from helpers import collect_files, configure_logger_none_channel
from datalog_page import DataLoggerPage

CASE_ID  = "TestCase_AcuHMI_003_02_case03"
LOGGER_N = 2


def test_case(pool, driver):
    dl_page = DataLoggerPage(driver)
    configure_logger_none_channel(
        dl_page, LOGGER_N,
        file_format="csv",
        file_length="1 minute",
        timestamp_fmt="Local Time String",
        name_fmt="UTC Timestamp",
        prefix="meter0_Logger2",
        interval="1 minute",
    )
    time.sleep(90)
    dirs  = [pool[p].data_dir for p in ["FTP", "SFTP", "HTTP", "HTTPS"] if p in pool]
    found = collect_files(dirs, logger_n=LOGGER_N)
    dl_page.disable_logger(LOGGER_N)
    assert not found, f"[{CASE_ID}] PostChannel=None，但发现文件：{found}"
