# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_02_case01
类别：A — Disable 验证
功能：Logger2 处于 Disable 状态时，FTP / SFTP 接收目录不应出现任何文件

【预置条件】
  1. setup_env.py 已在后台运行：FTP / SFTP / HTTP 服务器已启动，
     Post Channel 1=FTP / 2=SFTP / 3=HTTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录

【测试步骤】
  1. 进入 Data Loggers 页面，将 Logger2 设为 Disable 并保存
  2. 等待 90 秒，覆盖至少一个文件推送窗口
  3. 检查 FTP / SFTP 协议数据接收目录下的 .csv / .json 文件

【预期结果】
  - FTP 接收目录下不出现任何文件
  - SFTP 接收目录下不出现任何文件
"""
import time
from helpers import collect_files
from datalog_page import DataLoggerPage

CASE_ID      = "TestCase_AcuHMI_003_02_case01"
LOGGER_N     = 2
CHECK_PROTOS = ["FTP", "SFTP"]


def test_case(pool, driver):
    DataLoggerPage(driver).disable_logger(LOGGER_N)
    time.sleep(90)
    dirs  = [pool[p].data_dir for p in CHECK_PROTOS if p in pool]
    found = collect_files(dirs, logger_n=LOGGER_N)
    assert not found, f"[{CASE_ID}] Logger{LOGGER_N} 已 disable，但发现文件：{found}"
