# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case01
类别：A — Disable 验证
功能：Rapid Logger 处于 Disable 状态时，FTP / SFTP 接收目录不应出现任何文件

【预置条件】
  1. setup_env.py 已在后台运行：FTP / SFTP 服务器已启动，Post Channel 1=FTP / 2=SFTP 已配置
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 FTP / SFTP 接收目录

【测试步骤】
  1. 进入 Data Log → Data Loggers → Rapid Logger，将 Enable 设为 Disable 并保存
  2. 等待 90 秒
  3. 检查 FTP / SFTP 接收目录下是否有文件

【预期结果】
  - FTP 接收目录下不出现任何文件
  - SFTP 接收目录下不出现任何文件
"""
import time
from datalog_page import RapidLoggerPage
from helpers import collect_files

CASE_ID      = "TestCase_AcuHMI_003_04_case01"
CHECK_PROTOS = ["FTP", "SFTP"]


def test_case(pool, driver):
    RapidLoggerPage(driver).disable_rapid_logger()
    time.sleep(90)
    dirs  = [pool[p].data_dir for p in CHECK_PROTOS if p in pool]
    found = collect_files(dirs)
    assert not found, f"[{CASE_ID}] Rapid Logger 已 disable，但 FTP/SFTP 目录发现文件：{found}"
