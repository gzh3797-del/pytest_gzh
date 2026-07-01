# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_03_case12
类别：F — 推送中断与恢复验证
功能：Logger3 推送 HTTP/JSON 正常后，禁用 Logger 模拟中断，恢复后验证文件继续产生

【预置条件】
  1. setup_env.py 已在后台运行：HTTP 服务器已启动，Post Channel 3=HTTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 HTTP 接收目录

【测试步骤】
  1. 进入 Data Loggers 页面，对 Logger3 进行如下配置并保存：
       Enable               = True
       Post Channel         = 3（HTTP）
       Log File Format      = JSON
       Log File Length      = 5 minute
       Log File Name Format = Time Interval Format
       Log File Name Prefix = meter2_Logger3
       Log Interval         = 1 minute
  2. 等待最长 450 秒，轮询 HTTP 接收目录直至出现 .json 文件（验证正常推送）
  3. 将 Logger3 设为 Disable（模拟中断）
  4. 等待 90 秒，检查 HTTP 目录文件数量不增加
  5. 重新启用 Logger3（恢复推送），等待新文件出现
  6. 验证恢复后 HTTP 目录有新文件产生

【预期结果】
  2. HTTP 目录在 450s 内收到文件（正常推送）
  3-4. Logger3 禁用期间 HTTP 目录文件数不增加
  5-6. 重新启用后 HTTP 目录出现新文件
"""
import os
import time
from datalog_page import DataLoggerPage, DataLoggerConfig
from datalog_server_verifier import wait_for_files
from helpers import collect_files, PROTO_TO_CHANNEL, LENGTH_TIMEOUT

CASE_ID     = "TestCase_AcuHMI_003_03_case12"
LOGGER_N    = 3
PROTOCOL    = "HTTP"
FILE_FORMAT = "json"
FILE_LENGTH  = "1 minute"
TS_FMT      = "ISO8601 Format"
NAME_FMT    = "Time Interval Format"
PREFIX      = "meter2_Logger3"
INTERVAL    = "1 minute"

_CFG = DataLoggerConfig(
    channel_index=PROTO_TO_CHANNEL[PROTOCOL],
    enabled=True,
    log_file_format=FILE_FORMAT,
    log_file_length=FILE_LENGTH,
    timestamp_format=TS_FMT,
    log_file_name_format=NAME_FMT,
    log_file_name_prefix=PREFIX,
    log_interval=INTERVAL,
)


def test_case(pool, driver):
    dl_page = DataLoggerPage(driver)
    target_dir = os.path.normpath(pool[PROTOCOL].data_dir)
    timeout = LENGTH_TIMEOUT.get(FILE_LENGTH, 450)

    # 阶段 1：正常推送验证
    dl_page.configure_logger(LOGGER_N, _CFG)
    initial_files = wait_for_files([target_dir], timeout=timeout)
    assert initial_files, (
        f"[{CASE_ID}] 阶段1：正常推送超时 {timeout}s 内未收到文件"
    )

    # 阶段 2：禁用（模拟中断），验证文件停止
    dl_page.disable_logger(LOGGER_N)
    count_at_disable = len(collect_files([target_dir], logger_n=LOGGER_N))
    time.sleep(90)
    count_after_wait = len(collect_files([target_dir], logger_n=LOGGER_N))
    assert count_after_wait == count_at_disable, (
        f"[{CASE_ID}] 阶段2：Logger 禁用后仍产生新文件"
        f"（禁用时 {count_at_disable} 个，等待后 {count_after_wait} 个）"
    )

    # 阶段 3：重新启用，验证恢复推送
    dl_page.configure_logger(LOGGER_N, _CFG)
    resumed = wait_for_files([target_dir], min_files=count_at_disable + 1, timeout=timeout)
    assert len(resumed) > count_at_disable, (
        f"[{CASE_ID}] 阶段3：重新启用后 {timeout}s 内未出现新文件"
    )

    dl_page.disable_logger(LOGGER_N)
