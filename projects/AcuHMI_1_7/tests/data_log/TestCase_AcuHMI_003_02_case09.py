# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_02_case09
类别：C — 正向推送验证
功能：Logger2 通过 SFTP 推送 JSON 文件，验证文件到达及格式正确性

【预置条件】
  1. setup_env.py 已在后台运行：SFTP 服务器（端口 2222）已启动，
     Post Channel 2=SFTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 SFTP 接收目录

【测试步骤】
  1. 进入 Data Loggers 页面，对 Logger2 进行如下配置并保存：
       Enable               = True
       Post Channel         = 2（SFTP）
       Log File Format      = JSON
       Log File Length      = 5 minute
       Log File Name Format = Time Interval Format
       Log File Name Prefix = meter1_Logger2
       Log Interval         = 1 minute
  2. 等待最长 450 秒，轮询 SFTP 接收目录直至出现 .json 文件
  3. 验证文件扩展名为 .json
  4. 验证文件名格式：meter1_Logger2 + 开始时间戳 + 结束时间戳（Time Interval Format）
  5. 验证相邻行时间戳间隔约为 1 minute（±10%）
  6. 验证文件数据覆盖时长约为 5 minute（±10%）
  7. 将 Logger2 设为 Disable

【预期结果】
  - SFTP 接收目录内收到至少一个 .json 文件
  - 文件名符合 Time Interval Format（meter1_Logger2<start><end>.json）
  - 相邻行间隔与 Log Interval（1 minute）一致
  - 文件数据覆盖时长与 Log File Length（5 minute）一致
"""
from helpers import run_push_case

CASE_ID     = "TestCase_AcuHMI_003_02_case09"
LOGGER_N    = 2
PROTOCOL    = "SFTP"
FILE_FORMAT = "json"
FILE_LENGTH  = "1 minute"
TS_FMT      = "UTC Seconds"
NAME_FMT    = "Time Interval Format"
PREFIX      = "meter1_Logger2"
INTERVAL    = "1 minute"


def test_case(pool, driver):
    run_push_case(
        CASE_ID, LOGGER_N, PROTOCOL, FILE_FORMAT,
        FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
        pool, driver,
    )
