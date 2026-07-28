# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_03_case04
类别：C — 正向推送验证
功能：Logger3 通过 FTP 推送 CSV 文件，验证文件到达及格式正确性

【预置条件】
  1. setup_env.py 已在后台运行：FTP 服务器（端口 2121）已启动，
     Post Channel 1=FTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 FTP 接收目录

【测试步骤】
  1. 进入 Data Loggers 页面，对 Logger3 进行如下配置并保存：
       Enable               = True
       Post Channel         = 1（FTP）
       Log File Format      = CSV
       Log File Length      = 1 minute
       Timestamp Format     = Local Time String
       Log File Name Format = UTC Timestamp
       Log File Name Prefix = meter0_Logger3
       Log Interval         = 1 minute
  2. 等待最长 120 秒，轮询 FTP 接收目录直至出现 .csv 文件
  3. 验证文件扩展名为 .csv
  4. 验证文件名格式：meter0_Logger3 + 14 位时间戳（UTC Timestamp）
  5. 验证文件内容中时间戳列符合 Local Time String 格式
  6. 验证相邻行时间戳间隔约为 1 minute（±10%）
  7. 验证文件数据覆盖时长约为 1 minute（±10%）
  8. 将 Logger3 设为 Disable

【预期结果】
  - FTP 接收目录内收到至少一个 .csv 文件
  - 文件名符合格式：meter0_Logger3<YYYYMMDDHHmmss>.csv
  - 所有时间戳列符合 Local Time String 格式
  - 相邻行间隔与 Log Interval（1 minute）一致
  - 文件数据覆盖时长与 Log File Length（1 minute）一致
"""
from helpers import run_push_case

CASE_ID      = "TestCase_AcuHMI_003_03_case04"
LOGGER_N     = 3
PROTOCOL     = "FTP"
FILE_FORMAT  = "csv"
FILE_LENGTH  = "1 minute"
TS_FMT       = "Local Time String"
NAME_FMT     = "UTC Timestamp"
PREFIX       = "meter0_Logger3"
INTERVAL     = "1 minute"


def test_case(pool, driver):
    run_push_case(
        CASE_ID, LOGGER_N, PROTOCOL, FILE_FORMAT,
        FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
        pool, driver,
    )
