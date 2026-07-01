# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_01_case18
类别：C — 正向推送验证
功能：Logger1 通过 HTTP 推送 JSON 文件，验证文件到达及格式正确性

【预置条件】
  1. setup_env.py 已在后台运行：HTTP 服务器（端口 8080）已启动，
     Post Channel 3=HTTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 HTTP 接收目录

【测试步骤】
  1. 进入 Data Loggers 页面，对 Logger1 进行如下配置并保存：
       Enable               = True
       Post Channel         = 3（HTTP）
       Log File Format      = JSON
       Log File Length      = 10 minute
       Timestamp Format     = ISO8601 Format
       Log File Name Format = Time Interval Format
       Log File Name Prefix = meter2_logger1
       Log Interval         = 1 minute
  2. 等待最长 720 秒，轮询 HTTP 接收目录直至出现 .json 文件
  3. 验证文件扩展名为 .json
  4. 验证文件名格式：meter2_logger1 + 开始时间戳 + 结束时间戳（Time Interval Format）
  5. 验证文件内容中 time/timestamp 字段符合 ISO8601 格式（yyyy-MM-ddTHH:mm:ss）
  6. 验证相邻记录时间戳间隔约为 1 minute（±10%）
  7. 验证文件数据覆盖时长约为 10 minute（±10%）
  8. 将 Logger1 设为 Disable

【预期结果】
  - HTTP 接收目录内收到至少一个 .json 文件
  - 文件名符合 Time Interval Format（meter2_logger1<start><end>.json）
  - 所有记录时间戳符合 ISO8601 格式
  - 相邻记录间隔与 Log Interval（1 minute）一致
  - 文件数据覆盖时长与 Log File Length（10 minute）一致
"""
from helpers import run_push_case

CASE_ID      = "TestCase_AcuHMI_003_01_case18"
LOGGER_N     = 1
PROTOCOL     = "HTTP"
FILE_FORMAT  = "json"
FILE_LENGTH  = "1 minute"
TS_FMT       = "ISO8601 Format"
NAME_FMT     = "Time Interval Format"
PREFIX       = "meter2_logger1"
INTERVAL     = "1 minute"


def test_case(pool, driver):
    run_push_case(
        CASE_ID, LOGGER_N, PROTOCOL, FILE_FORMAT,
        FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
        pool, driver,
    )
