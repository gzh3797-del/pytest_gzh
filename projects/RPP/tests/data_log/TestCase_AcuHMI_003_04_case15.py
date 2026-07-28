# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_04_case15
类别：C — 正向推送验证（JSON，5 seconds 间隔，含断连恢复观察）
功能：Rapid Logger 通过 HTTP/HTTPS 推送 JSON，验证日志存储路径、数据格式、推送频率及数据准确性；
     同时观察短暂断开下游设备连接（13 s 内）和重新上电（20 s 后）后的恢复行为。

【预置条件】
  1. HTTP 服务器已启动，Post Channel 3=HTTP 已配置并 Enabled
  2. conftest.py driver fixture 已完成网关 Web 登录
  3. conftest.py clear_dirs fixture 已清空 HTTP 接收目录

【测试步骤】
  1. 配置 Rapid Logger：
       Enable=True  PostChannel=3(HTTP/HTTPS)  Format=JSON  Length=1 minute
       NameFormat=Time Interval Format  Prefix=meter2_RapidLogger  Interval=5 seconds
  2. 等待最长 120s，检查远端 HTTP/HTTPS 目录是否有 Rapid Logger 日志文件
  3. [手动] 断开下游 Modbus 设备网络连接，13 秒以内再检查目录（
     期望：目录中仍有文件，且时间戳间隔保持正常）
  4. [手动] 重新连接网络设备，20 秒后上电，检查目录是否有新文件恢复（
     期望：目录中出现新文件，时间戳连续）
  5. 将 Rapid Logger 设为 Disable

【预期结果】
  步骤 2：远端目录中出现 Rapid Logger JSON 文件，
          文件内数据格式、行数与配置及实际设备采集值一致
  步骤 3：断连 13 s 内，远端目录中的日志文件时间戳间隔列一致（自动化只验证步骤 2）
  步骤 4：重连后恢复推送新文件（自动化只验证步骤 2）

注意：步骤 3/4 需要物理硬件操作（断/接 Modbus 设备网络），
      本自动化脚本仅完成步骤 1-2 的验证。
"""
from helpers import run_rapid_push_case

CASE_ID     = "TestCase_AcuHMI_003_04_case15"
PROTOCOL    = "HTTP"
FILE_FORMAT = "json"
FILE_LENGTH = "1 minute"
TS_FMT      = "UTC Seconds"
NAME_FMT    = "Time Interval Format"
PREFIX      = "meter2_RapidLogger"
INTERVAL    = "5 seconds"


def test_case(pool, driver):
    run_rapid_push_case(
        CASE_ID, PROTOCOL, FILE_FORMAT,
        FILE_LENGTH, TS_FMT, NAME_FMT, PREFIX, INTERVAL,
        pool, driver, full_verify=True,
    )
