# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_08_case03
类别：A — Post Historical Data CSV 推送（HTTP，ISO8601 Format + Time Interval Format）
功能：配置 Post Channel 3（HTTP），CSV 格式，Log File Length=1 hour，Log Interval=10 min，
      点击 Post 等待进度到 100%，检查 HTTP 服务器文件的格式与内容

【预置条件】
  1. setup_env.py 已在后台运行，HTTP 服务器已启动
  2. conftest.py driver fixture 已完成网关 Web 登录及 Post Channel 3=HTTP 配置
  3. conftest.py clear_dirs fixture 已在本用例执行前清空 HTTP 数据目录

【测试步骤】
  1. 进入 Post Historical Data 页面
  2. Post Channel = Post Channel 3（HTTP）
  3. Device = 第一台可用设备
  4. Log File Format = CSV
  5. Timestamp Format = ISO8601 Format
  6. Log File Name Format = Time Interval Format
  7. Log File Name Prefix = meter2_logger1
  8. Log File Length = 1 hour
  9. Log Interval = 10 minutes
  10. 点击 Post，等待进度框到 100%
  11. 检查 HTTP 目录收到的文件

【预期结果】
  - HTTP 目录下出现 .csv 文件
  - 文件名格式符合 Time Interval Format 规则，前缀为 meter2_logger1
  - CSV 中 Timestamp 列为 ISO8601 格式
  - Log Interval 约 600 s，Log File Length 约 1 hour
"""
from helpers import run_post_historical_case

CASE_ID = "TestCase_AcuHMI_003_08_case03"


def test_case(pool, driver):
    run_post_historical_case(
        case_id=CASE_ID,
        protocol="HTTP",
        file_format="csv",
        file_length="1 hour",
        timestamp_fmt="ISO8601 Format",
        name_fmt="Time Interval Format",
        prefix="meter2_logger1",
        interval="10 minutes",
        pool=pool,
        driver=driver,
    )
