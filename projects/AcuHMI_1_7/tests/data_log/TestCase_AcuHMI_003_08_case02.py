# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_08_case02
类别：A — Post Historical Data CSV 推送（SFTP，UTC Seconds + Time Interval Format）
功能：配置 Post Channel 2（SFTP），CSV 格式，Log File Length=30 min，Log Interval=5 min，
      点击 Post 等待进度到 100%，检查 SFTP 服务器文件的格式与内容

【预置条件】
  1. setup_env.py 已在后台运行，SFTP 服务器已启动
  2. conftest.py driver fixture 已完成网关 Web 登录及 Post Channel 2=SFTP 配置
  3. conftest.py clear_dirs fixture 已在本用例执行前清空 SFTP 数据目录

【测试步骤】
  1. 进入 Post Historical Data 页面
  2. Post Channel = Post Channel 2（SFTP）
  3. Device = 第一台可用设备
  4. Log File Format = CSV
  5. Timestamp Format = UTC Seconds
  6. Log File Name Format = Time Interval Format
  7. Log File Name Prefix = meter1_logger1
  8. Log File Length = 30 minutes
  9. Log Interval = 5 minutes
  10. 点击 Post，等待进度框到 100%
  11. 检查 SFTP 目录收到的文件

【预期结果】
  - SFTP 目录下出现 .csv 文件
  - 文件名格式符合 Time Interval Format 规则，前缀为 meter1_logger1
  - CSV 中 Timestamp 列为 UTC Seconds（Unix 时间戳）格式
  - Log Interval 约 300 s，Log File Length 约 30 min
"""
from helpers import run_post_historical_case

CASE_ID = "TestCase_AcuHMI_003_08_case02"


def test_case(pool, driver):
    run_post_historical_case(
        case_id=CASE_ID,
        protocol="SFTP",
        file_format="csv",
        file_length="30 minutes",
        timestamp_fmt="UTC Seconds",
        name_fmt="Time Interval Format",
        prefix="meter1_logger1",
        interval="5 minutes",
        pool=pool,
        driver=driver,
    )
