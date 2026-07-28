# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_08_case04
类别：A — Post Historical Data CSV 推送（FTP，Log File Length=1 day）
功能：配置 Post Channel 1（FTP），CSV 格式，Log File Length=1 day，Log Interval=1 min，
      点击 Post 等待进度到 100%，检查 FTP 服务器文件的格式与内容

【预置条件】
  1. setup_env.py 已在后台运行，FTP 服务器已启动
  2. conftest.py driver fixture 已完成网关 Web 登录及 Post Channel 1=FTP 配置
  3. conftest.py clear_dirs fixture 已在本用例执行前清空 FTP 数据目录

【测试步骤】
  1. 进入 Post Historical Data 页面
  2. Post Channel = Post Channel 1（FTP）
  3. Device = 第一台可用设备
  4. Log File Format = CSV
  5. Timestamp Format = Local Time String
  6. Log File Name Format = UTC Timestamp
  7. Log File Name Prefix = meter0_logger1
  8. Log File Length = 1 day
  9. Log Interval = 1 minute
  10. 点击 Post，等待进度框到 100%
  11. 检查 FTP 目录收到的文件

【预期结果】
  - FTP 目录下出现 .csv 文件
  - 文件名格式符合 UTC Timestamp 规则，前缀为 meter0_logger1
  - CSV 中 Timestamp 列为 Local Time String 格式
  - Log Interval 约 60 s，Log File Length 约 1 day
"""
from helpers import run_post_historical_case

CASE_ID = "TestCase_AcuHMI_003_08_case04"


def test_case(pool, driver):
    run_post_historical_case(
        case_id=CASE_ID,
        protocol="FTP",
        file_format="csv",
        file_length="1 day",
        timestamp_fmt="Local Time String",
        name_fmt="UTC Timestamp",
        prefix="meter0_logger1",
        interval="1 minute",
        pool=pool,
        driver=driver,
    )
