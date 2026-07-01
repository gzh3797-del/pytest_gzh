# -*- coding: utf-8 -*-
"""
用例 ID：TestCase_AcuHMI_003_08_case12
类别：D — Post Historical Data 错误验证（Log File Name Prefix 超过 20 字符）
功能：Log File Name Prefix 填入超过 20 字符的字符串，预期 Post 操作报错

【预置条件】
  1. setup_env.py 已在后台运行，HTTP 服务器已启动
  2. conftest.py driver fixture 已完成网关 Web 登录及 Post Channel 3=HTTP 配置
  3. conftest.py clear_dirs fixture 已在本用例执行前清空 HTTP 数据目录

【测试步骤】
  1. 进入 Post Historical Data 页面
  2. Post Channel = Post Channel 3（HTTP）
  3. Device = 第一台可用设备
  4. Log File Format = JSON
  5. Log File Name Format = Time Interval Format
  6. Log File Name Prefix = meter2_logger101234567890（超过 20 字符）
  7. Log File Length = 1 hour
  8. Log Interval = 10 minutes
  9. 点击 Post

【预期结果】
  - 出现错误提示（Prefix 超出长度限制）
  - HTTP 目录下不出现任何文件
"""
from helpers import run_post_historical_case

CASE_ID = "TestCase_AcuHMI_003_08_case12"

# 超过 20 字符的 prefix（用例原文 "meter2_logger101234567890"，共 25 字符）
LONG_PREFIX = "meter2_logger101234567890"


def test_case(pool, driver):
    run_post_historical_case(
        case_id=CASE_ID,
        protocol="HTTP",
        file_format="json",
        file_length="1 hour",
        timestamp_fmt="UTC Seconds",
        name_fmt="Time Interval Format",
        prefix=LONG_PREFIX,
        interval="10 minutes",
        pool=pool,
        driver=driver,
        expect_error=True,
    )
