r"""Function_AC meter_014_01_case15
用例标题: boot界面信息查看（升级波特率默认9600）
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: 电表进入boot模式, 查看电表LCD界面信息显示是否正确
预期结果: 信息显示正确(具体显示样式需求侧待确认)

生成说明: MANUAL: 电表本体 LCD boot 界面目视项, 无寄存器可断言(物理观察项不计自动判据); 且预期'如何显示待确认'——需求未冻结。
"""
from pathlib import Path

import pytest

from comm.ctl_acuview.gui_driver import is_session_locked

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

pytestmark = [
    pytest.mark.skipif(is_session_locked(), reason="会话锁屏/远程断开, 无法驱动 Acuview2 界面"),
]


def test_014_01_case15():
    pytest.skip("MANUAL 目视项: 电表LCD boot界面显示, 无寄存器判据且预期样式待需求确认")
