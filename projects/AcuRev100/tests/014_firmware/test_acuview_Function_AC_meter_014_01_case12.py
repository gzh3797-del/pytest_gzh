r"""Function_AC meter_014_01_case12
用例标题: 修改slave id后进行RTU升级
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: 修改slave id为2, RTU连接到上位机; 进行RTU升级
预期结果: RTU升级成功

生成说明: 暂 skip 存根: 改 SlaveID 后 Acuview 需以 slave=2 重建连接会话(Add Connection 无该表项), 且 016 已实证 SlaveID 变更需专用 runner(2026-07-15 踩坑: spinBox 截断+失联)。待 016 SlaveID runner 重构后一并接入。
"""
from pathlib import Path

import pytest

from comm.ctl_acuview.gui_driver import is_session_locked

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

pytestmark = [
    pytest.mark.skipif(is_session_locked(), reason="会话锁屏/远程断开, 无法驱动 Acuview2 界面"),
]


def test_014_01_case12():
    pytest.skip("需 slave=2 专用连接会话+SlaveID 专用还原 runner(016 重构中), 暂不自动")
