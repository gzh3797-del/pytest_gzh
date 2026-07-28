r"""Function_AC meter_014_01_case5
用例标题: 升级过程中断开连接，待通讯正常后能继续升级
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表 / 4、铅封解锁
测试步骤: 升级到40%关上位机→重开续升(成功)→再升级至clearing时给电表掉电→上电续升(成功)
预期结果: 关闭瞬间提示升级失败; 重开能升级成功; clearing掉电进Bootloader; 上电后能升级成功

生成说明: 暂 skip 存根: 需 OCR 盯进度到 40% 精准杀 Acuview 进程 + clearing 时刻源掉电(自供电表)两段时序配合, 属专项联调(1320 同类 case4_01 亦为 manual)。引擎就绪后转自动。
"""
from pathlib import Path

import pytest

from comm.ctl_acuview.gui_driver import is_session_locked

PROJECT_ROOT = Path(__file__).resolve().parents[2]
TEST_CONFIG = str(PROJECT_ROOT / "config.yaml")

pytestmark = [
    pytest.mark.skipif(is_session_locked(), reason="会话锁屏/远程断开, 无法驱动 Acuview2 界面"),
]


def test_014_01_case5():
    pytest.skip("中断续升需'40%杀进程+clearing掉电'两段精准时序, 待专项联调(参照1320 case4_01 manual先例)")
