r"""Function_AC meter_017_01_case3
用例标题: 首次连接上位机进入system status，factory reset(密码权限管理)
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: System Status 点 Factory Reset -> 输错密码(提示错误, 不执行) -> 输对密码(恢复出厂成功)。
预期结果: 错误密码提示错误且不恢复出厂; 正确密码恢复出厂成功。

生成说明(待硬件补/设计): 密码弹窗 runner + Factory Reset 按钮。
🔴🟡 双门禁: ①红线-恢复出厂(用户本轮未授权自动执行, 见 015_01_case3) ②密码弹窗 runner(输错不执行/输对执行)。
错误密码分支可 Modbus 判据(回读默认值*未*被重置); 正确分支=恢复出厂(需授权+在场)。
"""
import pytest

pytestmark = pytest.mark.skip(reason="🔴🟡待: 红线恢复出厂(未授权)+密码弹窗runner; 见015_01_case3/017_01_case1")


def test_017_01_case3():
    pytest.skip("待: 红线恢复出厂授权+在场 + 密码弹窗runner; 见docstring")
