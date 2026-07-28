r"""Function_AC meter_017_01_case4
用例标题: 首次连接上位机进入system status，设置时间和clear run time/load time(密码权限管理)
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: System Status 设时间 -> 输错密码(提示错误) -> 输对密码(设置成功) ->
          重连后遍历 clear run time / clear load time(均需输密码)。
预期结果: 错误密码提示错误; 正确密码设置成功; clear run/load time 首次均需密码。

生成说明(待硬件补/设计): 密码弹窗 runner + System Status 的 Time/Run Time Clear/Load Time Clear。
可 Modbus 判据锚点: DEVICE_RUN_TIME(0x1019=4121)/DEVICE_LOAD_TIME(0x101B=4123) 是否随密码结果被清;
时间设置可读 RTC 寄存器验证。clear run/load time 属重置类(与 011 计时 case10/11 的 4404/4405 相关, 非强破坏)。
🟡 门禁: 密码弹窗 runner(三态)。见 017_01_case1。
"""
import pytest

pytestmark = pytest.mark.skip(reason="🟡待: 补密码弹窗runner(三态); 同017_01_case1")


def test_017_01_case4():
    pytest.skip("待硬件补密码弹窗runner(设时间/clear run·load time需密码) + RTC回读; 见docstring")
