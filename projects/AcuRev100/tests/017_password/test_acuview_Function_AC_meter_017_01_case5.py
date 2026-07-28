r"""Function_AC meter_017_01_case5
用例标题: 上位机设置密码0001、1000、9999，用设置密码可进入对应目录，错误密码不可以
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表
测试步骤: 密码设置界面设 0001/1000/9999, 每次用新密码登录成功; 输错4位密码登录失败。
预期结果: 三个密码均设置成功且能登录; 错误密码提示错误、登录失败。

生成说明(待硬件补/设计): PASSWORD 寄存器(0x1000=4096, R/W 可回读)。可 Modbus 判据: 设密码后回读 4096=新值。
"用新密码登录成功/错误密码失败"属 GUI 权限行为(密码弹窗 runner)。PASSWORD 在 forbid_write('PASSWORD' 子串),
经 allow_write=['PASSWORD'] 放行。⚠️ 用例结束必须还原密码=0(出厂默认), 否则后续 elevate(写死0000)会失效。
🟡 门禁: 密码弹窗登录 runner + OCR 读提示; 且需确认 Acuview 是否有独立"设密码"控件(General Password 字段
是 elevate 输入口, 设新密码入口待真机确认)。
"""
import pytest

pytestmark = pytest.mark.skip(reason="🟡待: 确认Acuview设密码入口控件 + 密码弹窗登录runner + 还原PW=0; 见docstring")


def test_017_01_case5():
    pytest.skip("待硬件确认设密码控件 + 登录runner + Modbus回读4096 + 还原PW; 见docstring")
