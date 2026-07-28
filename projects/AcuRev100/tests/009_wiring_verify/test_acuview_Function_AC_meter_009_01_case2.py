r"""Function_AC meter_009_01_case2
用例标题: 1E2W 测量5分钟 切2E3W1P 再切回1E2W, 检查HMI/上位机能量全置0, 电压电流测量正确
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: 配 1E2W(Va=120V,Ia=80A,0°)测5min -> Current&Wiring 改 Service 下拉切 2E3W1P
          (弹"是否清除能量"选清除) -> 再切回 1E2W -> 检查能量=0 且 V/I 测量正确。
预期结果: 切接线后能量全部清 0, 电压电流测量正确。

生成说明(待硬件补实现): 复合流程=改 Wiring Service 下拉(comboBox@0x1042=4162, OCR)触发
"切接线清能量"弹窗 -> _confirm_dialogs 确认清除 -> Modbus 回读能量=0 + V/I 区间。
🔴🟡 三重门禁: 红线(能量清零, 已授权需在场放行) + OCR(下拉+弹窗) + 需新增"改下拉触发弹窗"复合流程
(引擎现有 run_button_action_case 仅覆盖固定按钮; 切接线弹窗流程明天在硬件上补一个 runner)。
V/I 测量部分当前板 ADC 损坏为已知 FAIL。
"""
import pytest

# 2026-07-15 结论: 上位机(当前 Acuview debug 版)尚未实现"切接线时清/不清能量"的选择功能,
#   无法自动化本用例。需求已确认存在该功能, 待 Acuview 正式版发布后再补 009 切接线弹窗复合 runner。
pytestmark = pytest.mark.skip(reason="上位机未实现切接线清/不清能量选择功能(需求已确认), 待正式版发布")


def test_009_01_case2():
    pytest.skip("上位机切接线清能量选择功能未实现, 待正式版; 见docstring")
