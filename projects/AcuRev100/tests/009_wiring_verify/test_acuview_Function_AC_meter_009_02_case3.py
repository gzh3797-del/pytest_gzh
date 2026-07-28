r"""Function_AC meter_009_02_case3
用例标题: 2E3W1P 测5min 切1E2W(能量清0且数据同1E2W一致) 再切回2E3W1P 验证系统功率=A+C
预置条件: 1、Acuview2上位机 / 2、RTU串口线 / 3、AcuRev-100电表(铅封解锁)
测试步骤: 配 2E3W1P 测5min -> 改 Service 下拉切 1E2W(清能量) -> 检查能量=0、数据同 1E2W、
          多出参数显示"-" -> 切回 2E3W1P(Va/Vc=120V,Ia=50A,Ic=150A,0°) -> Iavg=100A、Psys=Pa+Pc。
预期结果: 切接线能量清 0; 多出参数显示"-"(MANUAL); 2E3W1P 下 Iavg=100A、系统功率=A+C 相之和。

生成说明(待硬件补实现): 同 009_01_case2——改 Wiring Service 下拉(4162,OCR)触发清能量弹窗确认 +
Modbus 回读能量=0 + 2E3W1P 下 Iavg/系统功率区间。"-"显示项属 GUI 目视 MANUAL。
🔴🟡 门禁: 红线清能量(在场放行) + OCR + 切接线弹窗复合流程(硬件补)。V/I/P 测量部分 ADC 损坏已知 FAIL。
"""
import pytest

# 2026-07-15 结论: 上位机(当前 Acuview debug 版)尚未实现"切接线时清/不清能量"的选择功能,
#   无法自动化本用例。需求已确认存在该功能, 待 Acuview 正式版发布后再补。
pytestmark = pytest.mark.skip(reason="上位机未实现切接线清/不清能量选择功能(需求已确认), 待正式版发布")


def test_009_02_case3():
    pytest.skip("上位机切接线清能量选择功能未实现, 待正式版; 见docstring")
