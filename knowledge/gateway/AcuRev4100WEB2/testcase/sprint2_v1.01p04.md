# AcuRev-4100-WEB2 Sprint 2 测试用例摘要

**文件**：AcuRev-4100-WEB2 Sprint 2_测试用例_v1.02_20260509.xlsx
**软件版本**：v1.01p04
**测试日期**：2026.05.26 - 2026.05.27
**版本目标**：新增需求 + 影响模块

## 版本测试概况

> 版本测试第二天，测试用例执行 100%。发现 14 个 BUG。

| 指标 | 数值 |
|------|------|
| 本版本执行用例数 | 168 |
| Passed | 158 |
| Failed | 9 |
| Unavailable | 1 |
| 执行率 | 93.3%（168/180） |
| 全量用例总数 | 1202 |

## 模块执行情况（本版本在测）

| 模块 | 用例总数 | 执行数 | Passed | Failed | Unavailable |
|------|--------|--------|--------|--------|-------------|
| 设备数据协议转换 | 133 | 121 | 119 | 1 | 1 |
| Service | 11 | 11 | 10 | 1 | 0 |
| Wiring Check | 36 | 36 | 29 | 7 | 0 |
| **合计** | **180** | **168** | **158** | **9** | **1** |

## 本版本发现问题（14 条）

| 问题单号 | 模块 | 摘要 |
|---------|------|------|
| A4WS-225 | Wiring Check | Show Only Issues 仍显示 Phase Order ABC |
| A4WS-224 | Wiring Check | 报告中 Measurement 乱码，Phase Order 显示 unKnow，Device Name 显示 Device |
| A4WS-223 | Wiring Check | 铅封状态下 Confirm Nominal Voltage 弹窗额定电压下发成功（预期应失败） |
| A4WS-222 | AWS IoT | 配置错误证书后 Test Connection 仍返回成功 |
| A4WS-221 | SNMP | 只选择 AcuRev4100 上报，SNMP Walk 实际重复上报两次 |
| A4WS-220 | Wiring Check | 1E2W 接线检查 Phase Order 显示 Unknown（预期不显示） |
| A4WS-219 | Wiring Check | DELTA 接线下电压测量值显示相电压（预期线电压） |
| A4WS-218 | Troubleshooting | 选择 4100 主表修改 CT Model 后 Save 未成功 |
| A4WS-217 | About | Installation Record 选择 IOM 设备点击新建页面报错 |
| A4WS-216 | Wiring Check | Va=0&Ia=1 时 Wiring Status 显示 PhaseError（预期 Phase Shift） |
| A4WS-215 | AcuCloud | AcuCloud 平台 PT Installation Table 与 WEB2 显示不一致 |
| A4WS-213 | Wiring Check | 接线检查页面 UserChannel 名显示不全 |
| A4WS-212 | MQTT | 设备参数导入导出后 Test MQTT 测试失败 |
| A4WS-211 | BACnet/IP | BACnet/IP 设备参数配置 Save 概率报错 |

## 问题单回归情况（4 条，全部通过）

| 问题单号 | 来源版本 | 摘要 | 结论 |
|---------|---------|------|------|
| A4WS-206 | v1.01p03 | SNMP Port 异常提示语错误 | 通过 |
| A4WS-207 | v1.01p03 | 设备列表主表 Interface 显示 "RS485 1" | 通过 |
| A4WS-208 | v1.01p03 | 设备列表主表 DeviceName 显示 "AcuRev41001" | 通过 |
| A4WS-209 | v1.01p03 | 虚拟设备公式不能选择不同 4100 的同一参数 | 通过 |

## 全量用例级别分布

| 级别 | 数量 |
|------|------|
| LV0 | 90 |
| LV1 | 371 |
| LV2 | 491 |
| LV3 | 220 |
| LV4 | 30 |
| **合计** | **1202** |

## 全量模块用例分布

| 模块 | 用例数 |
|------|--------|
| 接入设备日志管理 | 195 |
| 设备数据协议转换 | 188 |
| 多用户管理 | 148 |
| 系统设置（System Setting） | 113 |
| AcuREV4100参数配置 | 101 |
| 设备管理（Management） | 62 |
| AcuIOM-AIO参数配置 | 60 |
| 系统诊断（Diagnostics） | 58 |
| About | 51 |
| 软件升级 | 46 |
| Wiring Check | 36 |
| 接入设备数据采集 | 24 |
| Emergency Mode | 20 |
| 配置管理 | 20 |
| AcuIOM-DIO参数配置 | 17 |
| Advanced Commission | 16 |
| Service | 16 |
| Device Connection | 13 |
| 性能测试 | 10 |
| 稳定性测试 | 5 |
| 安全管理 | 3 |
| **合计** | **1202** |
