# AcuRev-1320 — 三相电表测试项目

## 测试设计必读（固定结构速查区）

> strategy-design / testcase-design / coverage-check 的「知识库提取表」以本区块为权威来源，正文为详情补充。正文变更涉及下列字段时必须同步更新本区块。

| 字段 | 内容 |
|------|------|
| 设备清单 | AcuRev-1320，两型号必须分开覆盖：**1321**（精简：仅 200ms 聚合、无 PQ/录波、仅 200ms 谐波）/ **1322**（全功能：5 种聚合、PQ Event+Waveform、4 种聚合谐波）；扩展：ACM-41-WEB2（USB 虚拟串口） |
| 协议与端口 | Modbus TCP / Modbus RTU（9600~115200）/ BACnet MS/TP / BACnet IP / HTTPS Web 页面（强制 HTTPS，自签名证书，功能与上位机对齐）/ USB WEB2（Modbus RTU，FC 0x66，720 寄存器/帧） |
| 接线方式 | 6 种：1E2W / 2E3W 1Phase / 2E3W Network / 2E3W Delta / 3E4W Y / 3E4W Delta；**特殊点：2E3W 1Phase 用 A+C（非 A+B）；2E3W Delta 仅测 Ia/Ic（不含 Ib）** |
| 容量与边界 | Independent Input Channel 4 路 + 2 Summation；Custom Read ≤125 寄存器；PQ 事件 ≤200 条（含 3V+4I 录波）；告警 20 组（2 参数 AND/OR）≤200 条日志；Audit/SYS Log 各 5000 条；Data Post 单文件 ≤100KB；RTC 断电 1 周 |
| 共存/联动约束 | **AcuCloud 与 DataLog 互斥**（AcuCloud 5min 推送 CSV）；Audit Log 铅封保护不可清除；铅封时禁止固件更新；Dual Source Energy 电网/光伏由 DI 或通信切换；闪变 Pinst/Pst/Plt 三相各自输出 |
| 安全与默认状态 | 强制 HTTPS；标准铅封 + 非标准铅封（可配置保护范围） |
| 高频缺陷模式 | → 详见 bugs/INDEX.md |
| 测试环境要点 | 接线实测相/CT 速查 → knowledge/shared/wiring_reference.md；接线检测算法判据 → web2 项目 wiring_check_v1.05.md |

## 产品背景

AcuRev-1320 是新一代导轨安装三相电表，面向工业/商业分表计量，具备高精度测量与多种通信功能。通过 USB 扩展口与 ACM-41-WEB2 模块配合，可扩展北向协议。

**需求文档：**
- 软件需求规格 v1.00：`requirements/raw/AcuRev-1320软件需求规格说明书 v1.00 20260422.docx` → 摘要：`requirements/summaries/v1.00_20260422.md`
- 时间与时区配置：`requirements/raw/AcuRev-1320时间与时区配置.docx`
- 型号区别显示需求：`requirements/raw/AcuRev-1320型号区别的显示需求.xlsx`
- 独立模式参数显示：`requirements/raw/1320独立模式不同接线方式参数显示 v1.00 20260427.xlsx`
- 非独立模式参数显示：`requirements/raw/1320非独立模式不同接线方式参数显示 v1.00 20260427.xlsx`

**Bug 追踪：**
- 精简索引：`bugs/INDEX.md`

---

## 功能模块

| 模块 | 说明 |
|------|------|
| 交流电参数测量 | 频率/电压/电流/功率/电能/谐波/需量/TOU，精度 0.2% |
| 电压闪变（Flicker） | Pinst/Pst/Plt，三相各自输出 |
| Independent Input Channel | 4 路电流通道 + 2 个 Summation，可独立配置计算电压 |
| Dual Source Energy | 电网/光伏双源电能分开记录，DI 或通信控制切换 |
| Custom Read | 用户配置最多 125 个寄存器，一次性读取 |
| 电能质量（PQ） | Sag/Swell/Interruption/Current Swell；最多 200 条；含录波（3V+4I） |
| 告警 | 20 组，2 参数 AND/OR，最多 200 条告警日志 |
| 记录与统计 | Max/Min / SOE Log / Data Log（3 datalogger + 1 trendlogger） |
| Audit Log / SYS Log | 各 5000 条；Audit Log 铅封保护不可清除 |
| AcuCloud | 5 分钟推送，CSV，与 DataLog 互斥 |
| Data Post | Datalog 数据推送，HTTP/HTTPS，文件 ≤100KB |
| Input & Output | DI/DO/RO + 脉冲计数/能量脉冲/告警指示 |
| 时间功能 | RTC 断电保持 1 周；SNTP 同步；DST |
| 通信协议 | Modbus TCP / RTU（9600~115200）/ BACnet MS/TP / BACnet IP |
| USB WEB2 | USB 虚拟串口 Modbus RTU，FC 0x66（720 寄存器/帧） |
| HTTPs Web 页面 | 强制 HTTPS，自签名证书；功能与上位机对齐 |
| 铅封 | 标准铅封 + 非标准铅封（可配置范围） |
| 固件更新 | RS485 / 以太网，铅封时禁止 |

---

## 型号差异（1321 vs 1322）

| 特性 | 1321 | 1322 |
|------|------|------|
| 聚合数据 | 仅 200ms（10/12 cycle） | 5 种聚合（1C/10/12C/150/180C/10min/2hr） |
| PQ Event / Waveform | 无 | 有 |
| 谐波 | 仅 200ms | 4 种聚合谐波 |

---

## 接线方式

1E2W / 2E3W 1Phase / 2E3W Network / 2E3W Delta / 3E4W Y / 3E4W Delta

各接线方式实测相/CT 速查 → `knowledge/shared/wiring_reference.md`
完整接线检测算法规格（缺失/反接/相移判据） → `knowledge/gateway/AcuRev4100WEB2/requirements/summaries/wiring_check_v1.05.md`

> 1320 特殊点：2E3W 1Phase 用 A+C（非 A+B）；2E3W Delta 仅测 Ia/Ic（不含 Ib）。

---

## 关联项目

- **ACM-41-WEB2**：通信扩展模块（WEB2），通过 USB 与本表通信 → `knowledge/gateway/AcuRev4100WEB2/context.md`
