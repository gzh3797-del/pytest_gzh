# AcuRev-4100 — 多回路电表测试项目

## 测试设计必读（固定结构速查区）

> strategy-design / testcase-design / coverage-check 的「知识库提取表」以本区块为权威来源，正文为详情补充。正文变更涉及下列字段时必须同步更新本区块。

| 字段 | 内容 |
|------|------|
| 设备清单 | 单设备项目：AcuRev-4100 本体（24 路电流输入、12 个用户通道、3 相电压）；通信扩展：ACM-41-WEB2（→ knowledge/gateway/AcuRev4100WEB2/context.md） |
| 协议与端口 | Modbus TCP（≤4 并发客户端，端口非标准，见 config.py）；Modbus RTU（RS485，1200~115200）；BACnet MS/TP（与 Modbus RTU 共用 RS485，分时使用） |
| 接线方式 | 5 种：1E2W / 2E3W 1Phase（12 用户）/ 2E3W Network（12）/ 3E4W Y（8，最常用）/ 2E3W Delta（12） |
| 容量与边界 | DI×4（脉冲计数 ≤50Hz，v1.02 起由 100Hz 修订）；DO×8（能量脉冲/告警）；RO×2（Latch/Pulse）；AI/AO 0-20mA/0-10V；频率 42.5~69Hz；TOU 4 费率/12 季节/14 时刻表；电流输入 mV（333mV/RCT，0~400mV）或 mA（80/100mA，0~120mA） |
| 共存/联动约束 | BACnet MS/TP 与 Modbus RTU 分时使用（同一 RS485 主口） |
| 安全与默认状态 | 无专项（见需求摘要 v1.03） |
| 高频缺陷模式 | → 详见 bugs/INDEX.md |
| 测试环境要点 | 精度基准：电压/电流 0.1%、功率 0.2%、有功电能 IEC 62053-22 Class 0.2s（update 100ms）、谐波 1%（2~31 次，可选 2~63 次）、频率 0.002Hz |

## 产品背景

AcuRev-4100 是导轨安装多回路电表，支持 24 路电流输入、12 个用户通道、3 相电压测量。面向北美/欧洲工业/商业多回路分表计量场景。通过通信扩展模块（ACM-41-WEB2）可接入网关测试体系。

**需求文档：**
- 产品需求规格 v1.03：`requirements/raw/AcuRev4100产品需求规格说明书 v1.03 20260519 2.docx` → 摘要：`requirements/summaries/v1.03_20260519.md`
- Sprint 2（USB + WEB2 交互）v1.02：`requirements/raw/4100 Sprint 2 软件需求规格书 v1.02 20260403.docx` → 摘要：`requirements/summaries/sprint2_v1.02_20260403.md`

**Bug 追踪：**
- 精简索引：`bugs/INDEX.md`

---

## 核心测量能力

| 类别 | 精度 | 说明 |
|------|------|------|
| 电压（相/线） | 0.1% | Van/Vbn/Vcn + Vab/Vbc/Vca |
| 电流 | 0.1% | 24 路输入 + 12 用户通道 + 系统 |
| 有功/无功/视在功率 | 0.2% | 各路/用户/相/系统总 |
| 有功电能 | IEC 62053-22 Class 0.2s | 进/出/净/总，update 100ms |
| 谐波 | 1% | 2~31次（可选 2~63次）；% + 幅度 + 角度 |
| 频率 | 0.002 Hz | 42.5~69 Hz |
| 需量 | — | Block/Sliding Window；记录最大需量 |
| TOU | — | 尖峰谷平4费率，12季节，14时刻表 |

---

## 支持接线方式

| 方式 | 用户数 | 典型场景 |
|------|------|---------|
| 1E2W | — | 单相 120VLN |
| 2E3W 1Phase | 12 | 北美 120/240V 家用/商用 |
| 2E3W Network | 12 | 三相三线网络 |
| 3E4W Y | 8 | 三相四线（最常用） |
| 2E3W Delta | 12 | 三相三线Delta |

---

## IO 配置

| IO | 数量 | 功能 |
|----|------|------|
| DI | 4 | 状态指示 / 脉冲计数（最大 50Hz） |
| DO | 8 | 能量脉冲输出 / 告警指示 |
| RO | 2 | 告警指示 / 继电器控制（Latch/Pulse） |
| AI | — | 0-20mA/0-10V，3段分段线性转换 |
| AO | — | 0-20mA/0-10V，跟随实时参数 |

---

## 通信协议

- **Modbus TCP**：以太网，最多 4 个并发客户端，端口非标准（参考 config.py）
- **Modbus RTU**：RS485 主口，波特率 1200~115200
- **BACnet MS/TP**：RS485 主口（与 Modbus RTU 分时使用）

---

## 电气规格（测试关键参数）

| 项目 | 规格 |
|------|------|
| 电流输入类型 | mV（333mV/RCT，0~400mV）或 mA（80mA/100mA，0~120mA） |
| DI 脉冲频率上限 | 50Hz（v1.02 起由 100Hz 修订） |
| 工作温度 | -25°C ~ +70°C |
| 防护等级 | IP41 本体，IP54 HMI 前面板 |
| 认证 | UL 61010-1，IEC 62052-31 |

---

## 关联项目

- **ACM-41-WEB2**：通信扩展模块，附加在本表本体，提供 BACnet/IP、Modbus TCP、MQTT 等北向协议 → `knowledge/gateway/AcuRev4100WEB2/context.md`
