# AcuCloud 知识库结构总览

> 生成时间：2026-05-19  
> 基于 2026-05-18 全模块 UI 实测  
> 环境：http://acucloud-test-451397146.cn-northwest-1.elb.amazonaws.com.cn  
> 测试账号：renjie.jiao@accuenergy.com（SUPER 角色，org: AG PROYECTOS Y SERVICIOS, S.A.）

---

## 一、知识库文件目录

```
acucloud_knowledge/
├── README.md                → 知识库索引入口（文件列表 + 快速参考）
├── STRUCTURE.md             → 本文件（系统结构总览）
├── FLOWCHART.md             → 系统全局流程图（Mermaid 格式）
│
├── 平台基础
│   ├── 01_overview.md       → 平台概述、技术架构（Vue3/Element Plus/InfluxDB）
│   ├── 02_navigation.md     → 完整路由结构与导航映射
│   └── 14_api.md            → API 认证、请求格式、核心端点
│
├── 功能模块
│   ├── 03_home.md           → 首页（地图、离线告警横幅、组织切换）
│   ├── 04_installation.md   → 安装管理（设施/设备/计量点/告警/Portfolio）
│   ├── 05_billing.md        → 计费（费率/账户/自动/手动/Utility Bills/Zoho）
│   ├── 06_utility_bills.md  → 公用事业账单（账单录入/分析/OCR）
│   ├── 07_analysis.md       → 能耗分析（Energy/Realtime/M&V/Heatmap/Schedule）
│   ├── 08_power_quality.md  → 电能质量（THD/PF/电压/PQ Tier）
│   ├── 09_carbon_model.md   → 碳排放模型（Scope1/2/排放因子/Plan要求）
│   ├── 10_report.md         → 报告管理（能耗/计费/Utility Bills/TOU）
│   ├── 11_data.md           → 数据管理（导出/导入/编辑/VEE/转发）
│   ├── 12_admin.md          → 组织管理员（用户/订阅/Customization）
│   ├── 13_super_admin.md    → 超级管理员（全局管理/Plan/Device Model/Wiring）
│   ├── 15_gas_model.md      → 天然气模型（数据格式/单位转换/m³/cuft）
│   ├── 16_dashboard.md      → 仪表盘（Widget/kVA/Sankey/Auto Refresh）
│   └── 17_alerts.md         → 告警体系（类型/SMS/Rate Alarm/日志）
│
└── 参考资料
    ├── 18_subscription_plan.md → 订阅 Plan 功能矩阵（Free/Lite/AcuBilling/Plus/AcuEMS）
    └── 19_ui_details.md        → 精确 UI 细节（按钮名/表头/过滤器，2026-05-18 实测）
```

---

## 二、系统模块层级结构

```
AcuCloud 平台
├── 🏠 Home                          → 03_home.md
│   ├── 组织概览地图
│   ├── 离线设备告警横幅
│   └── Organization/Facility 视图切换
│
├── 📊 Dashboard                     → 16_dashboard.md
│   ├── Organization View Dashboard（跨设施）
│   ├── Facility View Dashboard（单设施）
│   └── Widget 类型
│       ├── Energy Analysis（kWh/kVA）
│       ├── Data Widget（实时数据表格）
│       ├── Alert Widget（活跃告警）
│       ├── Billing Statistics（账单统计）
│       ├── Sankey Diagram（能流图）
│       ├── Canvas Feature（自定义画布）
│       ├── Energy Intensity（单位面积能耗）
│       └── Energy Generation（发电量）
│
├── 📦 Installation                  → 04_installation.md / 19_ui_details.md
│   ├── Facilities（设施管理）
│   │   ├── 创建/编辑/删除
│   │   ├── 多层级 Facility（父/子）
│   │   └── 自动生成 TOTAL 设备和计量点
│   ├── Devices（设备管理）
│   │   ├── Physical（Acuvim II/III、物理电表）
│   │   ├── Gateway（AcuLink810 网关）
│   │   ├── Calculated（计算表，公式: $deviceId_SN.paramCode_unit）
│   │   ├── Manual Metering（手动抄表）
│   │   └── Single Parameter（单参数设备）
│   ├── Meter Points（计量点）
│   │   ├── 过滤：Facilities / Devices / Subscription Tier（3个"Please Select"下拉）
│   │   └── 类型：Physical / TOTAL / Calculated
│   ├── Alerts（告警规则）            → 17_alerts.md
│   │   ├── Alert Event（告警事件）
│   │   └── Alert Rules（告警规则）
│   └── Portfolio（项目集）
│
├── 📈 Analysis                      → 07_analysis.md
│   ├── Energy（能耗趋势）
│   ├── Accumulative（累计）
│   ├── Realtime（实时 + 预测基线）
│   ├── Heatmap（热力图）
│   ├── Schedule（排程分析 + TOU 对比）
│   ├── Annotations（数据标注）
│   └── M&V（测量与验证，需 Plus/AcuEMS）
│
├── ⚡ Power Quality                  → 08_power_quality.md
│   ├── THD（电流/电压总谐波失真）
│   ├── Power Factor（功率因数）
│   ├── Voltage（电压指标）
│   └── PQ Meter Point Tier（订阅配置）
│
├── 🌿 Carbon Model                  → 09_carbon_model.md
│   ├── Scope 1（直接排放）
│   ├── Scope 2（电力间接排放）
│   ├── Emission Factors（Global/Facility/Custom）
│   ├── Intensity & Emission Tab
│   └── Facility Emissions Tab
│
├── 💰 Billing                       → 05_billing.md
│   ├── Rates（费率配置）
│   │   ├── LDC（电网公司费率）
│   │   ├── TOU（分时电价）
│   │   └── Tiered（阶梯定价）
│   ├── Accounts（账户配置，绑定 Facility+MP+Rate）
│   ├── Auto Billing（自动账单）
│   ├── Manual Billing（手动账单）
│   ├── Previous Bills（历史账单）
│   └── Billing Analysis（账单分析）
│       ├── Trend Analysis
│       ├── TOU Allocation Analysis
│       └── Submeter Rate Comparison
│
├── 🧾 Utility Bills                 → 06_utility_bills.md / 05_billing.md
│   ├── Account（公用事业账户配置）
│   ├── Utility Bills（账单录入）
│   ├── Utility Bills Analysis（对比分析）
│   │   ├── Energy Graph
│   │   ├── Costs Graphs
│   │   ├── Energy Trend (EUI)
│   │   └── Proportional Split
│   ├── OCR Upload（PDF 智能识别）
│   ├── Templates（账单模板）
│   ├── Alerts（账单超额告警）
│   └── Budget Tracking（预算追踪）
│
├── 📄 Report                        → 10_report.md
│   ├── Configs（报告配置）
│   │   ├── Energy Report
│   │   ├── Billing Report
│   │   ├── Portfolio Report
│   │   ├── Power Quality Report（BETA）
│   │   └── Utility Bills Report（Plus/AcuEMS）
│   ├── Generate（手动触发）
│   └── Previous Reports（历史查询）
│
├── 💾 Data                          → 11_data.md
│   ├── Data Exports（原始数据导出，xlsx/zip，保留30天）
│   ├── Data Forwards（实时转发 MQTT/HTTP/FTP）
│   ├── Data Import（历史数据补录）
│   ├── Energy Exports（聚合能耗导出）
│   ├── Max Min Export（最大/最小值统计）
│   ├── Data Edit（手动修正历史数据）
│   └── Manual VEE（验证/估算/编辑 VEE）
│
├── 🔔 Logs / Alert Logs             → 17_alerts.md
│   └── 告警日志（类型/设施/时间过滤，分页排序）
│
├── ⚙️ Admin                         → 12_admin.md / 19_ui_details.md
│   ├── Users（组织用户管理）
│   ├── Organization（组织信息）
│   ├── Tenants（租户管理）
│   ├── Subscription（订阅配置）
│   └── Customization（White Labeling/Logo）
│
└── 🔑 Super Admin                   → 13_super_admin.md / 19_ui_details.md
    ├── Users（全平台用户管理）
    ├── Parameter（参数定义）
    ├── Device Model（设备型号管理）
    ├── Organization（全平台组织管理）
    ├── Subscription（全平台订阅管理）
    ├── Plan（Plan 类型配置）
    ├── Release Center（版本发布）
    ├── Global Emission Factors（全球排放因子）
    ├── Global Device Search（跨组织设备 SN 搜索）
    └── Wiring Check（接线检测）
```

---

## 三、核心数据实体关系

```
Organization（组织）
└── Subscription Plan（订阅 Plan）           → 18_subscription_plan.md
    └── Subscription Tier（标准/PQ 计量点层级）

Organization
└── Facility（设施）
    ├── Device（设备）                        → 04_installation.md
    │   ├── Physical / Gateway / Calculated / Manual / SingleParam
    │   └── Meter Point（计量点）
    │       ├── 绑定 Rate Structure → Billing Account
    │       ├── 参与 Analysis 分析
    │       ├── 参与 Data Export/Forward
    │       └── 参与 Report 配置
    └── TOTAL Device（自动聚合，Facility 创建时自动生成）
```

---

## 四、能源类型支持矩阵

| 模块 | Electricity | Water | Gas |
|------|:-----------:|:-----:|:---:|
| Installation / Device | ✅ | ✅ | ✅ |
| Meter Point | ✅ | ✅ | ✅ |
| Billing / Rates | ✅ LDC/TOU/Tiered | ✅ Water Rate | ✅ Gas Rate |
| Billing / Manual | ✅ | ✅ | ✅ |
| Analysis | ✅ | ✅ | ✅ |
| Power Quality | ✅（专属） | ❌ | ❌ |
| Carbon Model | ✅（Scope 2 主要来源） | ❌ | ✅（Scope 1） |
| Gas Model（单位转换） | ❌ | ❌ | ✅ m³/cuft |
| VEE | ✅ 电力 VEE | ✅ 水 VEE | ✅ 气体 VEE |
| Report | ✅ | ✅ Water Report | ✅ |
| Dashboard Widget | ✅ | ✅ | ✅ |

---

## 五、订阅 Plan 功能对照（快速参考）

| 功能 | Free | Lite | AcuBilling | AcuPQ | Plus | AcuEMS |
|------|:----:|:----:|:----------:|:-----:|:----:|:------:|
| Installation + Dashboard | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| 基础 Analysis | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Billing（电） | ❌ | ✅ | ✅ | ❌ | ✅ | ✅ |
| Billing（水/气） | ❌ | ❌ | ✅ | ❌ | ✅ | ✅ |
| Power Quality | ❌ | ❌ | ❌ | ✅ | ❌ | ✅ |
| Carbon Model | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| M&V / Analysis 高级 | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Utility Bills Analysis/Report | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| SMS 通知 | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Zoho 集成 | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| 最小数据间隔 | 1h | 15min | 5min | 5min | 5min | 5min |

> 详细矩阵见 → 18_subscription_plan.md

---

## 六、告警类型速查

| 告警类型 | 触发条件 | 通知渠道 | 适用设备 | 文档 |
|---------|---------|---------|---------|------|
| Offline | 超时无数据 | Email / SMS | 所有 | 17_alerts.md |
| Parametric | 参数超阈值 | Email / SMS | Physical/Gateway/Calc | 17_alerts.md |
| Flatline | 数据长期固定不变 | Email | 所有 | 17_alerts.md |
| Rate Alarm | 费率超阈值 | Email | 电力设备 | 17_alerts.md |
| Meter Resealing | 水表密封异常 | Email | Water 专属 | 17_alerts.md |
| After Hours | 非工时段超阈值 | Email | 所有 | 17_alerts.md |
| Utility Bill Alert | 账单金额超预算 | Email / SMS | 公用事业账单 | 05_billing.md |

---

## 七、天然气模型关键参数

| 维度 | 非网关设备 | 网关设备 |
|------|-----------|---------|
| 上报格式 | CSV | JSON |
| 上报字段 | TOTAL_VOLUME_m3 | TOTAL_VOLUME_cuft |
| 单位换算 | 无需转换 | cuft × 0.0283168 → m³ |
| 存储单位 | m³（InfluxDB） | m³（InfluxDB） |
| 上报大小限制 | ≤ 3MB | 无限制 |

> 详细说明见 → 15_gas_model.md

---

## 八、设备删除规则

| 删除对象 | 前置条件 | 强关联（禁止删除） | 弱关联（警告后可继续） |
|---------|---------|------------------|-------------------|
| Meter Point | — | 被账单/报告引用 | — |
| Device | 先删所有 Meter Point | 810 网关且有下挂设备 | 被账单/报告/数据导出引用 |
| Facility | 先删所有 Device | — | — |

> 删除顺序：Meter Point → Device → Facility

---

## 九、关键 API 端点

| 功能 | 方法 | 端点 | 文档 |
|------|------|------|------|
| 设备列表 | POST | `/api/v1/devices/list` | 14_api.md |
| 设施列表 | POST | `/api/v1/facilities/list` | 14_api.md |
| 计量点列表 | POST | `/api/v1/meterpoints/list` | 14_api.md |
| 能耗查询 | POST | `/api/v1/analysis/energy` | 14_api.md |
| 账单列表 | POST | `/api/v1/billing/previousBills/list` | 14_api.md |
| 菜单权限树 | GET | `/system-menu/tree` | 14_api.md |

---

## 十、UI 注意事项（测试必读）

| 位置 | 注意点 | 来源 |
|------|--------|------|
| Meter Points 过滤器 | 三个 el-select 的 placeholder 均为 "Please Select"，需按**位置索引**（0=Facilities/1=Devices/2=Subscription Tier）区分 | 19_ui_details.md |
| Data Forwards 按钮 | 系统原文拼写错误：`+Add Data Fowrards`（Fowrards 非 Forwards） | 19_ui_details.md |
| Device Detail Tab | 正确名称：OverView / Latest Readings / Wiring / Documents / Alert / Photos / Replacement Records | 04_installation.md |
| Calculated Meter 公式 | 纯常数公式（如"0"）被 API 拒绝（code:501, msg: Calculated Meter Formula is invalid） | 04_installation.md |
| 登录后切换组织 | 使用 el-cascader 选择器；输入后选项出现在 `.el-cascader__suggestion-item` | 01_overview.md |
| 未订阅状态 | `/system-menu/tree` 返回 `code: 603`，大多数页面仍可通过 URL 直接访问 | 18_subscription_plan.md |

---

## 十一、角色权限速查

| 角色 | Installation | Billing | Analysis | Data | Admin | Super Admin |
|------|:------------:|:-------:|:--------:|:----:|:-----:|:-----------:|
| SUPER | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Org Admin | ✅ | ✅ | ✅ | ✅ | ✅（本组织） | ❌ |
| Normal 用户 | 查看 | 查看 | 查看 | 仅导出 | ❌ | ❌ |
| Tenant 用户 | ❌ | 查看自己账单 | 查看（Facility 范围） | ❌ | ❌ | ❌ |

---

## 十二、跨文档引用索引

| 主题 | 主文档 | 补充文档 |
|------|--------|---------|
| 设备创建完整流程 | 04_installation.md | 19_ui_details.md |
| Utility Bills 完整功能 | 05_billing.md | 06_utility_bills.md |
| 告警配置与查询 | 17_alerts.md | 04_installation.md（规则配置 Tab）|
| 订阅 Plan 与功能限制 | 18_subscription_plan.md | 12_admin.md / 13_super_admin.md |
| 天然气数据采集 | 15_gas_model.md | 04_installation.md / 07_analysis.md |
| 电能质量 Tier 配置 | 08_power_quality.md | 18_subscription_plan.md |
| 碳模型 Plan 要求 | 09_carbon_model.md | 18_subscription_plan.md |
| TOU 报告 | 05_billing.md | 10_report.md |
| VEE 数据处理 | 11_data.md | 15_gas_model.md（气体 VEE）|
| UI 精确细节（自动化测试用） | 19_ui_details.md | — |
| 系统流程图（可视化） | FLOWCHART.md | — |
