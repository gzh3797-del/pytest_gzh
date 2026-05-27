# AcuCloud 知识库全局流程图

## 系统模块总览

```mermaid
graph TB
    %% 顶层入口
    Login["🔐 Login<br>renjie.jiao@accuenergy.com"]
    
    Login --> OrgView["Organization View<br>(当前组织: AG PROYECTOS)"]
    Login --> FacView["Facility View<br>(单设施视图)"]

    %% 主模块
    OrgView --> INST["📦 Installation<br>#/installation"]
    OrgView --> DASH["📊 Dashboard<br>#/dashboard"]
    OrgView --> ANAL["📈 Analysis<br>#/analysis"]
    OrgView --> BILL["💰 Billing<br>#/billing"]
    OrgView --> UTIL["🧾 Utility Bills<br>#/utilityBills"]
    OrgView --> RPT["📄 Report<br>#/report"]
    OrgView --> DATA["💾 Data<br>#/data"]
    OrgView --> LOGS["📝 Logs<br>#/logs"]
    OrgView --> ADMIN["⚙️ Admin<br>#/admin"]
    OrgView --> SA["🔑 Super Admin<br>#/superAdmin"]
    OrgView --> PQ["⚡ Power Quality<br>#/powerQuality"]
    OrgView --> CM["🌿 Carbon Model<br>#/carbonModel"]

    %% Installation 子模块
    INST --> FAC["Facilities<br>设施管理"]
    INST --> DEV["Devices<br>设备管理"]
    INST --> MP["Meter Points<br>计量点"]
    INST --> ALT["Alerts配置<br>告警规则"]
    INST --> PORT["Portfolio<br>项目集"]

    %% Devices 子类型
    DEV --> PHY["Physical<br>物理电表/网关"]
    DEV --> CALC["Calculated<br>计算表"]
    DEV --> MANUAL["Manual Metering<br>手动抄表"]
    DEV --> SINGLE["Single Parameter<br>单参数设备"]

    %% Analysis 子模块
    ANAL --> ENERGY["Energy<br>能耗趋势"]
    ANAL --> ACCUM["Accumulative<br>累计分析"]
    ANAL --> RT["Realtime<br>实时+预测基线"]
    ANAL --> HEAT["Heatmap<br>热力图"]
    ANAL --> SCHED["Schedule<br>排程分析"]
    ANAL --> ANN["Annotations<br>标注"]
    ANAL --> MV["M&V<br>测量与验证"]

    %% Billing 子模块
    BILL --> RATES["Rates<br>费率配置"]
    BILL --> ACCT["Accounts<br>账户配置"]
    BILL --> AUTO["Auto Billing<br>自动计费"]
    BILL --> MANU["Manual Billing<br>手动计费"]
    BILL --> PREV["Previous Bills<br>历史账单"]
    BILL --> BA["Billing Analysis<br>账单分析"]

    %% Data 子模块
    DATA --> DEXP["Data Exports<br>数据导出"]
    DATA --> DFWD["Data Forwards<br>数据转发"]
    DATA --> DIMP["Data Import<br>数据导入"]
    DATA --> DEDIT["Data Edit<br>数据编辑"]
    DATA --> VEE["Manual VEE<br>验证/估算/编辑"]
    DATA --> MMEXP["Max Min Export<br>最值导出"]
```

---

## 数据采集链路

```mermaid
graph LR
    HW["🔌 硬件设备<br>Acuvim II/III<br>AcuLink810<br>Gas Meter<br>Water Meter"]
    
    HW -->|"5min/15min<br>JSON/CSV"| RECV["📡 Receiver<br>数据接收服务"]
    
    RECV -->|"存储"| INFLUX["🗄️ InfluxDB<br>时序数据库"]
    RECV -->|"处理"| VEE2["VEE<br>验证/估算/编辑"]
    
    VEE2 --> INFLUX
    
    INFLUX -->|"API"| APIV1["Backend API<br>/api/v1/*"]
    
    APIV1 -->|"WebSocket/HTTP"| FE["Vue 3 前端<br>SPA"]
    
    FE --> USER["👤 用户界面"]
```

---

## 设备类型与数据流

```mermaid
graph TB
    subgraph "物理设备"
        A1["Acuvim II<br>电力 Electricity"]
        A2["Acuvim III<br>电力+电能质量 PQ"]
        A3["AcuLink810<br>Gateway 网关"]
        A4["Typical Gas Meter<br>天然气 Gas"]
        A5["Water Meter<br>水务 Water"]
    end

    subgraph "虚拟设备"
        V1["Calculated Meter<br>计算表（公式）"]
        V2["Manual Metering<br>手动抄表"]
        V3["Single Parameter<br>单参数设备"]
        V4["TOTAL<br>设施汇总（自动创建）"]
    end

    subgraph "数据单元"
        MP1["Meter Point<br>计量点"]
    end

    A1 --> MP1
    A2 --> MP1
    A3 -->|"下挂多个设备"| A1
    A4 --> MP1
    A5 --> MP1
    V1 -->|"引用其他MP"| MP1
    V2 --> MP1
    V3 --> MP1
    V4 -->|"聚合所有子MP"| MP1
```

---

## Billing 计费流程

```mermaid
graph TD
    RATESETUP["费率配置<br>Rates → LDC → Rate Structure<br>(电/水/气/污水)"]
    
    RATESETUP -->|"绑定"| ACCTCFG["账户配置<br>Accounts<br>Facility+MP+Rate"]
    
    ACCTCFG --> AUTOBILL["自动账单<br>Auto Billing<br>定时生成"]
    ACCTCFG --> MANBILL["手动账单<br>Manual Billing<br>按需触发"]
    
    AUTOBILL --> PREVBILL["历史账单<br>Previous Bills"]
    MANBILL --> PREVBILL
    
    PREVBILL --> EXPORT["Excel 导出<br>汇总/子工作簿/TOU Report"]
    
    PREVBILL --> BANALYSIS["账单分析<br>Billing Analysis"]
    
    PREVBILL --> SUBMIT["Submeter Bills<br>子表账单"]
    
    PREVBILL --> ZOHO["Zoho 集成<br>Billing → Zoho CRM/Books"]
```

---

## Utility Bills（公共事业账单）流程

```mermaid
graph TD
    UPLOAD["上传方式"]
    
    UPLOAD -->|"手动录入"| ENTRY["Bill Entry<br>手动输入账单数据"]
    UPLOAD -->|"PDF 上传"| OCR["OCR Upload<br>PDF 识别 → 人工审核 → 提交"]
    
    ENTRY --> BILLDB["账单数据库"]
    OCR --> BILLDB
    
    BILLDB --> UBA["Utility Bills Analysis<br>跨设施账单分析"]
    BILLDB --> UBRPT["Utility Bills Report<br>账单报告"]
    BILLDB --> UALERTS["Utility Bill Alerts<br>账单超额告警"]
    BILLDB --> BUDGET["Budget Tracking<br>预算对比追踪"]
    BILLDB --> TOTAL["Utility Bills Total<br>账单汇总统计"]
```

---

## 告警体系

```mermaid
graph LR
    subgraph "告警类型"
        OA["Offline Alert<br>设备离线"]
        PA["Parametric Alert<br>参数超阈值"]
        FA["Flatline Alert<br>数据固定不变"]
        RA["Rate Alarm<br>费率超阈值"]
        MA["Meter Resealing Alert<br>水表密封（水专用）"]
        AH["After Hours Alarm<br>非工作时段异常"]
        UBA2["Utility Bill Alerts<br>账单超额"]
    end

    subgraph "通知渠道"
        EMAIL["📧 Email 邮件"]
        SMS["📱 SMS 短信"]
    end

    subgraph "日志记录"
        LOG["Alert Logs<br>告警日志"]
        DEVC["Device 列表<br>Alarm Count"]
    end

    OA --> EMAIL
    PA --> EMAIL
    FA --> EMAIL
    OA --> SMS
    PA --> SMS
    RA --> EMAIL
    UBA2 --> EMAIL
    UBA2 --> SMS

    OA --> LOG
    PA --> LOG
    FA --> LOG
    RA --> LOG

    OA --> DEVC
    PA --> DEVC
    FA --> DEVC
```

---

## 订阅 Plan 与功能权限

```mermaid
graph TB
    subgraph "Plan 层级"
        FREE["Free<br>免费版"]
        LITE["Lite<br>基础版"]
        ACUB["AcuBilling<br>计费专项"]
        ACUPQ["AcuPQ<br>电能质量专项"]
        PLUS["Plus<br>增强版"]
        ACUEMS["AcuEMS<br>全功能版"]
    end

    subgraph "功能模块"
        F1["✅ 基础监控<br>Installation+Dashboard"]
        F2["✅ 基础 Billing（电）"]
        F3["✅ Water/Gas Billing"]
        F4["✅ Power Quality"]
        F5["✅ Carbon Model"]
        F6["✅ Utility Bills Analysis/Report"]
        F7["✅ M&V + Budget Tracking"]
        F8["✅ SMS 通知"]
        F9["✅ Billing Zoho 集成"]
    end

    FREE --> F1
    LITE --> F1
    LITE --> F2
    ACUB --> F1
    ACUB --> F2
    ACUB --> F3
    ACUPQ --> F1
    ACUPQ --> F4
    PLUS --> F1
    PLUS --> F2
    PLUS --> F3
    PLUS --> F5
    PLUS --> F6
    PLUS --> F7
    PLUS --> F8
    PLUS --> F9
    ACUEMS --> F1
    ACUEMS --> F2
    ACUEMS --> F3
    ACUEMS --> F4
    ACUEMS --> F5
    ACUEMS --> F6
    ACUEMS --> F7
    ACUEMS --> F8
    ACUEMS --> F9
```

---

## 知识库文件索引

```
acucloud_knowledge/
├── 01_overview.md          → 平台概述、技术架构、测试环境
├── 02_navigation.md        → 导航结构、路由规则
├── 03_home.md              → 首页
├── 04_installation.md      → 设施/设备/计量点管理（完整版）
├── 05_billing.md           → 计费模块（Rates/Accounts/Auto/Manual/Utility Bills）
├── 06_utility_bills.md     → 公共事业账单（原版）
├── 07_analysis.md          → 能耗分析（Energy/Realtime/M&V/Schedule/Heatmap）
├── 08_power_quality.md     → 电能质量分析 + PQ Meter Point Tier
├── 09_carbon_model.md      → 碳排放模型 + Plan 要求
├── 10_report.md            → 报告管理（能耗/计费/Utility Bills/TOU）
├── 11_data.md              → 数据管理（导出/导入/编辑/VEE/转发）
├── 12_admin.md             → 组织管理员（用户/订阅/自定义）
├── 13_super_admin.md       → 超级管理员（全局管理/Plan/Device Model）
├── 14_api.md               → API 接口规范
├── 15_gas_model.md         → 天然气模型（数据格式/单位转换/分析）
├── 16_dashboard.md         → 仪表盘（Widget/kVA/Sankey/Auto Refresh）
├── 17_alerts.md            → 告警体系（类型/通知/日志）
├── 18_subscription_plan.md → 订阅 Plan 详细对照（功能矩阵/Tier）
└── FLOWCHART.md            → 本文件（系统流程图）
```
