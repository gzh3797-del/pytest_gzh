# AcuCloud 测试环境知识库

> 初次探索时间：2026-05-14  
> 知识库更新时间：2026-05-18（全模块 UI 重探、修正 Device Detail Tab 名称、新增 UI 精确细节文档）  
> 环境：http://acucloud-test-451397146.cn-northwest-1.elb.amazonaws.com.cn  
> 登录账号：renjie.jiao@accuenergy.com（SUPER 角色）  
> 当前组织：AG PROYECTOS Y SERVICIOS, S.A.（orgId: 431，Reseller 类型）

## 文档目录

| 文件 | 内容 | 更新情况 |
|------|------|---------|
| [01_overview.md](01_overview.md) | 平台概述、技术架构、环境配置 | 初版 |
| [02_navigation.md](02_navigation.md) | 完整导航结构与路由映射 | 初版 |
| [03_home.md](03_home.md) | 首页 — 组织概览与设施地图 | 初版 |
| [04_installation.md](04_installation.md) | 设施/设备/计量点/告警/删除优化/CALC多公式 | ✅ 已增补 |
| [05_billing.md](05_billing.md) | 计费/费率/Utility Bills/OCR/Budget/Zoho | ✅ 已增补 |
| [06_utility_bills.md](06_utility_bills.md) | 公用事业账单（原版） | 初版 |
| [07_analysis.md](07_analysis.md) | 能耗/Realtime预测/M&V/Schedule TOU/Compare | ✅ 已增补 |
| [08_power_quality.md](08_power_quality.md) | 电能质量 + PQ Meter Point Tier 订阅 | ✅ 已增补 |
| [09_carbon_model.md](09_carbon_model.md) | 碳排放模型 + Plan 支持矩阵 | ✅ 已增补 |
| [10_report.md](10_report.md) | 报告/Utility Bills Report/TOU Report/水报告 | ✅ 已增补 |
| [11_data.md](11_data.md) | 数据导出/导入/编辑/VEE/Max Min Export | ✅ 已增补 |
| [12_admin.md](12_admin.md) | 用户/订阅Plan矩阵/PQ Tier/White Labeling | ✅ 已增补 |
| [13_super_admin.md](13_super_admin.md) | 超级管理员/Plan类型/SN搜索/接线配置 | ✅ 已增补 |
| [14_api.md](14_api.md) | API 接口与认证 | 初版 |
| [15_gas_model.md](15_gas_model.md) | 天然气模型（数据格式/单位转换/分析） | ✅ 新增 |
| [16_dashboard.md](16_dashboard.md) | 仪表盘（Widget/kVA/Sankey/自动刷新） | ✅ 新增 |
| [17_alerts.md](17_alerts.md) | 告警体系（类型/SMS/Rate Alarm/日志） | ✅ 新增 |
| [18_subscription_plan.md](18_subscription_plan.md) | 订阅Plan详细对照（功能矩阵/Tier配置） | ✅ 新增 |
| [FLOWCHART.md](FLOWCHART.md) | 系统全局流程图（Mermaid格式） | ✅ 新增 |
| [19_ui_details.md](19_ui_details.md) | 全模块精确 UI 细节：按钮名/表头/过滤器（2026-05-18 实测） | ✅ 新增 |

## 知识来源

| 来源 | 内容 |
|------|------|
| 现场界面探索（2026-05-14） | 截图、DOM结构、API响应 |
| 全功能测试用例_ORG.xlsx | Installation/Billing/Analysis/Report/Data/Admin/SuperAdmin |
| 第二阶段测试用例.xlsx | Dashboard/MV/Power_Quality/Alert Status/Billing_Water/Utility_Bills |
| 第三阶段测试用例.xlsx | Plan/Carbon Model/Device删除/Utility Bills/Facility UI/Alert Logs |
| 第四阶段测试用例.xlsx | Gas/PQ Meter Tier/Calculated多公式/Realtime/OCR/TOU/Budget/Zoho |
| Playwright 自动化脚本 | 设备创建流程（CALC公式/Manual Metering/Reading Cycle）的技术细节 |

## 快速参考

### 测试账号
- 邮箱：`renjie.jiao@accuenergy.com`
- 密码：`n4a6pJ7oKwRHZaPdowtfk`

### 已创建的测试设备（Facility: 2026051412, facilityId: 6241）
- Physical: TEST2026051412PHY01、TEST2026051412PHY02（Acuvim II）
- Physical: TEST2026051412PHY03（Acuvim III）
- Gateway: TEST2026051412GW01（AcuLink810）
- Single Parameter: TEST2026051412SP01
- Calculated Meter: calc_2026051412（公式: $2516418_TEST2026051412PHY02.VLNa_V）
- Manual Metering: manual_2026051412

### 关键 API 端点

| 功能 | 端点 |
|------|------|
| 设备列表 | POST `/api/v1/devices/list` |
| 设施列表 | POST `/api/v1/facilities/list` |
| 计量点列表 | POST `/api/v1/meterpoints/list` |
| 能耗查询 | POST `/api/v1/analysis/energy` |
| 账单列表 | POST `/api/v1/billing/previousBills/list` |

### 截图目录

所有截图位于 `screenshots/` 目录下：
- `screenshots/`：初版截图（01_login_filled.png 等）
- `screenshots/create/`：设备创建过程截图（D3_*, D4_*, FINAL_* 等）
