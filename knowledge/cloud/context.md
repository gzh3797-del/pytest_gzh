# AcuCloud — 云端能源管理平台测试

## 测试设计必读（固定结构速查区）

> strategy-design / testcase-design / coverage-check 的「知识库提取表」以本区块为权威来源，正文为详情补充。正文变更涉及下列字段时必须同步更新本区块。

| 字段 | 内容 |
|------|------|
| 设备清单 | 物理设备：Acuvim II / Acuvim III（智能电表）、AcuLink810（网关，**有下挂设备时禁止直接删除**）；软件设备类型：Calculated（公式）/ Manual Metering / Single Parameter / TOTAL（自动生成） |
| 协议与端口 | Web 前端（Vue 3 + Element Plus，Hash 路由）+ /api/v1、/api/report REST API（Bearer Token 认证）；数据转发 MQTT / HTTP / FTP |
| 接线方式 | 无（云平台）；Super Admin 含接线检测管理功能 |
| 容量与边界 | **订阅 Plan 6 级功能限制（本项目核心测试维度）**：Free（无 Billing，数据间隔≥1h）/ Lite（仅电力 Billing，≥15min）/ AcuBilling（含水/气 Billing，≥5min）/ AcuPQ（Power Quality，≥5min）/ Plus（+Carbon+UB Analysis+M&V）/ AcuEMS（全功能）；能源类型 3 种：电力/水/天然气 |
| 共存/联动约束 | 功能可用性随 Plan 变化——**每个功能的用例必须评估 Plan 维度**（低 Plan 不可见/不可用）；SMS 告警仅 Plus 及以上；完整矩阵 → requirements/raw/18_subscription_plan.md |
| 安全与默认状态 | Bearer Token（Header Authorization）；角色体系（SUPER / Admin / 租户等，White Labeling） |
| 高频缺陷模式 | → 详见 bugs/INDEX.md（如有） |
| 测试环境要点 | 前端 http://acucloud-test-451397146.cn-northwest-1.elb.amazonaws.com.cn（AWS 宁夏区）；测试账号见下方「项目背景」节（SUPER 角色） |

## 项目背景

AcuCloud 是爱博精电开发的 SaaS 能源管理平台，面向工业/商业设施，提供能耗监控、计费自动化、电能质量分析、碳排放追踪和报告推送。支持电力/水/天然气三种能源类型，通过智能电表（Acuvim II/III）和网关（AcuLink810）采集数据。

- 测试环境：http://acucloud-test-451397146.cn-northwest-1.elb.amazonaws.com.cn
- 测试账号：renjie.jiao@accuenergy.com（SUPER 角色，org: AG PROYECTOS Y SERVICIOS, S.A.）
- 技术栈：Vue 3 + Element Plus，Hash 路由，/api/v1 REST API，AWS 宁夏区

## 当前模块

| 模块 | 功能概述 | 需求文档 |
|------|---------|---------|
| Home | 组织概览地图、离线告警横幅 | requirements/raw/03_home.md |
| Dashboard | Widget（Energy/Sankey/Canvas 等 8 种）、Org/Facility 双视图 | requirements/raw/16_dashboard.md |
| Installation | 设施/设备/计量点/告警规则/Portfolio 管理 | requirements/raw/04_installation.md |
| Analysis | Energy/Realtime+预测/Heatmap/Schedule/M&V | requirements/raw/07_analysis.md |
| Power Quality | THD/PF/Voltage、PQ Tier 订阅 | requirements/raw/08_power_quality.md |
| Carbon Model | Scope1/2、排放因子、Plan 要求 Plus+ | requirements/raw/09_carbon_model.md |
| Billing | Rates（LDC/TOU/Tiered）、Auto/Manual、账单分析 | requirements/raw/05_billing.md |
| Utility Bills | 账单录入、OCR、分析、预算追踪 | requirements/raw/06_utility_bills.md |
| Report | 能耗/计费/PQ/UB 报告配置与生成 | requirements/raw/10_report.md |
| Data | 导出/导入/编辑/VEE/转发（MQTT/HTTP/FTP） | requirements/raw/11_data.md |
| Alerts/Logs | 6 种告警类型、SMS（Plus+）、告警日志 | requirements/raw/17_alerts.md |
| Gas Model | 天然气单位转换（cuft→m³）、VEE | requirements/raw/15_gas_model.md |
| Admin | 用户/组织/租户/订阅/White Labeling | requirements/raw/12_admin.md |
| Super Admin | 全平台管理/Plan/Device Model/接线检测 | requirements/raw/13_super_admin.md |
| API | Bearer Token 认证、核心端点 | requirements/raw/14_api.md |

## 订阅 Plan（功能限制核心）

| Plan | 定位 | 关键限制 |
|------|------|---------|
| Free | 免费版 | 无 Billing，最小间隔 1h |
| Lite | 轻量版 | 仅电力 Billing，最小间隔 15min |
| AcuBilling | 计费专项 | 含水/气 Billing，最小间隔 5min |
| AcuPQ | 电能质量专项 | Power Quality，最小间隔 5min |
| Plus | 增强版 | Billing + Carbon + Utility Bills Analysis + M&V |
| AcuEMS | 全功能版 | 全部功能（含 PQ + Carbon） |

完整矩阵 → requirements/raw/18_subscription_plan.md

## 网络配置

| 项目 | 值 |
|------|-----|
| 前端 URL | http://acucloud-test-451397146.cn-northwest-1.elb.amazonaws.com.cn |
| API Base | /api/v1，/api/report |
| 云服务 | AWS 宁夏区（cn-northwest-1） |
| 认证方式 | Bearer Token，Header Authorization |

## 支持设备

| 设备 | 类型 | 说明 |
|------|------|------|
| Acuvim II | 智能电表 | Physical Device |
| Acuvim III | 智能电表 | Physical Device |
| AcuLink810 | 网关 | Gateway Device，有下挂设备时禁止直接删除 |

软件设备类型：Calculated（公式计算）/ Manual Metering / Single Parameter / TOTAL（自动生成）

## 需求文档

| 文件 | 说明 |
|------|------|
| requirements/raw/ | 原始探索文档（01~19 + STRUCTURE/FLOWCHART/README） |
| requirements/summaries/v1_20260518.md | 全模块需求摘要（2026-05-18 版本） |
