# AcuCloud 平台概述

## 产品定位

AcuCloud 是 Accuenergy（爱博精电）开发的**云端能源管理平台（Energy Management Software）**，面向工业/商业设施，提供：

- 能耗数据采集与监控（通过智能电表/网关）
- 设施级能耗分析与可视化
- 计费自动化
- 电能质量分析
- 碳排放追踪
- 报告自动生成与推送

## 测试环境信息

| 项目 | 值 |
|------|-----|
| 前端 URL | http://acucloud-test-451397146.cn-northwest-1.elb.amazonaws.com.cn |
| API Base URL | /api/v1 |
| 前端框架 | Vue 3 + Element Plus (el-*) |
| 打包工具 | Vite |
| 版本信息 | 1.0.0 - 2026/5/14 02:17:34 |
| 服务器 | nginx/1.31.0 |
| 云平台 | AWS 宁夏区（cn-northwest-1） |
| 构建哈希 | 1778725018398 |

## 当前登录账户信息

| 字段 | 值 |
|------|-----|
| 邮箱 | renjie.jiao@accuenergy.com |
| 当前组织 | AG PROYECTOS Y SERVICIOS, S.A. |
| 组织 ID | 431 |
| 组织类型 | Reseller（经销商） |
| 角色 | SUPER |
| 用户类型 | org（组织用户） |
| 是否租户 | 否 |
| 界面语言 | English |
| 当前设施 | receiver_20250807（facilityId: 5341） |
| 时区 | Etc/GMT-8 |
| SMS 告警 | 已启用 |
| 免费限额 | 未启用 |
| 计费功能 | 未启用（isBillingEnable: false） |

## 技术架构

```
浏览器（Vue 3 SPA）
    ↓ HTTP/HTTPS
nginx 反向代理
    ↓
后端 API（/api/v1/*）
    ↓
后端报告服务（/api/report/*）
```

### 前端特性

- 单页应用（Hash 路由：`/#/xxx`）
- sessionStorage 存储：`common`（含 token、当前组织、当前设施）
- localStorage 存储：`authToken`（加密存储）、`AcuLang`、`remember`、`isTenant`
- 认证方式：Bearer Token，Header `Authorization: Bearer <uuid>`
- 国际化：支持 English / Español / Français 等多语言

### API 特性

- 统一响应格式：`{ code: number, data: any, msg: string }`
- 成功代码：200
- 未登录：401
- 参数错误：501
- 未订阅：603（含提示信息）
- 账户锁定：连续 5 次密码错误即锁定

## 主要硬件生态

AcuCloud 主要与以下 Accuenergy 硬件集成：

| 设备类型 | 型号示例 |
|---------|---------|
| 智能电表 | Acuvim II、Acuvim III |
| 网关/数据集中器 | AcuLink810 |
| 气体计量 | Typical Gas Meter |

## 平台订阅状态

当前账号组织（orgId: 431）**未订阅 AcuCloud 服务**，系统菜单树（`/system-menu/tree`）返回 code 603，提示联系 yanling.cao@accuenergy.com 订阅。尽管如此，大部分功能模块在测试环境下仍可正常访问。
