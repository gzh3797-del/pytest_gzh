# Admin 管理员

**路由**：`#/admin/users`

## 概述

Admin 模块是组织级别的管理面板，管理当前组织下的用户、子组织、租户、订阅和定制化配置。

## 子页面（Tab）

| Tab | 说明 |
|-----|------|
| **Users** | 用户管理（组织内用户） |
| **Organization** | 子组织管理 |
| **Tenants** | 租户管理 |
| **Subscription** | 订阅管理 |
| **Customization** | 界面/功能自定义配置 |

---

## Users（用户管理）

### 创建用户

| 用户类型 | 说明 |
|---------|------|
| **普通 User（user）** | 基础权限，只读/有限操作 |
| **Org Admin** | 组织管理员，可管理本组织内所有内容 |
| **Tenant** | 租户用户（Facility View 限定） |

**创建流程（Add User）：**
1. 点击 Add User
2. 输入 user name、Email
3. 选择组织和角色
4. 提交后用户收到初始密码邮件
5. 用户可用初始密码登录并修改密码

**必填字段：** Email（格式校验）、角色

### 用户权限说明

| 角色 | Dashboard | Installation | Analysis | Billing | Data | Admin |
|------|-----------|-------------|---------|---------|------|-------|
| Normal User | ✅（受限） | ✅（只读） | ✅ | ❌ | ✅（只查看） | ❌ |
| Org Admin | ✅ | ✅（全部） | ✅ | ✅ | ✅ | ✅ |
| Tenant | ✅（仅 Facility View） | ❌ | ✅（受限） | ❌ | ❌ | ❌ |

### Two Factor Authentication（双因素认证）

- 状态：Enabled / Disabled
- 默认：Disabled
- 开启后登录需额外验证码

### 用户表格字段

| 字段 | 说明 |
|------|------|
| userName | 用户名 |
| Email | 邮箱地址 |
| Login Count | 登录次数 |
| Last Login Time | 最近登录时间 |
| Last Time Password Changed | 最近改密时间 |
| Two Factor Authentication | 双因素认证状态 |
| Action | 历史记录 / 查看权限 / 编辑 / 删除 |

---

## Organization（组织管理）

管理当前组织下的子组织结构（层级组织树）：
- 创建子组织
- 配置子组织权限范围
- 子组织 Admin 管理本子组织内的 Facility 和设备

---

## Tenants（租户管理）

管理租户账号（Tenant 类型用户，区别于 org 类型）：
- 租户仅能访问被分配的 Facility
- 租户无 Installation 权限
- 租户可查看 Dashboard 和 Analysis（设施范围内）

---

## Subscription（订阅管理）

### Plan 类型

| Plan | 说明 |
|------|------|
| **Free** | 免费版，功能受限 |
| **Lite** | 基础版，支持基础 Billing |
| **AcuBilling** | 计费专项 |
| **AcuPQ** | 电能质量专项 |
| **Plus** | 增强版，支持高级功能 |
| **AcuEMS** | 最全功能版，面向能源管理系统 |

### Plan 功能对照表

| 功能 | Free | Lite | AcuBilling | AcuPQ | Plus | AcuEMS |
|------|------|------|-----------|-------|------|--------|
| Billing（基础） | ❌ | ✅ | ✅ | ❌ | ✅ | ✅ |
| Water/Gas Billing | ❌ | ❌ | ✅ | ❌ | ✅ | ✅ |
| Power Quality | ❌ | ❌ | ❌ | ✅ | ❌ | ✅ |
| Carbon Model | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Utility Bills Analysis | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Utility Bills Report | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Submeter Bills | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| M&V | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Budget Tracking | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |

### Subscription Tier（订阅层级）

每个 Plan 下可配置以下 Tier：
- **Standard Meter Point Tier**：标准计量点数量和费率
- **Power Quality Meter Point Tier**（第四阶段新增）：PQ 专项计量点订阅

**Admin 配置 Subscription：**
1. 进入 Admin → Subscription
2. 选择 Plan 类型
3. 配置各 Tier 的数量（Number of Subscription）和费率（Rate）
4. 提交后订阅生效

**历史订阅兼容：** 新增 Power Quality Tier 后，历史订阅默认 0，不影响已有订阅。

### 订阅 Interval Limit（ACU-820）

- 订阅支持设置 Interval 限制
- 不同 Plan 对应不同的数据采集间隔限制

---

## Customization（自定义）

支持以下自定义配置：

| 配置项 | 说明 |
|--------|------|
| 组织 Logo | 上传组织 Logo（用于界面品牌化） |
| 界面颜色主题 | 自定义主色调 |
| 自定义邮件模板 | 自定义告警/报告邮件格式 |
| 默认语言 | 设置组织默认显示语言 |
| **White Labeling（第三阶段）** | 平台品牌白标（自定义平台名称、Logo、颜色） |

---

## 测试要点

| 测试项 | 验证点 |
|--------|--------|
| 创建普通 user | 邮件收到初始密码，可登录 |
| 创建 org admin | 权限正确，可管理本组织 |
| 删除用户 | 用户删除后无法登录 |
| 订阅 Plan 切换 | 功能开关随 Plan 变化 |
| Power Quality Tier 配置 | 数量和费率保存正确 |
| Zoho 账单验证 | Zoho 接收到订阅邮件 |
| White Labeling | Logo/颜色/名称展示正确 |
