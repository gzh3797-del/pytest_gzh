# Super Admin 超级管理员

**路由**：`#/superAdmin/users`

## 概述

Super Admin 是系统最高权限面板，管理所有组织、全局用户、设备型号、订阅计划等系统级资源。

## 子页面（Tab）

| Tab | 说明 |
|-----|------|
| **Users** | 系统级用户管理（全平台所有用户） |
| **Parameter** | 全局参数配置 |
| **Device Model** | 设备型号库管理 |
| **Organization** | 全局组织管理 |
| **Subscription** | 全局订阅管理 |
| **Plan** | 订阅计划配置 |
| **Release Center** | 版本发布中心 |
| **Global Emission Factors** | 全球碳排放因子管理 |
| **Global Device Search** | 全局设备搜索 |
| **Wiring Check** | 接线检测工具 |

---

## Users（系统用户管理）

### 与 Admin/Users 的区别

| 维度 | Admin/Users | Super Admin/Users |
|------|------------|-------------------|
| 范围 | 当前组织内用户 | 全平台所有用户 |
| Organization 列 | 无 | 有（显示所属组织） |
| 删除权限 | 部分用户有 | 视用户级别 |

### 用户表格字段

| 字段 | 说明 |
|------|------|
| User Name | 用户名 |
| Email | 邮箱 |
| Organization | 所属组织名称 |
| Login Count | 登录次数 |
| Last Login Time | 最近登录时间 |
| Last Time Password Changed | 最近改密时间 |
| Two Factor Authentication | 双因素认证状态 |
| Action | 操作 |

### 超级管理员显示优化（SuperAdmin User 显示优化，第二/三阶段）

- 用户列表分页显示优化
- 搜索过滤功能增强

---

## Plan（订阅计划管理）

### 历史 Plan 验证

- 系统原有 Plan 为具有所有权限的 Plan，类型为 **Others**
- 历史订阅正常展示原 Plan（向后兼容）

### Plan 类型详解

| Plan 名 | 描述 | Carbon Model | Utility Bills | Power Quality |
|---------|------|-------------|--------------|---------------|
| Free | 免费基础版 | ❌ | ❌ | ❌ |
| Lite | 轻量版 | ❌ | ❌ | ❌ |
| AcuBilling | 计费专项 | ❌ | ❌ | ❌ |
| AcuPQ | 电能质量 | ❌ | ❌ | ✅ |
| **Plus** | 增强版 | ✅ | ✅ | ❌ |
| **AcuEMS** | 能源管理全功能 | ✅ | ✅ | ✅ |
| Others | 历史遗留（全权限） | ✅ | ✅ | ✅ |

**Plan 在 Plan Detail 表格中的展示：**
- Main Bar 列：显示各模块是否在主菜单栏中出现
- Carbon Model 菜单仅 AcuEMS / Plus 支持
- Utility Bills Analysis 仅 Plus / AcuEMS 支持
- Utility Bills Report 仅 Plus / AcuEMS 支持
- Power Quality Meter Point Tier 仅相关 Plan 支持

---

## Subscription（全局订阅管理）

### SuperAdmin 新建订阅

1. 进入 Super Admin → Subscription
2. 选择组织
3. 配置 Plan 类型
4. 配置各 Tier：
   - Standard Meter Point Tier
   - Power Quality Meter Point Tier（第四阶段新增）
5. 配置 Zoho 集成信息（如需要）
6. 提交订阅

**Power Quality Meter Point Tier（新增订阅类型）：**
- 字段：Number of Subscription（数量）、Rate（费率）
- 新建时验证：提交成功、Zoho 邮件收到、账单 Power Quality Addons 正常计费
- 历史订阅默认值：数量和费率均为 0

### 订阅 Interval Limit（ACU-820）

- 不同订阅层级限制不同的数据采集间隔
- 超出限制时系统提示

---

## Device Model（设备型号库）

管理平台支持的所有硬件设备型号：

| 型号类别 | 示例 |
|---------|------|
| 智能电表 | Acuvim II、Acuvim III |
| 网关 | AcuLink810 |
| 气体计量 | Typical Gas Meter |
| 水表 | 各类 Water Model |

**Device Model 测试（Document/文档）：**
- Acuvim II Model 文档（第三阶段专项：Document_AcuvimII_Model）
- 1310 Model 文档（Document_1310_model）
- Manufacture 选项卡文档下载验证

---

## Parameter（全局参数）

系统级参数配置：
- API 接口参数
- 系统默认值
- 功能开关（Feature Flag）
- VEE 算法参数配置

---

## Global Emission Factors（全球排放因子）

为 Carbon Model 模块提供全球各地区的标准碳排放系数：
- 按地区/国家配置
- 按能源类型配置（电力/天然气/水）
- 单位：kg CO₂/kWh 或 kg CO₂/m³

---

## Global Device Search（全局设备搜索，ACU-SN）

**路由：** `#/superAdmin/globalDeviceSearch`

跨所有组织搜索特定设备：
- 按序列号（SN）搜索
- 按设备名搜索
- 返回：设备所属组织、设施、型号、状态
- **Global SN Search（第四阶段）：** 增强版全局序列号搜索，包含 SN 重复标签功能

---

## Release Center（发布中心）

管理平台版本发布：
- 新版本说明文档
- 功能更新记录（Changelog）
- 用户通知推送（Release Notes）

---

## Wiring Check（接线检测）

**路由：** `#/superAdmin/wiringCheck`（或 接线配置）

辅助工具，帮助现场工程师验证设备接线：
- 实时数据判断 CT/PT 方向
- 识别接线错误（如 CT 反接）
- 第四阶段新增接线配置功能

---

## 测试要点

| 测试项 | 验证点 |
|--------|--------|
| Plan 历史兼容 | 历史 Plan（Others 类型）正常展示 |
| 新建订阅 | 组织订阅创建成功，功能随 Plan 启用 |
| PQ Tier 新建 | 数量/费率保存正确，Zoho 接收邮件 |
| Global SN Search | 跨组织搜索 SN 返回正确结果 |
| SN 重复标签 | 重复 SN 设备被正确标记 |
| Device Model 文档 | 型号文档上传、列表、下载正常 |
| 排放因子 | 全球因子配置后 Carbon Model 计算正确 |
