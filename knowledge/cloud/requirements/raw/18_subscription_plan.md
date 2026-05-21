# Subscription Plan 订阅计划

## 概述

AcuCloud 采用分级订阅模式，不同 Plan 对应不同的功能权限和数据采集能力。

---

## Plan 类型

| Plan | 定位 | 关键功能 |
|------|------|---------|
| **Free** | 免费基础版 | 基础监控，无计费无高级分析 |
| **Lite** | 轻量版 | 基础 Billing（电力） |
| **AcuBilling** | 计费专项 | 完整 Billing（含水/气） |
| **AcuPQ** | 电能质量专项 | Power Quality 分析 |
| **Plus** | 增强版 | Billing + Carbon + Utility Bills Analysis + M&V |
| **AcuEMS** | 全功能能源管理 | 全部功能（含 PQ + Carbon + Utility Bills） |
| **Others** | 历史遗留 | 全权限（历史存量账户） |

---

## 功能-Plan 对照矩阵

| 功能模块 | Free | Lite | AcuBilling | AcuPQ | Plus | AcuEMS |
|---------|------|------|-----------|-------|------|--------|
| Installation | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Dashboard | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Analysis（基础） | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Analysis（M&V） | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Billing（电） | ❌ | ✅ | ✅ | ❌ | ✅ | ✅ |
| Billing（水/气） | ❌ | ❌ | ✅ | ❌ | ✅ | ✅ |
| Power Quality | ❌ | ❌ | ❌ | ✅ | ❌ | ✅ |
| Carbon Model | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Utility Bills Analysis | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Utility Bills Report | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Submeter Bills | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Budget Tracking | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| SMS 通知 | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Billing Zoho 集成 | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |
| Report（基础） | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Report（Utility Bills） | ❌ | ❌ | ❌ | ❌ | ✅ | ✅ |

---

## Subscription Tier（订阅层级）

每个组织的 Plan 下可配置多种 Tier：

### Standard Meter Point Tier

- 配置可创建/激活的标准计量点数量
- Number of Subscription：已订阅数量
- Rate：每计量点费率

### Power Quality Meter Point Tier（第四阶段新增）

- 专门用于 Power Quality 相关计量点的订阅计费
- 默认值：0（历史订阅不受影响）
- SuperAdmin 操作：新建 → 配置数量和费率 → Zoho 验证

---

## 订阅流程

### SuperAdmin 端（全局订阅配置）

1. Super Admin → Subscription
2. 选择目标组织
3. 选择 Plan 类型
4. 配置各 Tier 数量和费率
5. 配置 Zoho 集成（如需要）
6. 提交 → 订阅激活

### Admin 端（组织订阅分配）

1. Admin → Subscription
2. 查看当前组织的订阅状态
3. 在已有 Tier 配额内分配给用户/设施

---

## 订阅 Interval Limit（ACU-820）

| Plan | 最小 Interval | 最大 Interval |
|------|-------------|-------------|
| Free | 1h | 1day |
| Lite | 15min | 1day |
| AcuBilling | 5min | 1day |
| AcuPQ | 5min | 1day |
| Plus | 5min | 1day |
| AcuEMS | 5min | 1day |

> 超出 Interval Limit 时系统提示订阅限制。

---

## 历史 Plan 兼容

- 历史系统 Plan 类型为 **Others**，具有所有权限
- 历史订阅正常展示，不被新 Plan 结构影响
- 升级 Plan 后功能解锁，降级后功能被限制（历史数据保留）

---

## 未订阅状态

当前账号组织未订阅时：
- 菜单树（`/system-menu/tree`）返回 `code: 603`
- 提示联系 `yanling.cao@accuenergy.com` 订阅
- 大部分功能在测试环境下仍可直接访问（URL 输入方式绕过菜单限制）
