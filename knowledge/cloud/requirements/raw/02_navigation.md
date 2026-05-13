# 导航结构与路由

## 左侧导航栏（侧边栏）

AcuCloud 使用固定左侧边栏导航，包含图标+文字标签。

```
├── Home                    #/home
├── Dashboards              #/dashboard/list
├── Installation            #/installation/devices   (默认落到 Devices 子标签)
├── Billing                 #/billing/accounts
├── Utility Bills [BETA]    #/utilityBills/utilityBillsAccount
├── Analysis                #/analysis/energy
├── Power Quality [BETA]    #/powerQuality
├── Carbon Model [BETA]     #/carbonModel/carbonAnalysis
├── Report                  #/report/configs/0
├── Data                    #/data/dataExports
├── Logs                    #/log
├── Admin                   #/admin/users
└── Super Admin             #/superAdmin/users
```

## 各模块子路由详情

### Installation（安装管理）
路由前缀：`#/installation/`

| 子页面 | URL |
|--------|-----|
| Facilities（设施） | `#/installation/facility` |
| Devices（设备） | `#/installation/devices` |
| Meter Points（计量点） | `#/installation/meterPoints` |
| Alerts（告警） | `#/installation/alerts` |
| Portfolio（项目集） | `#/installation/portfolio` |

### Billing（计费）
路由前缀：`#/billing/`

| 子页面 | URL |
|--------|-----|
| Accounts（账户配置） | `#/billing/accounts` |
| Auto Billing（自动计费） | `#/billing/autoBilling` |
| Manual Billing（手动计费） | `#/billing/manualBilling` |
| Previous Bills（历史账单） | `#/billing/previousBills` |
| Billing Analysis（账单分析） | `#/billing/billingAnalysis` |
| Rates（费率配置） | 通过右上角 Rates 按钮访问 |

### Utility Bills [BETA]
路由前缀：`#/utilityBills/`

| 子页面 | URL |
|--------|-----|
| Account | `#/utilityBills/utilityBillsAccount` |
| Utility Bills | `#/utilityBills/utilityBills` |
| Utility Bills Analysis | `#/utilityBills/utilityBillsAnalysis` |

### Analysis（分析）
路由前缀：`#/analysis/`

| 子页面 | URL |
|--------|-----|
| Energy（能耗） | `#/analysis/energy` |
| Heatmap（热力图） | 通过 Energy 页面 tab 切换 |
| Schedule（排程） | 通过 Energy 页面 tab 切换 |
| Annotations（标注） | 通过 Energy 页面 tab 切换 |
| M&V（测量与验证） | 通过 Energy 页面 tab 切换 |

### Carbon Model [BETA]
路由前缀：`#/carbonModel/`

| 子页面 | URL |
|--------|-----|
| Analysis（碳分析） | `#/carbonModel/carbonAnalysis` |
| Emission Factor（排放因子） | 通过 tab 切换 |

### Report（报告）
路由前缀：`#/report/`

| 子页面 | URL |
|--------|-----|
| Configs（配置） | `#/report/configs/0` |
| Generate（生成） | 通过 tab 切换 |
| Previous Reports（历史报告） | 通过 tab 切换 |

报告类型 tab：Energy Report / Billing Report / Portfolio Report / Power Quality Report [BETA]

### Data（数据）
路由前缀：`#/data/`

| 子页面 | URL |
|--------|-----|
| Data Exports | `#/data/dataExports` |
| Data Forwards | `#/data/dataForwards` |
| Data Import | `#/data/dataImport` |
| Energy Exports | `#/data/energyExports` |
| Max Min Export | `#/data/maxMinExport` |
| Data Edit | `#/data/dataEdit` |
| Manual VEE | `#/data/manualVee` |

### Admin（管理员）
路由前缀：`#/admin/`

| 子页面 | URL |
|--------|-----|
| Users（用户） | `#/admin/users` |
| Organization（组织） | `#/admin/organization` |
| Tenants（租户） | `#/admin/tenants` |
| Subscription（订阅） | `#/admin/subscription` |
| Customization（自定义） | `#/admin/customization` |

### Super Admin（超级管理员）
路由前缀：`#/superAdmin/`

| 子页面 | URL |
|--------|-----|
| Users（用户） | `#/superAdmin/users` |
| Parameter（参数） | `#/superAdmin/parameter` |
| Device Model（设备型号） | `#/superAdmin/deviceModel` |
| Organization（组织） | `#/superAdmin/organization` |
| Subscription（订阅） | `#/superAdmin/subscription` |
| Plan（计划） | `#/superAdmin/plan` |
| Release Center（发布中心） | `#/superAdmin/releaseCenter` |
| Global Emission Factors | `#/superAdmin/globalEmissionFactors` |
| Global Device Search | `#/superAdmin/globalDeviceSearch` |
| Wiring Check | `#/superAdmin/wiringCheck` |

## 顶栏功能

| 控件 | 功能 |
|------|------|
| 地球图标 + English | 语言切换（英/西/法等） |
| Home 图标 | 快速跳转首页 |
| Help 图标 | 帮助文档 |
| 邮件图标 | 消息通知 |
| 人像图标 | 用户菜单（个人信息/退出） |
| Organization View 开关 | 切换组织视图 / 设施视图 |
| 右上组织选择器 | 切换当前组织（多组织场景） |

## 404 路由（guessed 路由无效）

以下路由在本账号下返回 404，说明路由名与猜测不符或未授权：
- `#/dashboard/list`（Dashboards）
- `#/installation/facility`（Installation/Facility 子页）
- `#/log`（Logs）
