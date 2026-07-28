# User Management / Password Policy — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/userManagement/passwordPolicy` |
| 路由名 | `passwordPolicy` |
| 面包屑 | AcuHMI-1-7 / User Management / Password Policy |
| 顶级模块 | User Management |

## 2. 页面用途

配置全局密码策略（复杂度、历史、有效期、锁定规则）。

## 3. 交互元素清单 / 表单字段

| 字段 | 类型 | 定位策略 1 | 默认 | 范围/说明 |
|------|------|-----------|------|-----------|
| Upper and Lower Case | checkbox(Required) | `getByRole('checkbox',{name:'Required'})`（第1个/按分组标题定位） | 勾选 | 必须含大小写 |
| Numbers and Letters | checkbox(Required) | 分组 "Numbers and Letters" 内 checkbox | 勾选 | 必须含字母+数字 |
| Special Characters | checkbox(Required) | 分组 "Special Characters" 内 checkbox | 勾选 | 必须含特殊字符 (!@#$%^) |
| Password History | textbox | `getByRole('textbox',{name:'Enter Password History'})` | 1 | **1–32**，1=无限制（不可重用前 N 个） |
| Minimum Password Age | textbox | `getByRole('textbox',{name:'Enter Minimum Password Age'})` | 0 | **0–90** days，0=无限制 |
| Password Expires | textbox | `getByRole('textbox',{name:'Enter Password Expires'})` | 0 | **0–90** days，0=永不过期 |
| Minimum Password Length | textbox | `getByRole('textbox',{name:'Enter Minimum Password Length'})` | 8 | **8–64** |
| Grace Period | textbox | `getByRole('textbox',{name:'Enter Grace Period'})` | 65535 | **0–65535** days，0=无宽限 |
| Maximum Failed Attempts | textbox | `getByRole('textbox',{name:'Enter Maximum Failed Attempts'})` | 0 | **0–30**，0=永不锁定 |
| Failed Login Attempt Window | textbox | `getByRole('textbox',{name:'Enter Failed Login Attempt Window'})` | 0 | **0–86400** seconds |
| Failed Login Wait | textbox | `getByRole('textbox',{name:'Enter Failed Login Wait'})` | 0 | **0–86400** seconds，0=不自动解锁 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 自动化测试要点

- 每个数值字段的边界与"0/1=无限制"语义校验（丰富的边界值用例源）。
- 3 个复杂度 Required 复选框各自独立切换。
- 策略变更影响 User Configuration 新增/改密的密码校验（联动）。

## 5. 机器可解析摘要

```json
{
  "route": "/userManagement/passwordPolicy",
  "name": "passwordPolicy",
  "title": "Password Policy",
  "module": "User Management",
  "fields": {
    "Upper and Lower Case": {"type":"checkbox","default":true},
    "Numbers and Letters": {"type":"checkbox","default":true},
    "Special Characters": {"type":"checkbox","default":true},
    "Password History": {"type":"text","range":[1,32],"note":"1=none"},
    "Minimum Password Age": {"type":"text","range":[0,90],"unit":"days"},
    "Password Expires": {"type":"text","range":[0,90],"unit":"days","note":"0=never"},
    "Minimum Password Length": {"type":"text","range":[8,64]},
    "Grace Period": {"type":"text","range":[0,65535],"unit":"days"},
    "Maximum Failed Attempts": {"type":"text","range":[0,30],"note":"0=never lockout"},
    "Failed Login Attempt Window": {"type":"text","range":[0,86400],"unit":"seconds"},
    "Failed Login Wait": {"type":"text","range":[0,86400],"unit":"seconds"}
  },
  "buttons": ["Save"]
}
```
