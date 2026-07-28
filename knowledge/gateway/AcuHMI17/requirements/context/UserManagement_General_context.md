# User Management / General — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/userManagement/general` |
| 路由名 | `userManagementGeneral` |
| 面包屑 | AcuHMI-1-7 / User Management / General |
| 顶级模块 | User Management |

## 2. 页面用途

用户管理通用设置（会话超时）。User Management 二级 tab：General / User Configuration / Role Configuration / Password Policy / Password Management。

## 3. 交互元素清单 / 表单字段

| 元素 | 类型 | 定位策略 1 | 默认 | 校验 |
|------|------|-----------|------|------|
| Session Timeout | textbox | `getByRole('textbox',{name:'Session Timeout'})` | 0 | 必填(*)，范围 **0–60** 分钟，0=永不超时 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 自动化测试要点

- 超时范围校验（0/60 边界、61 越界、非数字）；0=never 语义断言。

## 5. 机器可解析摘要

```json
{
  "route": "/userManagement/general",
  "name": "userManagementGeneral",
  "title": "General",
  "module": "User Management",
  "fields": {"Session Timeout": {"type":"text","range":[0,60],"unit":"minutes","note":"0=never"}},
  "buttons": ["Save"],
  "sub_tabs": ["General","User Configuration","Role Configuration","Password Policy","Password Management"]
}
```
