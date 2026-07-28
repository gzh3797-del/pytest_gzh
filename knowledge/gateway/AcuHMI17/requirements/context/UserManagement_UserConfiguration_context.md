# User Management / User Configuration — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 列表路由 | `/#/userManagement/userConfiguration/list`（`userConfigurationList`） |
| 新增路由 | `/#/userManagement/userConfiguration/add`（`userConfigurationAdd`） |
| 编辑路由 | `/#/userManagement/userConfiguration/edit`（`userConfigurationEdit`） |
| 面包屑 | AcuHMI-1-7 / User Management / User Configuration |
| 顶级模块 | User Management |

## 2. 页面用途

管理系统用户：列出、新增、编辑、锁定/解锁、删除用户及其角色。

## 3. 列表页 (list)

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Add User | button | `getByRole('button',{name:'Add User'})` | 进入 add 页 |
| 用户表 | table | 列: Username/Role/Register Date/Expiration Date/Last Login Time/Status/Lock/Action | — |
| 行 Lock | button | `getByRole('row',{name:/test_user/}).getByRole('button',{name:'Lock'})` | 锁定用户（admin 自身无 Lock） |
| 行 Action | button×2 | 行末图标 | 编辑/删除 |
| Username 排序 | columnheader | `getByRole('columnheader',{name:'Username'})` | 排序 |

- 列语义：Status(Active/…)、Expiration Date(no restrict/日期)、Role(admin/view/自定义角色)。

## 4. 新增/编辑页 (add/edit)

| 字段/元素 | 类型 | 定位策略 1 | 校验/说明 |
|-----------|------|-----------|-----------|
| Username | textbox | `getByRole('textbox',{name:'Username'})` | 必填(*)，唯一 |
| Password | textbox | `getByRole('textbox',{name:'Password'})` | 必填(*)，受 Password Policy 约束 |
| Repeat Password | textbox | `getByRole('textbox',{name:'Repeat Password'})` | 必填(*)，须与 Password 一致 |
| Role | combobox | `getByRole('combobox',{name:'Role'})` | 必选(*)，--Select Role-- |
| Override Password Policy | checkbox | `getByRole('checkbox',{name:'Override Password Policy'})` | 未勾 |
| Multiple Login | checkbox | `getByRole('checkbox',{name:'Multiple Login'})` | 默认勾选 |
| Override Password Expire | checkbox | `getByRole('checkbox',{name:'Override Password Expire'})` | 未勾 |
| Save / Cancel | button | `getByRole('button',{name:'Save'})` / `{name:'Cancel'}` | 提交/取消 |

## 5. 页面状态与分支

| 状态 | 说明 |
|------|------|
| admin 自身行 | 无 Lock 按钮（不可锁自己） |
| 密码不一致 | Password≠Repeat 应报错 |
| edit 页 | 预填现有用户数据（经列表 Action 编辑进入） |

## 6. 自动化测试要点

- 新增：用户名唯一、两次密码一致、密码策略校验（除非 Override）、角色必选。
- Lock/Unlock 状态切换；删除二次确认。
- 三个 checkbox 各自独立行为（原子操作）。

## 7. 机器可解析摘要

```json
{
  "routes": {"list":"/userManagement/userConfiguration/list","add":"/userManagement/userConfiguration/add","edit":"/userManagement/userConfiguration/edit"},
  "name": "userConfiguration",
  "title": "User Configuration",
  "module": "User Management",
  "list_columns": ["Username","Role","Register Date","Expiration Date","Last Login Time","Status","Lock","Action"],
  "form_fields": {
    "Username": {"type":"text","required":true,"unique":true},
    "Password": {"type":"password","required":true},
    "Repeat Password": {"type":"password","required":true,"match":"Password"},
    "Role": {"type":"select","required":true},
    "Override Password Policy": {"type":"checkbox"},
    "Multiple Login": {"type":"checkbox","default":true},
    "Override Password Expire": {"type":"checkbox"}
  },
  "row_actions": ["Lock","edit","delete"]
}
```
