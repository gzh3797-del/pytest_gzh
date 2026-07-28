# User Management / Role Configuration — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 列表路由 | `/#/userManagement/roleConfiguration/list`（`roleConfigurationList`） |
| 新增路由 | `/#/userManagement/roleConfiguration/add`（`roleConfigurationAdd`） |
| 编辑路由 | `/#/userManagement/roleConfiguration/edit`（`roleConfigurationEdit`） |
| 面包屑 | AcuHMI-1-7 / User Management / Role Configuration |
| 顶级模块 | User Management |

## 2. 页面用途

定义角色及其对各功能模块的权限级别（view / edit / 无）。

## 3. 列表页 (list)

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Add Role | button | `getByRole('button',{name:'Add Role'})` | 进入 add 页 |
| 角色权限表 | table | 见下列 | 每行一角色 |
| 行 Action | button×2 | 行末图标 | 编辑/删除（admin 内置角色无 Action） |

**权限列**：Role Name / User / Device / Data Log / Alarm Log / System Settings / Protocol / Maintenance / Diagnostics / Firmware Update。单元格值为权限级别：`view` / `edit`（或无权限）。

- 示例：view 角色所有模块=view；admin 角色所有模块=edit（内置，不可删）。

## 4. 新增/编辑页 (add/edit)

- 输入角色名 + 为上述每个功能模块选择权限级别（view/edit/none，通常为下拉或单选/复选）。
- Save/Cancel。
- edit 页经列表 Action 进入，预填现有角色权限。

## 5. 自动化测试要点

- 新增角色：名称唯一 + 各模块权限配置组合。
- 内置 admin 角色不可编辑/删除（负向用例）。
- 角色变更影响用户可访问范围（与 User Configuration、菜单可见性联动）。

## 6. 机器可解析摘要

```json
{
  "routes": {"list":"/userManagement/roleConfiguration/list","add":"/userManagement/roleConfiguration/add","edit":"/userManagement/roleConfiguration/edit"},
  "name": "roleConfiguration",
  "title": "Role Configuration",
  "module": "User Management",
  "permission_modules": ["User","Device","Data Log","Alarm Log","System Settings","Protocol","Maintenance","Diagnostics","Firmware Update"],
  "permission_levels": ["view","edit","none"],
  "buttons": ["Add Role"],
  "row_actions": ["edit","delete"],
  "builtin_roles": ["admin(edit all, locked)","view(view all)"]
}
```
