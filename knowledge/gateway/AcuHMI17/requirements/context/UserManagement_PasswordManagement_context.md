# User Management / Password Management — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 列表路由 | `/#/userManagement/passwordManagement/list`（`passwordManagementList`） |
| 编辑路由 | `/#/userManagement/passwordManagement/edit`（`passwordManagementEdit`） |
| 面包屑 | AcuHMI-1-7 / User Management / Password Management |
| 顶级模块 | User Management |

## 2. 页面用途

管理员为各用户重置/修改密码。

## 3. 列表页 (list)

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| 用户表 | table | 列: Username/Role/Register Date/Expiration Date/Last Login Time/Status/Action | 只读+操作 |
| 行 Action | button | 行末图标按钮 | 进入 edit（修改该用户密码） |
| Username 排序 | columnheader | `getByRole('columnheader',{name:'Username'})` | 排序 |

## 4. 编辑页 (edit)

- 修改所选用户密码：含 New Password / Repeat Password 字段 + Save/Cancel（受 Password Policy 约束）。经列表行 Action 进入。
- **字段定位坑**：密码字段 label 实际文本带星号 `Password*`，`get_by_label("Password", exact=True)` 在部分表单匹配不到，优先用 `get_by_placeholder("Enter Password")` / `get_by_placeholder("Enter Repeat Password")`。
- 目标用户行精确匹配：`page.locator("tbody tr")` 按 `td:first-child` 文本**精确等于**用户名（避免匹配到含该字样的其它行，如 admin）。

### 4.1 改「自身」密码 → Current User Password 验证弹窗（2026-07 实测）

改**自己**的密码（admin 或非admin 皆然）：进入 edit 页会弹出验证弹窗 `.password-verify-dialog`（标题 "Current User Password"），内含**无 label** 的 `input[type=password]`（placeholder="Please input"）+ Cancel/Confirm。

- 正确流程：先在弹窗输入**当前密码** → 点弹窗内 Confirm → 再填 Password/Repeat → Save。
- ⚠️ **不先确认弹窗就点 Save**：弹 "Unauthenticated user, please log in!"，改密**不生效**。
- ⚠️ 弹窗遮罩 `.el-overlay` 会拦截 Save 的 Playwright actionability：Confirm 后需等所有 `.el-overlay` `display:none`/`offsetParent===null` 再点 Save（必要时 `page.evaluate` JS click 绕过）。
- 弹窗出现与否与 session 验证状态有关：新建独立 context 登录后进 edit 页会弹；同一已登录 session 的 page 可能不弹（已缓存验证态）——用例需兼容两种。

选择器速查：
| 元素 | 定位 |
|------|------|
| 验证弹窗 | `.password-verify-dialog`（或 `.el-overlay [role='dialog']` 含 `input[type=password]`）|
| 弹窗密码输入 | 弹窗内 `input[type=password]`（placeholder="Please input"，无 label）|
| 弹窗 Confirm/Cancel | 弹窗内 `button` 文本 == `Confirm` / `Cancel` |
| New/Repeat 密码 | `get_by_placeholder("Enter Password")` / `get_by_placeholder("Enter Repeat Password")` |

## 5. 自动化测试要点

- 从列表进入某用户 edit → 新密码两次一致 + 满足 Password Policy → Save。
- 密码历史限制（不可重用前 N 个，见 Password Policy 联动）。
- **改密类用例目标新密码必须固定硬编码**（禁止变化值），便于失败时人工恢复。
- **改 admin 本账号密码为高危**（易锁死设备 + 触发安全沙箱），建议手工验证、不自动化。

### 5.1 改密后是否自动登出（2026-07 实测，与手工用例预期有出入）

| 操作 | 验证弹窗 | 改密后自动登出 |
|------|:-------:|:-------------:|
| admin 改**自身**密码 | 有（新 context）| ✅ 跳 `/#/login` |
| admin 改**他人**密码 | 无 | ❌ 会话保持，停留原页 |
| 非admin 改**自己**密码 | 有 | ❌ 停留列表页 |
| 非admin 改**他人**密码 | 有（填**操作者自己**的当前密码，非被改对象）| —（不登出）|

对应用例：`projects/AcuHMI_1_7/tests/ui/usermanagement/passwordchange/`（case1_04/05/06 已自动化通过；case1_03 admin 改自身**不做自动化**，无脚本，仅手工执行——改本账号高危 + 触发安全沙箱）。

## 6. 机器可解析摘要

```json
{
  "routes": {"list":"/userManagement/passwordManagement/list","edit":"/userManagement/passwordManagement/edit"},
  "name": "passwordManagement",
  "title": "Password Management",
  "module": "User Management",
  "list_columns": ["Username","Role","Register Date","Expiration Date","Last Login Time","Status","Action"],
  "edit_form": ["New Password","Repeat Password"],
  "edit_form_selectors": {"new":"getByPlaceholder('Enter Password')","repeat":"getByPlaceholder('Enter Repeat Password')"},
  "self_change_verify_dialog": {"selector":".password-verify-dialog","title":"Current User Password","input":"input[type=password] placeholder='Please input'","buttons":["Confirm","Cancel"],"note":"改自身密码时出现；不确认直接Save会报Unauthenticated且不生效；overlay拦截Save需等overlay消失或JS click"},
  "auto_logout_after_change": {"admin_self":true,"admin_other":false,"nonadmin_self":false,"nonadmin_other":false},
  "row_actions": ["change password"]
}
```
