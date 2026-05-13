# AcuHMI-1-7 User Management 页面结构文档

> 设备地址：`https://192.168.2.199`  
> 路径：顶部导航 **AcuHMI-1-7** 按钮 → 左侧菜单 **User Management**  
> 最后更新时间：2026-05-08（依据页面截图校准）  
> 自动化规则见：[autotest_generativerule.md](autotest_generativerule.md)

---

## 一、整体导航结构

### 1.1 两套左侧菜单

登录后，顶部 Header 包含两个导航入口，对应不同的左侧菜单：

| 入口 | 左侧菜单内容 |
|------|-------------|
| **Devices** 按钮（全局视图） | Dashboard / Physical Devices / Virtual Devices / Web Devices / Alarm / Data Log |
| **AcuHMI-1-7** 按钮（设备上下文） | System Settings / Templates / Protocols / Maintenance / Diagnostics / **User Management** / Firmware Update |

> **User Management 在 AcuHMI-1-7 设备上下文中**，进入方式：点击顶部 Header 的 **AcuHMI-1-7** 按钮 → 左侧点击 **User Management**。

### 1.2 User Management 内部 Tab 结构

进入 User Management 后，顶部显示 5 个 Tab：

```
General | User Configuration | Role Configuration | Password Policy | Password Management
```

> **Tab 折叠行为（Element UI 响应式）**：当前激活的 Tab 若不是 Password Management，末位会折叠为 `...More`，Password Management 隐藏其中；当 General 或 Password Policy 为当前激活 Tab 时，5 个 Tab 全部直接可见。  
> **自动化定位**：通过 `page.get_by_role("menuitem", name="Password Management")` 点击，无需关注是否折叠。

### 1.3 完整导航路径

```
登录页 (/#/login)
└── Header
    ├── Devices 按钮 → 全局 Dashboard（左侧：Physical Devices / Virtual Devices 等）
    └── AcuHMI-1-7 按钮 → 设备上下文（左侧：System Settings / Templates / ... / User Management）
        └── User Management（左侧菜单点击）
            ├── Tab: General               (/#/userManagement/general)
            ├── Tab: User Configuration    (/#/userManagement/userConfiguration/list)
            ├── Tab: Role Configuration    (/#/userManagement/roleConfiguration)
            ├── Tab: Password Policy       (/#/userManagement/passwordPolicy)
            └── Tab: Password Management   (/#/userManagement/passwordManagement)
```

---

## 二、登录页 Login

| 元素 | 类型 | Playwright 定位建议 |
|------|------|-------------------|
| Enter User Name | 文本框 | `get_by_role("textbox", name="Enter User Name")` |
| Enter Password | 密码框 | `get_by_role("textbox", name="Enter Password")` |
| Sign In | 按钮 | `get_by_role("button", name="Sign In")` |
| Forgot password | 链接/按钮 | `get_by_text("Forgot password")` |

### 2.1 登录后行为

- 若使用默认密码登录，弹出"修改默认密码"提示框：
  - `Cancel`：关闭弹框，保持当前密码，进入主页（不退出登录）
  - `Yes, continue`（**admin 用户**）：跳转至修改密码页面
  - `Yes, continue`（**view / 非 admin 用户**）：直接退出登录

- 非 admin 用户提示文本：
  > "You are using the default password. Click Continue to log out.  
  > Please contact your administrator to change your password, then log in again."

### 2.2 Forgot Password 行为

| 操作用户 | 行为 |
|----------|------|
| admin | 弹框显示当前时间和 SN 号，需借助外部工具根据时间+SN 生成临时密码 |
| 非 admin | 弹框提示 `"Please contact your administrator for assistance"` |

### 2.3 EULA 弹框触发条件

| 场景 | 是否触发 |
|------|---------|
| 新建用户首次登录 | **是** |
| 已接受 EULA 的用户再次登录 | 否 |
| 恢复出厂后所有用户重新登录 | **是** |
| 固件升级且 EULA 版本变更后登录 | **是** |

---

## 三、User Management — General（常规设置）

**路径**：左侧 User Management → **General** Tab  
**截图参考**：`explore/01_user_management_landing.png`

| 元素 | 类型 | UI 显示 | 说明 |
|------|------|---------|------|
| Session Timeout * | 数字输入框 | 默认值 `10`，单位 `minutes` | 必填（* 标注）；hint：`0 for never timeout` / `Range: 0 - 60` |
| Save | 按钮 | 蓝绿色 | 保存成功提示 `"User management general configuration saved"` |

**业务逻辑**：超时后系统自动登出所有已登录用户，下次需重新登录。有效范围 0–60；0 = 永不超时；61+ 保存失败。

**Playwright 定位**：
```python
page.get_by_label("Session Timeout")   # 或 get_by_role("spinbutton")
page.get_by_role("button", name="Save")
```

---

## 四、User Management — User Configuration（用户配置）

**路径**：左侧 User Management → **User Configuration** Tab  
**截图参考**：`explore/uc_main.png`

### 4.1 用户列表列定义

| 列名 | 说明 |
|------|------|
| Username ↕ | 用户名（可排序） |
| Role | 所属角色（admin / view / 自定义） |
| Register Date | 注册时间 |
| Expiration Date | 账号到期时间（`no restrict` = 永不过期） |
| Last Login Time | 最近一次登录时间（从未登录则显示 `-`） |
| Status | 账号状态（Active / Locked / Disabled） |
| Lock | 仅对**非 admin 用户**显示红色 `Lock` 按钮；admin 行**无此按钮** |
| Action | 见下表 |

### 4.2 Action 列按钮规则（按行类型区分）

| 行类型 | 编辑按钮 | 删除按钮 | Lock 按钮 |
|--------|---------|---------|----------|
| **非 admin 用户行** | ✅ 有（方形编辑图标） | ✅ 有（方形删除图标） | ✅ 有（红色 `Lock`） |
| **admin 用户行** | ✅ 有 | ❌ 无 | ❌ 无 |

> **自动化注意**：定位删除按钮时须过滤行，避免误操作 admin 行：
> ```python
> row = page.locator("tbody").get_by_role("row").filter(has_text=username)
> row.get_by_role("button").last.click()   # 最后一个按钮 = 删除
> ```

### 4.3 Add User 按钮

```python
page.get_by_role("button", name="Add User")
```

### 4.4 Add User / Edit User 表单字段

> **注意**：Add User 表单是 **drawer/panel 结构，非标准 ARIA dialog**，须用页面级 `get_by_label` 定位，不可套 `get_by_role("dialog")`。

| 字段标签（页面实际） | 类型 | 必填 | 说明 |
|---------------------|------|------|------|
| Username | 文本框 | ✅ | 有效长度 1–40 字符；超 40 位保存失败 |
| Password | 密码框 | ✅ | 有效长度 6–128 字符（下限由 Password Policy 的 Minimum Password Length 决定） |
| Repeat Password | 密码框 | ✅ | 须与 Password 一致 |
| Role | 下拉选择（`--Select Role--`） | ✅ | 可选 admin / view / 自定义角色 |
| Expiration Date | 日期选择 | ❌ | 留空 = `no restrict`（永不过期） |
| Override Password Policy | 复选框 | ❌ | 勾选后该用户不受全局密码复杂度约束 |
| Override Password Expire | 复选框 | ❌ | 勾选后可覆盖全局密码过期策略 |
| Multiple Login | 复选框 | ❌ | 勾选后允许多浏览器窗口同时登录 |
| Save | 按钮 | — | 成功提示 `"Add success"` |
| Cancel | 按钮 | — | 关闭表单 |

**Role 下拉 Playwright 定位**（Element UI el-select，combobox 被 span 拦截）：
```python
page.get_by_text("--Select Role--", exact=True).click()
page.get_by_role("option", name="view").click()
```

### 4.5 删除确认弹框

- 按钮：`Yes, continue`（确认删除）、`Cancel`（取消）
- 定位：`page.get_by_role("button", name="Yes, continue")`

### 4.6 密码修改权限说明

| 操作者 | 修改对象 | Password Management 表单是否显示 Current Password |
|--------|----------|--------------------------------------------------|
| admin | 任意用户（含自身） | **不显示** |
| 非 admin | 自身 | **显示** |

---

## 五、User Management — Role Configuration（角色配置）

**路径**：左侧 User Management → **Role Configuration** Tab  
**截图参考**：`explore/rc_main.png`

### 5.1 角色列表列定义（按截图列顺序）

| 列名 | Playwright 列索引参考 | 说明 |
|------|----------------------|------|
| Role Name ↕ | 0 | 角色名称（可排序） |
| User | 1 | 用户管理模块权限 |
| Device | 2 | 设备模块权限 |
| Data Log | 3 | 数据日志权限 |
| Alarm Log | 4 | 告警日志权限 |
| System Settings | 5 | 系统设置权限 |
| Protocol | 6 | 协议权限 |
| Maintenance | 7 | 维护权限 |
| Diagnostics | 8 | 诊断权限 |
| Firmware Update | 9 | 固件升级权限（表头截断显示为 `Firmware U`） |
| Action | 10 | 见下表 |

### 5.2 Action 列按钮规则（按行类型区分）

| 行类型 | 编辑按钮 | 删除按钮（红色） |
|--------|---------|----------------|
| **自定义角色行**（如 view） | ✅ 有 | ✅ 有 |
| **admin 内置角色行** | ❌ 无 | ❌ 无 |

> **截图确认**：admin 行 Action 列为空，view 行有编辑 + 红色删除按钮。

### 5.3 权限选项值

| 选项 | 含义 |
|------|------|
| `none` | 无权限（用户无法访问该模块） |
| `view` | 只读权限 |
| `edit` | 可读写权限 |

### 5.4 内置角色（截图实测）

| 角色 | 所有模块权限 | 可编辑 | 可删除 |
|------|------------|--------|--------|
| admin | edit | ❌ | ❌ |
| view | view | ✅ | ✅ |

### 5.5 Add Role 按钮

```python
page.get_by_role("button", name="Add Role")
```

**角色名称规则**：仅支持数字、字母、下划线、空格；最大 40 字符；已有用户关联的角色无法删除。

---

## 六、User Management — Password Policy（密码策略）

**路径**：左侧 User Management → **Password Policy** Tab  
**截图参考**：`explore/pp_main.png`

### 6.1 复杂度要求（Complexity Requirements）

| 字段标签 | 类型 | 默认状态 | UI hint 文本 |
|---------|------|---------|-------------|
| Upper and Lower Case | 复选框 | ✅ Required | `If required, password must contain both upper and lower case characters` |
| Numbers and Letters | 复选框 | ✅ Required | `If required, password must contain at least an alphabet and a number` |
| Special Characters | 复选框 | ✅ Required | `If required, password must contain at least one non-alphanumeric character e.g. '@#$%'` |

### 6.2 密码历史与生命周期

| 字段标签 | 默认值 | 有效范围 | UI hint 文本 | 自动化注意 |
|---------|-------|---------|-------------|----------|
| Password History | 1 | **1–32** | `User cannot reuse any of their previous N passwords. 0 means no restriction` | ⚠️ **UI hint 有误**：hint 说 0 = 无限制，但**实测 0 为无效值，保存失败**；自动化以实测为准 |
| Minimum Password Age | 0 | **0–90** | `User must use a password for this many days before changing it again. 0 means no restriction` | 0 = 无限制 |
| Password Expiries | 0 | **0–90** | `Days until a user's password expires. 0 means never expires` | 0 = 永不过期 |
| Minimum Password Length | 6 | **6–64** | `Password must be at least 6 characters` | — |
| Grace Period | 0 | **0–65535** | `After expiration, user has this many days to login and change their password (must change at login before being prevented from logging in). 0 means no grace - users will be blocked from logging in as soon as the password expires` | 0 = 立即锁定 |

### 6.3 登录失败策略

| 字段标签（UI 实际） | 默认值 | 单位 | UI hint 文本 | 有效范围 |
|-------------------|-------|------|-------------|---------|
| Maximum Failed Attempts | 0 | 次 | （页面截图中位于 Grace Period 与 Failed Login Attempt Window 之间） | **0–30**；0 = 永不锁定 |
| Failed Login Attempt Window | 0 | seconds | `Number of seconds after which the current count of failed attempts is reset. 0 means never lockout` | **0–86400**；0 = 不统计窗口 |
| Failed Login Wait | 0 | seconds | `After a lockout due to getting Max Failed attempts within the Failed Login Attempt Window, the account will automatically be removed after N seconds. 0 means never auto lock` | **0–86400**；0 = 永不自动解锁 |

**逻辑说明**：
- `Maximum Failed Attempts = 0`：永不锁定
- `Failed Login Attempt Window = 0`：不设时间窗口，等效永不锁定
- `Failed Login Wait = 0`：被锁后不自动解锁，需管理员手动解锁
- 三字段需同时配置才能触发账户锁定

> 所有字段修改后点击页面底部 **Save** 生效。

---

## 七、User Management — Password Management（密码管理）

**路径**：左侧 User Management → **Password Management** Tab（有时折叠在 `...More` 内）  
**截图参考**：`explore/pm_main.png`

### 7.1 用户列表列定义

| 列名 | 说明 |
|------|------|
| Username ↕ | 用户名（可排序） |
| Role | 所属角色 |
| Register Date | 注册时间 |
| Expiration Date | 到期时间（`no restrict` = 永不过期） |
| Last Login Time | 最近登录时间（从未登录则为空） |
| Status | 账号状态 |
| Action | 仅有 ✏️ 编辑图标（蓝绿色）；**无 Lock 列、无删除按钮** |

> **与 User Configuration 的区别**：Password Management 无 Lock 列，且 admin 行也有编辑按钮（用于改密）。

### 7.2 编辑密码表单字段（页面实际标签）

| 字段标签 | 类型 | 说明 |
|---------|------|------|
| Password | 密码框 | 新密码，须符合当前 Password Policy |
| Repeat Password | 密码框 | 须与 Password 一致 |
| Current Password | 密码框 | **仅当非 admin 用户修改自身密码时显示**；admin 修改任意用户时此字段不出现 |
| Save | 按钮 | 成功提示 `"User password changed"` |
| Cancel | 按钮 | 关闭表单 |

**Playwright 定位**：
```python
# 必须加 exact=True，否则 "Password" 会匹配到 "Repeat Password"
page.get_by_label("Password", exact=True).fill(new_pwd)
page.get_by_label("Repeat Password", exact=True).fill(new_pwd)
page.get_by_role("button", name="Save").click()
```

---

## 八、页面 URL 映射

| 页面 | URL Hash |
|------|----------|
| 登录 | `/#/login` |
| 全局 Dashboard | `/#/dashboard` 或 `/#/` |
| User Management / General | `/#/userManagement/general` |
| User Configuration | `/#/userManagement/userConfiguration/list` |
| Role Configuration | `/#/userManagement/roleConfiguration` |
| Password Policy | `/#/userManagement/passwordPolicy` |
| Password Management | `/#/userManagement/passwordManagement` |

> ⚠️ **以上 URL 不可通过 `page.goto()` 直接访问**（SPA 会被重定向），必须通过 UI 点击导航。  
> 标准导航封装见 [autotest_generativerule.md § 3.1](autotest_generativerule.md)。

---

## 九、恢复出厂设置

**路径**：AcuHMI-1-7 设备上下文 → 左侧 System Settings → Configuration Management → Reset

**行为**：
1. 点击 Reset 确认后，设备重启，所有配置恢复出厂默认
2. 重启后使用默认密码（`Admin@AABBCC` / `View@AABBCC`）登录，EULA 弹框重新触发
3. 登录后提示修改默认密码

> **自动化注意**：涉及恢复出厂的用例，代码可生成但标注 `@pytest.mark.skip(reason="涉及恢复出厂，需手动执行")`，不纳入 CI 自动运行。

---

## 十、EULA 功能说明

| 场景 | EULA 是否弹出 |
|------|--------------|
| 新建用户首次登录 | **是** |
| 已接受 EULA 的用户正常登录 | 否 |
| 恢复出厂后所有用户重新登录 | **是** |
| 固件升级且 EULA 版本变更后登录 | **是** |

**EULA 弹框操作**：
- 接受：登录成功，后续不再弹出
- 不接受（按钮文本待确认）：无法登录，下次仍弹出

---

## 十一、默认密码说明

| 项目 | 值 |
|------|-----|
| admin 默认密码格式 | `Admin@AABBCC`（AABBCC = 序列号后 6 位） |
| view 默认密码格式 | `View@AABBCC` |
| 测试环境序列号后 6 位 | `110001`（测试环境密码：`Admin@110001`） |

登录后弹框行为：
- `Cancel`：关闭弹框，可正常使用，下次登录仍提示
- `Yes, continue`（admin）：跳转修改密码页
- `Yes, continue`（view / 非 admin）：退出登录

---

## 十二、澄清汇总（2026-05-07）

| 编号 | 问题 | 用户答复 | 影响章节 |
|------|------|----------|---------|
| 1 | ARM-XXL 设备名称 | 均为笔误，应为 AcuHMI | 全文 |
| 2 | Role Configuration 是否有 `none` 权限选项 | **有** | 第五节 |
| 3 | Password History=0 是有效值还是无效值 | **无效值**，0 保存失败（注意 UI hint 描述有误） | 第六节 |
| 4 | `Maximum Failed Attempts` UI 字段名称 | 与用例一致，即 `Maximum Failed Attempts` | 第六节 |
| 5 | `Failed Login Attempt Window=0` 的含义 | 不设统计窗口，永不锁定 | 第六节 |
| 6 | Add User 弹框是否有 `Multiple Login` 字段 | **有** | 第四节 |
| 7 | Add User 弹框是否有覆盖密码过期字段 | **有**，字段名 `Override Password Expire` | 第四节 |
| 8 | EULA 触发条件 | 新建用户首次登录；恢复出厂/升级后重新登录 | 第十节 |
| 9 | 恢复出厂设置路径 | System Settings → Configuration Management → Reset | 第九节 |
| 10 | view 用户点击 "Yes, continue" 的行为 | **退出登录** | 第二节 |
| 11 | 非 admin 用户 Forgot password 弹框文本 | `"Please contact your administrator for assistance"` | 第二节 |
| 12 | 密码策略复杂度测试中未给出具体密码值的用例 | 密码符合策略规则即可随机生成，注释中备注清楚 | [autotest_generativerule.md § 2.6](autotest_generativerule.md) |
| 13 | case11_01（验证默认密码功能） | **不实现自动化** | — |
| 14 | case12_01/03/04/05/06（临时密码相关） | **不实现自动化**（外部工具依赖） | — |
| 15 | case01_19/20（重新启动权限） | **无效用例**，不实现自动化 | — |
| 16 | 时间依赖用例（等待 50/90 天、30/60 分钟） | **不实现自动化** | [autotest_generativerule.md § 2.6](autotest_generativerule.md) |
| 17 | case03_2（Minimum Password Age=1，1天内不可修改）| 可通过修改系统时间实现，**纳入自动化** | — |
