# System Settings / Access Control — 页面上下文

> 路由名 `whitelist`，UI 显示 **Access Control**（IP 白名单）。

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/systemSettings/whitelist` |
| 路由名 | `whitelist` |
| 面包屑 | AcuHMI-1-7 / System Settings / Access Control |
| 顶级模块 | System Settings |

## 2. 页面用途

IP 允许列表（白名单）：启用后仅列表内 IP/IP 段可访问网关。**条件分支页**（Disable→仅开关；Enable→显示列表+新增）。

## 3. 交互元素清单（主页面）

| 元素 | 类型 | 定位策略 1 | 默认 | 说明 |
|------|------|-----------|------|------|
| IP Allow List Enable | radiogroup | `getByRole('radiogroup',{name:'IP Allow List Enable'})` | Disable | 必选(*) |
| Add Allow List | button | `getByRole('button',{name:'Add Allow List'})` | — | 打开新增弹框（Enable 后可见） |
| 列表表格 | table | 列: No/Description/From IP/To IP/Action | No Data | Action 含编辑/删除 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 弹框：New Allow List ★

| 元素 | 类型 | 定位策略 | 说明 |
|------|------|----------|------|
| 标题 | heading | `getByRole('heading',{name:'New Allow List'})` | — |
| IP Range | radiogroup | `getByRole('radiogroup',{name:'IP Range'})` | Yes/No（默认 Yes）；**No→仅单一 IP Address\*（From/To 互斥隐藏）** |
| From Address | textbox | `getByRole('textbox',{name:'From Address'})` | IP Range=Yes 时；必填(*)，IP 格式 |
| To Address | textbox | `getByRole('textbox',{name:'To Address'})` | IP Range=Yes 时；必填(*)，IP 格式 |
| IP Address | textbox | `getByRole('textbox',{name:'IP Address'})` | **IP Range=No 时**；必填(*)，IP 格式 |
| Description | textbox | `getByRole('textbox',{name:'Description'})` | 可选 |
| Confirm / Cancel | button | `getByRole('button',{name:'Confirm'})` / `{name:'Cancel'}` | 提交/取消 |
| Close | button | `getByRole('button',{name:'Close this dialog'})` | 关闭 |

## 5. 页面状态与分支

| 状态 | 触发 | 结果 |
|------|------|------|
| Disable（默认） | 进入页面 | 仅 Enable 单选 + Save |
| Enable | 选 Enable | 显示 Add Allow List 按钮 + 列表 |
| IP Range = Yes | 弹框默认 | From/To Address 均需填 |
| IP Range = No | 切换 | 仅单一 **IP Address\***（From/To 隐藏，互斥展示） |

## 6. 自动化测试要点

- 条件显隐（Enable 才有列表）。
- 新增白名单：IP 格式校验、From≤To 逻辑、Description 可空。
- ⚠️ 启用白名单可能锁定访问——自动化中需保证测试机 IP 在列表内。

## 7. 机器可解析摘要

```json
{
  "route": "/systemSettings/whitelist",
  "name": "whitelist",
  "title": "Access Control",
  "module": "System Settings",
  "fields": {"IP Allow List Enable": {"type":"radio","default":"Disable"}},
  "conditional": {"when":"Enable","shows":["Add Allow List","list table"]},
  "table_columns": ["No","Description","From IP","To IP","Action"],
  "dialog": {"title":"New Allow List","fields":["IP Range(Yes/No)","From Address(ip,when Yes)","To Address(ip,when Yes)","IP Address(ip,when No)","Description"],"buttons":["Confirm","Cancel"]},
  "buttons": ["Add Allow List","Save"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/systemsettings/`。

### 加载态 API
- `GET /api/settings/whitelistConfig`、`GET /api/whitelist/list`

### pytest 选择器与控件
- 页面/导航文案为 **Access Control**，开关文案为 **IP Allow List Enable***（`el-radio` Enable/Disable，设备常态 Disable）。
- Enable 后：`page.get_by_role("button", name="Add Allow List")` + `el-table`（列 No/Description/From IP/To IP/Action，空态 `No Data`）。
- 弹窗：`page.get_by_role("dialog", name="New Allow List")`；字段选择器须 scope 到 dialog。
  - IP Range*：`el-radio` Yes/No，默认 Yes；**Yes→From/To Address\*，No→单一 IP Address\*（互斥展示）**。

### 校验时机（实测）
- From/To/IP Address：**blur 即校验**（IP 格式）。

### Element-Plus 通用坑
- `el-radio` 点 label 兜底。
- ⚠️ 历史过时：本页**无 Port Range / Protocol 字段**（旧用例中含 Protocol/Port Range 步骤的均已过时）。

### 高危
- ⚠️ 启用白名单可能锁定访问——自动化须保证测试机 IP 在列表内。
