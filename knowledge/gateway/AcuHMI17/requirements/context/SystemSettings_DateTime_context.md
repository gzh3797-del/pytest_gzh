# System Settings / Date & Time — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/systemSettings/dateTime` |
| 路由名 | `dateTime` |
| 面包屑 | AcuHMI-1-7 / System Settings / Date & Time |
| 顶级模块 | System Settings |

## 2. 页面用途

配置设备时间：NTP 自动同步开关、NTP 服务器、时区，或手动设置设备时钟。System Settings 二级 tab 栏含：Date & Time / Network / Access Control / Email / Alarm Notification / Certificate Management / (...More→) Configuration Management / Remote Access。

## 3. 交互元素清单 / 表单字段

| 字段/元素 | 类型 | 定位策略 1 | 默认/示例 | 说明 |
|-----------|------|-----------|-----------|------|
| NTP Enable | radiogroup | `getByRole('radiogroup',{name:'NTP Enable'})` | Enable | 必选(*)；Enable→NTP 服务器可用，Disable→手动设时钟 |
| Device Clock | date/time combobox | `getByRole('combobox',{name:'Device Clock'})` | 2026/07/06 08:49 PM | 时间选择器 |
| Sync | button | `getByRole('button',{name:'Sync'})` | — | 同步（显示 "Last updated at ..."） |
| NTP Server 1 | textbox | `getByRole('textbox',{name:'NTP Server 1'})` | time.apple.co | 必填(*)，最多 40 字符 |
| NTP Server 2 | textbox | `getByRole('textbox',{name:'NTP Server 2'})` | 空 | 可选，最多 40 字符 |
| NTP Server 3 | textbox | `getByRole('textbox',{name:'NTP Server 3'})` | 空 | 可选，最多 40 字符 |
| Time Zone | combobox | `getByRole('combobox',{name:'Time Zone'})` | Asia/Shanghai(CST) | 必选(*) |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 页面状态与分支

| 状态 | 触发 | 结果 |
|------|------|------|
| NTP Enable | 默认 | NTP Server 1–3 / Time Zone 可用；Device Clock 自动 |
| NTP Disable | 切换 | 需手动设置 Device Clock（NTP 服务器可能禁用，运行时确认） |

## 5. 校验规则要点

- NTP Server 1 必填，各 Server ≤40 字符。
- Time Zone 必选。

## 6. 自动化测试要点

- NTP Enable/Disable 切换的字段可用性联动。
- NTP Server 长度校验；Sync 后 "Last updated at" 时间戳更新。
- 时区下拉选项覆盖。

## 7. 机器可解析摘要

```json
{
  "route": "/systemSettings/dateTime",
  "name": "dateTime",
  "title": "Date & Time",
  "module": "System Settings",
  "fields": {
    "NTP Enable": {"type":"radio","default":"Enable"},
    "Device Clock": {"type":"datetime"},
    "NTP Server 1": {"type":"text","required":true,"maxlen":40},
    "NTP Server 2": {"type":"text","maxlen":40},
    "NTP Server 3": {"type":"text","maxlen":40},
    "Time Zone": {"type":"select","default":"Asia/Shanghai(CST)"}
  },
  "buttons": ["Sync","Save"],
  "sibling_tabs": ["Date & Time","Network","Access Control","Email","Alarm Notification","Certificate Management","Configuration Management","Remote Access"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/systemsettings/`。以下为真机实测事实。

### 进入路径
- 子页菜单：`page.get_by_role("menuitem", name="Date & Time").click()`；用例内可直接 hash 跳转 `page.goto(base + "#/systemSettings/dateTime")` 绕过。
- **"...More" 是 tooltip 弹出层**（非常驻 DOM）：Configuration Management / Remote Access 需先 `page.get_by_role("menuitem", name="...More").click()` 再点其内项（本页不涉及）。

### 加载态 API
- `GET /api/settings/ntpConfig`、`GET /api/command/getTime`

### pytest 选择器与控件
- NTP Enable*：`el-radio`，**点 label**（`page.get_by_text("Enable", exact=True)` scope 到该 `.el-form-item`），选中态判 class 含 `is-checked`。Disable 时 NTP Server 1/2/3 禁用。
- Device Clock*：`el-date-picker`（datetime），`page.get_by_role("combobox", name="Device Clock*")`。
- Sync：`page.get_by_role("button", name="Sync")`（点击即触发同步）。
- NTP Server 1/2/3：`page.get_by_placeholder("NTP Server 1")` 等。
- Time Zone*：**EP v2 select（`.el-select__input`），下拉 405 个 IANA 时区项**——用 `input.getAttribute('aria-controls')` 取专属 listId 后在该 popper 下查 `.el-select-dropdown__item`，**禁止全局 querySelectorAll**。默认 `Asia/Shanghai(CST)`，支持搜索。
- Save：`page.get_by_role("button", name="Save")`。

### 校验时机（实测）
- NTP Server 长度上限（40）：**无原生 maxlength，仅点 Save 后异步渲染错误**（blur 不报错）。

### Element-Plus 通用坑（本模块共用）
- `el-radio` 原生 click 被 `.el-radio__inner` 拦截 timeout → 优先点文案 label。
- 多组同名元素需父容器 scope。
