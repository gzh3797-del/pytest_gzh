# System Settings / Alarm Notification — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/systemSettings/alarm` |
| 路由名 | `alarm` |
| 面包屑 | AcuHMI-1-7 / System Settings / Alarm Notification |
| 顶级模块 | System Settings |

## 2. 页面用途

配置报警通知：蜂鸣、报警确认、邮件通知收件人与间隔，并可测试发信。

## 3. 交互元素清单 / 表单字段

| 字段/元素 | 类型 | 定位策略 1 | 默认/示例 | 校验/说明 |
|-----------|------|-----------|-----------|-----------|
| Alarm Beep | switch | `getByRole('switch',{name:'Alarm Beep'})` | ON | ON/OFF 开关 |
| Alarm Acknowledgement Enable | radiogroup | `getByRole('radiogroup',{name:'Alarm Acknowledgement Enable'})` | Enable | 必选(*) |
| Alarm Email Enable | radiogroup | `getByRole('radiogroup',{name:'Alarm Email Enable'})` | Enable | 必选(*)，控制收件人/间隔显隐 |
| Recipient 1 | textbox | `getByRole('textbox',{name:'Recipient 1'})` | ...@qq.com | 必填(*)，≤100 字符 |
| Recipient 2 | textbox | `getByRole('textbox',{name:'Recipient 2'})` | — | 可选，≤100 字符 |
| Recipient 3 | textbox | `getByRole('textbox',{name:'Recipient 3'})` | — | 可选，≤100 字符 |
| Email Interval | textbox | `getByRole('textbox',{name:'Email Interval'})` | 5 | 必填(*)，范围 **1–10** mins |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |
| Test Emails | button | `getByRole('button',{name:'Test Emails'})` | — | 测试发信 |

## 4. 页面状态与分支

| 状态 | 触发 | 结果 |
|------|------|------|
| Alarm Email Enable | 默认 | 显示 Recipient 1–3 + Email Interval |
| Alarm Email Disable | 切换 | 隐藏收件人/间隔（运行时确认） |

## 5. 校验规则要点

- Recipient 1 必填；各 Recipient ≤100 字符、邮箱格式。
- Email Interval：1–10 分钟。

## 6. 自动化测试要点

- Email Enable 显隐收件人/间隔分支。
- 收件人邮箱格式与长度校验、Interval 范围。
- Test Emails 依赖 Email 页 SMTP 配置；断言发信结果 toast。

## 7. 机器可解析摘要

```json
{
  "route": "/systemSettings/alarm",
  "name": "alarm",
  "title": "Alarm Notification",
  "module": "System Settings",
  "fields": {
    "Alarm Beep": {"type":"switch","default":"ON"},
    "Alarm Acknowledgement Enable": {"type":"radio","default":"Enable"},
    "Alarm Email Enable": {"type":"radio","default":"Enable","controls":["Recipient 1-3","Email Interval"]},
    "Recipient 1": {"type":"text","required":true,"maxlen":100},
    "Recipient 2": {"type":"text","maxlen":100},
    "Recipient 3": {"type":"text","maxlen":100},
    "Email Interval": {"type":"text","range":[1,10],"unit":"mins"}
  },
  "buttons": ["Save","Test Emails"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/systemsettings/`。

### 加载态 API
- `GET /api/settings/alarmEmailConfig`

### pytest 选择器与控件
- Alarm Beep*：`el-switch`，`page.get_by_role("switch", name="Alarm Beep*")`，默认 ON。
- Alarm Acknowledgement Enable* / Alarm Email Enable*：`el-radio`，默认 Enable。
- Recipient 1*/2/3：`page.get_by_placeholder("Enter Recipient 1")` 等（≤100 字符）。
- Email Interval*：`page.get_by_placeholder("Enter Email Interval")`，Range 1–10 mins。
- Save；**Test Emails（会真实发邮件，慎点）**。

### 校验时机（实测）
- **仅 Save 后异步报错**：Recipient 长度(100)/邮箱正则、Email Interval 范围(1–10)。

### Element-Plus 通用坑
- `el-radio` 点 label 兜底。

### 高危
- ⚠️ **Test Emails 会真实发信**——自动化中慎点/需 mock SMTP。依赖 Email 页 SMTP 配置。

### 补充（2026-07-17 联机实测，对应测试目录 `projects/RPP/tests/Alarm/`）
- 直连 `/#/systemSettings/alarm` 存在渲染竞态（可能渲染成 Date&Time 页内容）；稳定做法：先 `goto` `/#/systemSettings/dateTime`，再点顶部 menuitem 'Alarm Notification' 切入本页。
- Alarm Email Enable=Enable 且 Recipient 1 为空时，点 Save 会被必填校验拦下；自动化如需改 Alarm Acknowledgement Enable，须先把 Email 切到 Disable 再 Save，避免被邮箱必填卡住。
- Alarm Acknowledgement Enable 切 Disable 的联动：全局 Unacknowledged Alarms tab 隐藏、全局与设备详情页 Alarm Logs 的 Ack Status 列隐藏；切回 Enable 后历史 Acknowledged 记录不丢失。
