# System Settings / Email — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/systemSettings/email` |
| 路由名 | `email` |
| 面包屑 | AcuHMI-1-7 / System Settings / Email |
| 顶级模块 | System Settings |

## 2. 页面用途

配置发送邮件的 SMTP 服务器参数（用于报警通知等）。

## 3. 交互元素清单 / 表单字段

| 字段/元素 | 类型 | 定位策略 1 | 默认/示例 | 校验/说明 |
|-----------|------|-----------|-----------|-----------|
| Email Server | textbox | `getByRole('textbox',{name:'Email Server'})` | smtp.163.com | 必填(*)，合法 IP 或域名 |
| Email Port | textbox | `getByRole('textbox',{name:'Email Port'})` | 25 | 必填(*)，范围 **1–65535** |
| TLS/SSL | radiogroup | `getByRole('radiogroup',{name:'TLS/SSL'})` | Off | 必选(*) Auto/On/Off |
| Sender Name | textbox | `getByRole('textbox',{name:'Sender Name'})` | xiaoming | 可选，≤40 字符 |
| From Email Address | textbox | `getByRole('textbox',{name:'From Email Address'})` | 159xxxx4651@163.com | 必填(*)，≤100 字符 |
| Username | textbox | `getByRole('textbox',{name:'Username'})` | xiaoming123 | 必填(*)，≤40 字符 |
| Password | textbox(password+eye) | `getByRole('textbox',{name:'Password'})` | !@#AbC123 | ≤40 字符，含显隐图标 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 校验规则要点

- Email Server：合法 IP 或域名。
- Email Port：1–65535。
- From Email Address：≤100 字符（应为邮箱格式）。
- Sender Name / Username / Password：≤40 字符。
- TLS/SSL：三态 Auto/On/Off。

## 5. 自动化测试要点

- 端口范围、邮箱格式、字段长度校验。
- Password 显隐切换断言。
- 与 Alarm Notification 页联动（邮件通知依赖此配置）。

## 6. 机器可解析摘要

```json
{
  "route": "/systemSettings/email",
  "name": "email",
  "title": "Email",
  "module": "System Settings",
  "fields": {
    "Email Server": {"type":"text","required":true,"format":"ip_or_domain"},
    "Email Port": {"type":"text","required":true,"range":[1,65535]},
    "TLS/SSL": {"type":"radio","options":["Auto","On","Off"],"default":"Off"},
    "Sender Name": {"type":"text","maxlen":40},
    "From Email Address": {"type":"text","required":true,"maxlen":100},
    "Username": {"type":"text","required":true,"maxlen":40},
    "Password": {"type":"password","maxlen":40,"toggle":true}
  },
  "buttons": ["Save"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/systemsettings/`。

### 加载态 API
- `GET /api/settings/smtpConfig`

### pytest 选择器与控件
- 统一 `page.get_by_placeholder("Enter Email Server")` 等；TLS/SSL* 为 `el-radio` Auto/On/Off；Password 带显隐眼睛图标。
- ⚠️ **本页没有 Test Email 按钮**（Test Emails 在 Alarm Notification 页）。

### 校验时机（实测）
- **blur 即报错**：无（本页无纯格式 IP 类字段的 blur 报错记录）。
- **仅 Save 后异步报错**：长度上限类（Sender Name 40 / From Email Address 100 / Username 40 / Password 40）、数值范围类（Email Port 1–65535，实测填 99999 blur 无错）、邮箱正则类（From Email Address）。

### Element-Plus 通用坑
- `el-radio`（TLS/SSL）点 label 兜底。
