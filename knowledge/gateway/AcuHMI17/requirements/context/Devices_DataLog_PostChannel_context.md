# Devices / Data Log / Post Channels (1/2/3) — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/dataLog/postChannels/postChannel1`（`postChannel1`；同构 `postChannel2`/`postChannel3`） |
| 父路由 | `/#/dataLog/postChannels`（`postChannels`，分组入口） |
| 面包屑 | Devices / Data Log / Post Channels / Post Channel 1 |
| 上下文 | Devices 侧 |

> **本文档覆盖 Post Channel 1/2/3**（结构一致）。

## 2. 页面用途

配置数据上报通道（FTP/HTTP 等）。若该通道被某 Data Logger 引用，则 Enable 被锁定（不可禁用）。

## 3. 交互元素清单

| 字段 | 类型 | 定位策略 1 | 默认/示例 | 校验/说明 |
|------|------|-----------|-----------|-----------|
| Post Channel N Enable | radiogroup | `getByRole('radiogroup',{name:'Post Channel 1 Enable'})` | Enable | 被 Data Logger 引用时 `disabled`（提示 "Cannot be disabled. Used by Data Logger 1."） |
| Post Method | combobox | `getByRole('combobox',{name:'Post Method'})` | FTP | 上报方式（FTP/HTTP/HTTPS/… 切换字段） |
| FTP URL | textbox | `getByRole('textbox',{name:'FTP URL'})` | 192.168.2.45（FTP:// 前缀） | 必填(*)，合法 IP 或域名 |
| FTP Port | textbox | `getByRole('textbox',{name:'FTP Port'})` | 2121 | 必填(*)，范围 **1–65535** |
| Enable anonymous mode | checkbox | `getByRole('checkbox',{name:'Enable anonymous mode'})` | 未勾 | 勾选后隐藏用户名/密码 |
| FTP User Name | textbox | `getByRole('textbox',{name:'FTP User Name'})` | datalog | 必填(*)，≤40 字符 |
| FTP password | textbox(password+eye) | `getByRole('textbox',{name:'FTP password'})` | datalog123 | 必填(*)，≤40，含显隐图标 |
| Test Post Channel | button | `getByRole('button',{name:'Test Post Channel'})` | — | 测试上报 |
| Clear Post Channel Logs | button | `getByRole('button',{name:'Clear Post Channel Logs'})` | — | 清空日志 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 页面状态与分支 ★

| 状态 | 说明 |
|------|------|
| 被 Logger 引用 | Enable radio disabled，不可禁用 |
| Post Method 切换 | FTP→FTP 字段；HTTP/HTTPS→URL/认证等（运行时确认） |
| Enable anonymous mode | 勾选后无需用户名/密码 |

## 5. 自动化测试要点

- Post Method 切换字段显隐（核心分支）；FTP Port 范围、URL 格式、账号≤40 校验。
- 匿名模式勾选后用户名/密码显隐；密码显隐图标。
- Test Post Channel 结果提示（依赖真实 FTP，需 mock）；被引用时禁用断言。
- 三个 Post Channel（1/2/3）行为一致，可参数化。

## 6. 机器可解析摘要

```json
{
  "route": "/dataLog/postChannels/postChannel1",
  "name": "postChannel1",
  "title": "Post Channel 1",
  "context_side": "devices",
  "covers": ["postChannel1","postChannel2","postChannel3"],
  "fields": {
    "Post Channel Enable": {"type":"radio","default":"Enable","disabled_when":"used by a Data Logger"},
    "Post Method": {"type":"select","default":"FTP"},
    "FTP URL": {"type":"text","required":true,"format":"ip_or_domain"},
    "FTP Port": {"type":"text","required":true,"range":[1,65535]},
    "Enable anonymous mode": {"type":"checkbox"},
    "FTP User Name": {"type":"text","required":true,"maxlen":40},
    "FTP password": {"type":"password","required":true,"maxlen":40,"toggle":true}
  },
  "conditional": {"Post Method":"switches transport fields","anonymous":"hides credentials"},
  "buttons": ["Test Post Channel","Clear Post Channel Logs","Save"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/datalog/`（postchannel 子目录）。

### 进入路径
- 父 `div.el-sub-menu__title`（`Post Channels`）→ 子 `.el-menu-item`（`Post Channel 1/2/3`）。

### pytest 选择器与控件
- Enable：`el-radio`，`.el-radio` filter Enable/Disable，选中态 class 含 `is-checked`。**被某 Data Logger 引用时 Enable 锁定不可禁用**（提示 "Cannot be disabled. This Post Channel is used by Data Logger 1."）。
- **本 Channel 在 Data Logger 配置页下拉中的表现**：选项文案 `Post Channel N`（**带空格**）；当本 Channel Disable 时，选项仍在下拉但带 `is-disabled`（不移除），点击无效。断言时先 presence 再判 disabled。

### 框架坑与兜底
- `el-menu` popper 遮挡 → 坐标点击兜底；`el-radio` 点 label 兜底。

### 保存与成功判定
```python
page.get_by_role("button", name="Save").click(); page.wait_for_timeout(1500)
assert page.locator(".el-message--error").count() == 0
```

### 高危
- ⚠️ **保存 Post Channel 可能触发服务重启**，读取前留足等待。
- Test Post Channel 依赖真实 FTP/HTTP 目标（需 mock）；Clear Post Channel Logs 为清空类操作。
