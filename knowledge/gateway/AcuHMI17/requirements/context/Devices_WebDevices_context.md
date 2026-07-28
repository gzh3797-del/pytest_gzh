# Devices / Web Devices — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/webDevices` |
| 路由名 | `webDevices` |
| 面包屑 | Devices / Web Devices |
| 上下文 | Devices 侧 |

## 2. 页面用途

管理"Web 设备"（通过 URL 链接的外部 Web 界面设备）。

## 3. 交互元素清单（列表）

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Add Device | button | `getByRole('button',{name:'Add Device'})` | 打开新增弹框 |
| 设备表 | table | 列: Device Name / Serial Number / Model / URL / Action | 初始 No Data |
| 分页 | button | `getByRole('button',{name:'Go to next page'})` | 单页 disabled |

## 4. 弹框：Add Device ★

| 字段 | 类型 | 定位策略 1 | 校验 |
|------|------|-----------|------|
| Device Name | textbox | `getByRole('textbox',{name:'Device Name'})` | 必填(*)，≤40 字符 |
| Serial Number | textbox | `getByRole('textbox',{name:'Serial Number'})` | 必填(*)，唯一，≤40 字符 |
| Model | textbox | `getByRole('textbox',{name:'Model'})` | 必填(*)，≤40 字符 |
| URL 协议 | combobox | group "URL" 内 combobox | https:// / http:// |
| URL 地址 | textbox | `getByRole('textbox',{name:'---Enter URL---'})` | 必填(*)，≤300 字符 |
| Confirm / Cancel | button | `getByRole('button',{name:'Confirm'})` / `{name:'Cancel'}` | 提交/取消 |
| Close | button | `getByRole('button',{name:'Close this dialog'})` | 关闭 |

## 5. 自动化测试要点

- 新增弹框各字段必填/长度/唯一校验；URL 协议下拉 + 地址；URL ≤300。
- 新增后出现在列表；行 Action 编辑/删除。

## 6. 机器可解析摘要

```json
{
  "route": "/webDevices",
  "name": "webDevices",
  "title": "Web Devices",
  "context_side": "devices",
  "table_columns": ["Device Name","Serial Number","Model","URL","Action"],
  "dialog": {"title":"Add Device","fields":{"Device Name":{"maxlen":40},"Serial Number":{"maxlen":40,"unique":true},"Model":{"maxlen":40},"URL":{"protocol_select":["https://","http://"],"maxlen":300}},"buttons":["Confirm","Cancel"]},
  "buttons": ["Add Device"]
}
```
