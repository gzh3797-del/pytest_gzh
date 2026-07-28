# Protocols / MQTT / General — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/mqtt/general` |
| 路由名 | `general` |
| 面包屑 | AcuHMI-1-7 / Protocols / MQTT / General |
| 顶级模块 | Protocols → MQTT |

## 2. 页面用途

配置 MQTT 客户端连接 Broker 的基础参数（地址/端口/ClientID/保活/会话/超时），并可现场测试连接。MQTT 下有 5 个子页：General、User Credential、SSL/TLS、Last Will and Testament、Topic and Parameter Selection（经协议栏 `MQTT▼` 下拉进入）。

## 3. 交互元素清单 / 表单字段

| 字段/元素 | 类型 | 定位策略 1 | 定位策略 2 | 默认/示例 | 校验/说明 |
|-----------|------|-----------|-----------|-----------|-----------|
| MQTT Enable | radiogroup | `getByRole('radiogroup',{name:'MQTT Enable'})` | 文本 "MQTT Enable*" 邻接 | Enable | 必选 |
| Broker Address | textbox | `getByRole('textbox',{name:'Broker Address'})` | placeholder "Enter Broker Address" | www.accu.com | 必填(*) |
| Broker Port | textbox | `getByRole('textbox',{name:'Broker Port'})` | placeholder "Enter Broker Port" | 1883 | 必填(*)，端口号 |
| Client ID | textbox | `getByRole('textbox',{name:'Client ID'})` | placeholder "Enter Client ID" | pytest-mqtt-001 | 必填(*) |
| Generate Client ID | button | `getByRole('button',{name:'Generate Client ID'})` | Client ID 右侧按钮 | — | 自动生成 ClientID |
| Keep Alive | textbox | `getByRole('textbox',{name:'Keep Alive'})` | 后缀单位 "s" | 60 | 必填(*)，单位秒 |
| Clean Session | radiogroup | `getByRole('radiogroup',{name:'Clean Session'})` | 文本邻接 | Yes | 必选 Yes/No |
| Timeout | textbox | `getByRole('textbox',{name:'Timeout'})` | 后缀单位 "s" | 30 | 必填(*)，单位秒 |
| Save | button | `getByRole('button',{name:'Save'})` | 底部主按钮 | — | 保存 |
| Test MQTT | button | `getByRole('button',{name:'Test MQTT'})` | Save 右侧 | — | 测试连接（触发结果提示） |

## 4. 表单字段与校验规则

- 全部字段带 `*` 为必填。
- Broker Port：端口范围（通常 1–65535，1883 为 MQTT 默认）。
- Keep Alive / Timeout：正整数秒。
- Client ID：可手填或点 Generate 自动生成。

## 5. 页面状态与分支

| 状态 | 说明 |
|------|------|
| Enable | 显示全部连接字段 |
| Disable | 关闭 MQTT（字段可能禁用/隐藏，需运行时确认） |
| Test 成功/失败 | 点 Test MQTT 后弹出连接结果提示 |

## 6. 自动化测试要点

- 必填校验：清空各 `*` 字段后 Save 应报错。
- Generate Client ID 应生成非空且格式合法的 ID。
- Test MQTT 依赖真实 Broker，测试环境需 mock 或指向可达 broker；断言结果 toast。
- 与 SSL/TLS、User Credential 子页联动（启用 TLS 时端口通常 8883）。

## 7. 机器可解析摘要

```json
{
  "route": "/protocols/mqtt/general",
  "name": "general",
  "title": "General",
  "module": "Protocols/MQTT",
  "fields": {
    "MQTT Enable": {"type":"radio","default":"Enable"},
    "Broker Address": {"type":"text","required":true},
    "Broker Port": {"type":"text","required":true,"default":1883},
    "Client ID": {"type":"text","required":true,"has_generate_button":true},
    "Keep Alive": {"type":"text","required":true,"unit":"s","default":60},
    "Clean Session": {"type":"radio","options":["Yes","No"],"default":"Yes"},
    "Timeout": {"type":"text","required":true,"unit":"s","default":30}
  },
  "buttons": ["Generate Client ID","Save","Test MQTT"],
  "sibling_pages": ["general","credential","ssl","testament","deviceToPublish"]
}
```
