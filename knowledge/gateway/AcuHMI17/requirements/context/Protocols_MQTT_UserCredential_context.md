# Protocols / MQTT / User Credential — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/mqtt/credential` |
| 路由名 | `credential` |
| 面包屑 | AcuHMI-1-7 / Protocols / MQTT / User Credential |
| 顶级模块 | Protocols → MQTT |

## 2. 页面用途

配置 MQTT 连接 Broker 使用的用户名/密码认证凭据。

## 3. 交互元素清单 / 表单字段

| 字段/元素 | 类型 | 定位策略 1 | 定位策略 2 | 说明 |
|-----------|------|-----------|-----------|------|
| Username | textbox | `getByRole('textbox',{name:'Username'})` | placeholder "Enter Username" | 用户名（示例为空） |
| Password | textbox | `getByRole('textbox',{name:'Password'})` | placeholder "Enter Password" | 密码 |
| Save | button | `getByRole('button',{name:'Save'})` | 底部主按钮 | 保存 |

## 4. 校验规则

- 无 `*` 星号标注（可能为可选凭据；具体是否必填需结合 Broker 是否要求认证，运行时校验）。

## 5. 自动化测试要点

- 输入用户名/密码 → Save，断言保存成功。
- 密码字段是否掩码显示可作断言点。
- 与 General 页 Test MQTT 联动（凭据错误应导致连接失败）。

## 6. 机器可解析摘要

```json
{
  "route": "/protocols/mqtt/credential",
  "name": "credential",
  "title": "User Credential",
  "module": "Protocols/MQTT",
  "fields": {"Username":{"type":"text"},"Password":{"type":"text"}},
  "buttons": ["Save"]
}
```
