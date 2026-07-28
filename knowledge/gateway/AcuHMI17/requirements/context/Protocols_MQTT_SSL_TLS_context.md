# Protocols / MQTT / SSL/TLS — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/mqtt/ssl` |
| 路由名 | `ssl` |
| 面包屑 | AcuHMI-1-7 / Protocols / MQTT / SSL/TLS |
| 顶级模块 | Protocols → MQTT |

## 2. 页面用途

配置 MQTT 连接的 TLS 加密。启用后需上传 CA / 客户端证书 / 私钥三类文件。**典型条件分支页**（Disable→仅显示开关；Enable→显示 3 个上传字段）。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 定位策略 2 | 说明 |
|------|------|-----------|-----------|------|
| Enable SSL | radiogroup | `getByRole('radiogroup',{name:'Enable SSL'})` | 文本 "Enable SSL*" 邻接 | 默认 **Disable** |
| Enable/Disable radio label | radio(label) | `page.locator('label').filter({hasText:/^Enable$/})` | radiogroup 内项 | ⚠️ Element-Plus radio 需点 label，不能直接点 input |
| CA File Browse | button+file | `group('CA File').getByRole('button',{name:'Browse'})` | "Choose file" 邻接 | 上传 CA（示例已传 ca.crt） |
| Cert File Browse | button+file | `group('Cert File').getByRole('button',{name:'Browse'})` | — | 上传客户端证书（client.crt） |
| Key File Browse | button+file | `group('Key File').getByRole('button',{name:'Browse'})` | — | 上传私钥（client.key） |
| Save | button | `getByRole('button',{name:'Save'})` | 底部主按钮 | 保存 |

## 4. 页面状态与分支 ★

| 状态 | 触发 | 可观察结果 |
|------|------|-----------|
| SSL Disable（默认） | 进入页面 | 仅显示 Enable SSL 单选 + Save |
| SSL Enable | 选 Enable | **显示 3 个上传组**：CA File / Cert File / Key File，各含 "Choose file" + Browse，且显示 "Uploaded File: xxx" |

## 5. 表单字段与校验规则

- `Enable SSL*`：必选，默认 Disable。
- 启用后 CA/Cert/Key 三文件通常必传（TLS 双向认证）。
- 已上传文件以 "Uploaded File: 文件名" 形式回显。

## 6. 自动化测试要点（重点）

- **条件显隐**：断言 Disable 时无上传字段、Enable 时出现 3 个上传组——这是核心分支用例。
- 文件上传通过 `browser_file_upload` / `<input type=file>` 完成；断言上传后回显文件名。
- Element-Plus 单选控件必须点击 `label`（radio input 被 `.el-radio__inner` 遮挡，直接点 input 会超时）。
- 与 General 页联动：启用 TLS 时 Broker Port 通常改 8883。

## 7. 机器可解析摘要

```json
{
  "route": "/protocols/mqtt/ssl",
  "name": "ssl",
  "title": "SSL/TLS",
  "module": "Protocols/MQTT",
  "fields": {"Enable SSL": {"type":"radio","default":"Disable"}},
  "conditional": {"when":"Enable SSL=Enable","shows":["CA File(upload)","Cert File(upload)","Key File(upload)"]},
  "buttons": ["Browse x3","Save"],
  "ui_note": "Element-Plus radio: click label not input"
}
```
