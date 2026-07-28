# Protocols / MQTT / Last Will and Testament — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/protocols/mqtt/testament` |
| 路由名 | `testament` |
| 面包屑 | AcuHMI-1-7 / Protocols / MQTT / Last Will and Testament |
| 顶级模块 | Protocols → MQTT |

## 2. 页面用途

配置 MQTT 遗嘱消息（LWT）：客户端异常断开时 Broker 代为发布的 Topic/QoS。**条件分支页**（Disable→仅开关；Enable→显示 Topic + QoS）。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 定位策略 2 | 说明 |
|------|------|-----------|-----------|------|
| Last Will Enable | radiogroup | `getByRole('radiogroup',{name:'Last Will Enable'})` | 文本邻接 | 默认 **Disable** |
| Enable/Disable radio | radio(label) | `page.locator('label').filter({hasText:/^Enable$/})` | — | Element-Plus，点 label |
| Topic | textbox | `getByRole('textbox',{name:'Topic'})` | placeholder "Enter Topic" | 必填(*)，示例 acurev4100/lwt |
| Qos | 自定义 combobox | `getByRole('combobox',{name:'Qos'})` | 文本 "Qos 2" 容器 | 必选(*)，Qos 0/1/2（默认 Qos 2） |
| Save | button | `getByRole('button',{name:'Save'})` | 底部主按钮 | 保存 |

## 4. 页面状态与分支 ★

| 状态 | 触发 | 可观察结果 |
|------|------|-----------|
| Disable（默认） | 进入页面 | 仅 Last Will Enable 单选 + Save |
| Enable | 选 Enable | 显示 Topic* 输入 + Qos* 下拉 |

## 5. 表单字段与校验规则

- `Last Will Enable*`：必选，默认 Disable。
- `Topic*`：启用后必填。
- `Qos*`：启用后必选，取值 Qos 0 / Qos 1 / Qos 2。

## 6. 自动化测试要点

- 条件显隐分支断言（Disable 无 Topic/QoS，Enable 出现）。
- Topic 必填校验；QoS 下拉选项覆盖。
- Element-Plus radio 点 label。

## 7. 机器可解析摘要

```json
{
  "route": "/protocols/mqtt/testament",
  "name": "testament",
  "title": "Last Will and Testament",
  "module": "Protocols/MQTT",
  "fields": {"Last Will Enable":{"type":"radio","default":"Disable"}},
  "conditional": {"when":"Last Will Enable=Enable","shows":["Topic(text,required)","Qos(select: 0/1/2,default 2)"]},
  "buttons": ["Save"]
}
```
