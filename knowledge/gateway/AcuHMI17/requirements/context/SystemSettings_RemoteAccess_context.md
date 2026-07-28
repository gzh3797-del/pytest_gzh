# System Settings / Remote Access — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/systemSettings/remoteAccess` |
| 路由名 | `remoteAccess` |
| 面包屑 | AcuHMI-1-7 / System Settings / Remote Access |
| 顶级模块 | System Settings（...More） |

## 2. 页面用途

启用/禁用设备远程访问功能。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 默认 | 说明 |
|------|------|-----------|------|------|
| Remote Access Enable | radiogroup | `getByRole('radiogroup',{name:'Remote Access Enable'})` | Disable | 必选(*)；Enable 可能显示额外连接参数（运行时确认） |
| Enable/Disable radio | radio(label) | `page.locator('label').filter({hasText:/^Enable$/})` | — | Element-Plus，点 label |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 自动化测试要点

- Enable/Disable 切换与保存；Enable 后是否出现附加字段的断言。

## 5. 机器可解析摘要

```json
{
  "route": "/systemSettings/remoteAccess",
  "name": "remoteAccess",
  "title": "Remote Access",
  "module": "System Settings",
  "fields": {"Remote Access Enable": {"type":"radio","default":"Disable"}},
  "buttons": ["Save"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/systemsettings/`。
> 入口：在 "...More" tooltip 弹出层内（先点 `menuitem "...More"` 再点子项，或直接 hash `#/systemSettings/remoteAccess`）。

### 加载态 API
- `GET /api/settings/deviceInfo`、`GET /api/settings/remoteDeviceAccess`

### pytest 选择器与控件
- Remote Access Enable*：`el-radio` Enable/Disable，设备常态 Disable，点 label 兜底。

### 未尽项（实测存疑，待补充）
- **未保存态切 Enable 不展开任何新字段/按钮**（仅本地状态，无 POST）；Manual Register / Refresh Status / Deregister 推测在保存并注册成功后才出现——**未探明**，需在允许保存的窗口期联机补充。
