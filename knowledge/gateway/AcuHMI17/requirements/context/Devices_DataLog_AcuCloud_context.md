# Devices / Data Log / AcuCloud — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/dataLog/acucloud` |
| 路由名 | `acucloud` |
| 面包屑 | Devices / Data Log / AcuCloud |
| 上下文 | Devices 侧 |

> Data Log 二级 tab：Data Loggers▼(dataLogger1/2/3, Data Log Parameter Config, Rapid Logger) / Post Channels▼(postChannel1/2/3) / Data Log Management / Post Historical Data / AcuCloud。

## 2. 页面用途

配置将设备数据上报至 AcuCloud 云平台。**条件分支页**（Disable→仅开关）。

## 3. 交互元素清单（Enable 后）

| 元素 | 类型 | 定位策略 1 | 默认/示例 | 说明 |
|------|------|-----------|-----------|------|
| AcuCloud Enable | radiogroup | `getByRole('radiogroup',{name:'AcuCloud Enable'})` | Disable | 必选(*)，控制显隐 |
| Module Serial Number | textbox(disabled) | `getByRole('textbox',{name:'Module Serial Number'})` | AHI2606080066 | 只读 + Copy 按钮 |
| Copy | button | `getByRole('button',{name:'Copy'})` | — | 复制序列号 |
| AcuCloud Token | textbox | `getByRole('textbox',{name:'AcuCloud Token'})` | — | 必填(*)，≤40 字符 |
| Link to AcuCloud | link | `getByRole('link',{name:'Link to AcuCloud'})` | https://acucloud.accuenergy.com/ | 外链 |
| Devices Selection 表 | table | group "Devices Selection" | — | checkbox/Device Name/Serial Number/Protocol/Online |
| Test AcuCloud | button | `getByRole('button',{name:'Test AcuCloud'})` | — | 测试连接 |
| Clear AcuCloud Post Logs | button | `getByRole('button',{name:'Clear AcuCloud Post Logs'})` | — | 清空上报日志 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |

## 4. 页面状态与分支

| 状态 | 触发 | 结果 |
|------|------|------|
| Disable（默认） | 进入页面 | 仅 AcuCloud Enable 单选 + Save |
| Enable | 选 Enable | 显示序列号/Token/设备选择/测试等 |

## 5. 自动化测试要点

- 条件显隐；Token 必填≤40；设备勾选。
- Copy 复制序列号；Test AcuCloud 结果提示（依赖真实云，需 mock）。
- Clear Post Logs 清空。

## 6. 机器可解析摘要

```json
{
  "route": "/dataLog/acucloud",
  "name": "acucloud",
  "title": "AcuCloud",
  "context_side": "devices",
  "fields": {
    "AcuCloud Enable": {"type":"radio","default":"Disable"},
    "Module Serial Number": {"type":"text","readonly":true},
    "AcuCloud Token": {"type":"text","required":true,"maxlen":40}
  },
  "device_table": ["checkbox","Device Name","Serial Number","Protocol","Online"],
  "buttons": ["Copy","Test AcuCloud","Clear AcuCloud Post Logs","Save"],
  "sub_tabs": ["Data Loggers","Post Channels","Data Log Management","Post Historical Data","AcuCloud"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/datalog/`。

### 进入路径
- 先确保在 Devices 视图（`page.locator("header span").filter(has_text="Devices").first.click()`），再点左侧 `.left-nav-item` `Data Log`；AcuCloud 为顶层子项 `.el-menu-item`（无父展开），或直达 `#/dataLog/acucloud`。

### pytest 选择器与控件
- AcuCloud Enable*：`el-radio` Enable/Disable（默认 Disable），点 label 兜底。
- Enable 后：Module Serial Number（只读）+ Copy、AcuCloud Token*（≤40）、设备选择表、Test AcuCloud、Clear AcuCloud Post Logs。

### 保存与成功判定
```python
page.get_by_role("button", name="Save").click(); page.wait_for_timeout(1500)
assert page.locator(".el-message--error").count() == 0
```

### 高危
- Test AcuCloud 依赖真实云（需 mock）；Clear AcuCloud Post Logs 为清空类操作。
