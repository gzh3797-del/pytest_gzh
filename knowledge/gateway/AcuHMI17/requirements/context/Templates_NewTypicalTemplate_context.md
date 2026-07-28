# Templates / New Typical Energy Meter Template — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/templates/newtypicalTemplateConfig` |
| 路由名 | `newtypicalTemplateConfig` |
| 面包屑 | AcuHMI-1-7 / Templates / New Typical Energy Meter Template |
| 顶级模块 | Templates |
| 相关动态页 | `typicalTemplateConfig/:type/:id`（编辑已有模板，结构相同，见文末） |

## 2. 页面用途

自定义典型电表模板的构建器：定义模板元信息、Modbus 数据块(Block)、参数(Parameter)的地址映射，最终创建模板。

## 3. 页面分区与元素

### 3.1 Device 区（可折叠）
| 字段 | 类型 | 定位 | 默认/示例 | 校验 |
|------|------|------|-----------|------|
| Template Name | textbox | `getByRole('textbox',{name:'Template Name'})` | 空 | 必填(*)，唯一 |
| Version | textbox | `getByRole('textbox',{name:'Version'})` | 空 | 必填(*)，同模板内唯一（如 v1.01） |
| Typical Model | combobox(disabled) | `getByRole('combobox',{name:'Typical Model'})` | Typical Energy Meter V2 | 只读 |
| Wiring Configuration | combobox | `getByRole('combobox',{name:'Wiring Configuration'})` | 3 Element 4 Wire Y | 必选(*) |

### 3.2 Block 区（可折叠）
| 字段 | 类型 | 定位 | 默认/示例 | 校验 |
|------|------|------|-----------|------|
| Function | combobox | `getByRole('combobox',{name:'Function'})` | ---Select Function--- | 必选(*)，Modbus 功能码 |
| Address Format | combobox | `getByRole('combobox',{name:'Address Format'})` | Hex | 必选(*) |
| Start | textbox(前缀0x) | `getByRole('textbox',{name:'Start'})` | 空 | 必填(*)，范围 **0x0–0xffff** |
| Count | textbox | `getByRole('textbox',{name:'Count'})` | 空 | 必填(*)，最小值 1 |
| Save Block | button | `getByRole('button',{name:'Save Block'})` | — | 保存该块到 Block Table |

### 3.3 Save 区
- 提示："Configuration completed? If you leave or refresh the page without saving it to database, all locally saved configurations will be discarded."
- `Create Template` button：`getByRole('button',{name:'Create Template'})` 最终创建模板。

### 3.4 Block Table
列：Index / Start(Hex) / Start(Dec) / Count / Function / Range / Action。初始 No Data。

### 3.5 Parameter Table
- `Display Tab(s)` combobox（Realtime 等）：切换参数分组显示。
- **Configured Parameter Table**：Parameter / Post Label / Range / Units / Address(Hex) / Address(Dec) / Multiplier / Action。初始 No Data。
- **Unconfigured Parameter Table**：同列；预置大量参数（Average Line Current I_avg_A A、Phase A Active Power P_a_kW kW 等，多页 6+ 页），每行 Action 按钮（配置该参数→移入 Configured）。

## 4. 页面状态与分支

| 状态 | 说明 |
|------|------|
| 未保存离开 | 本地配置将丢弃（提示文案） |
| Block 未保存 | Block Table 显示 No Data |
| 参数未配置 | 全在 Unconfigured 表；配置后移入 Configured 表 |

## 5. 自动化测试要点

- 主流程：填 Device 信息 → 定义 Block(Function/AddressFormat/Start/Count)→Save Block → 配置参数(Unconfigured→Configured)→ Create Template。
- 校验：模板名唯一、版本格式、Start 十六进制范围 0x0–0xffff、Count≥1。
- Display Tab(s) 切换参数集合；三表联动（Unconfigured→Configured）。
- 离开未保存提示（beforeunload/路由守卫）。

## 6. 机器可解析摘要

```json
{
  "route": "/templates/newtypicalTemplateConfig",
  "name": "newtypicalTemplateConfig",
  "title": "New Typical Energy Meter Template",
  "module": "Templates",
  "sections": {
    "Device": ["Template Name(req,unique)","Version(req,unique)","Typical Model(readonly)","Wiring Configuration(req)"],
    "Block": ["Function(req)","Address Format(req,Hex)","Start(req,0x0-0xffff)","Count(req,min1)","Save Block"],
    "Save": ["Create Template"]
  },
  "tables": ["Block Table","Configured Parameter Table","Unconfigured Parameter Table"],
  "param_columns": ["Parameter","Post Label","Range","Units","Address(Hex)","Address(Dec)","Multiplier","Action"],
  "edit_route": "/templates/typicalTemplateConfig/:type/:id"
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/templates/`。

### pytest 选择器与控件
- `.el-form-item` 按 `has_text` 过滤后取 `input` / `.el-select`：Template Name、Version、Typical Model、Wiring Configuration、Function、Address Format、Start、Count。
- Block 区：填 Function/Start/Count 后 `page.get_by_role("button", name="Save Block")`，再 `page.get_by_role("button", name="Create Template")` 提交。

### ★ 下拉全量枚举（实测真实文案）
- **Wiring Configuration（8 项，注意大小写、无复数 s）**：`3 Element 4 Wire Y`、`1 Element 2 Wire`、`2 Element 3 Wire 1 Phase`、`2 Element 3 Wire Network`、`2 Element 3 Wire Delta`、`3 Element 3 Wire Delta`、`3 Element 4 Wire Delta`、`2 1/2 Element 4 Wire Y`。
  - ⚠️ 手工步骤里的 “3 elements 4 Wire Y”（小写复数）**不存在**，正确为 `3 Element 4 Wire Y`。
- **Function（4 项）**：`READ_HOLDING_REGISTERS`、`READ_COILS`、`READ_DISCRETE_INPUTS`、`READ_INPUT_REGISTERS`。

### 结果反馈（实测）
- **创建成功 toast**：`.el-message--success`，文案 `Create Success.`（含句号，基础/派生同款）。断言 `expect(loc.first).to_be_visible()` + `inner_text()` 含 `Create Success`，避免旧 toast 残留假通过。

### 框架坑
- 同名/多组 `.el-select` 需父 `.el-form-item` scope；`.el-select__item` 选项用 `aria-controls`/可见性过滤。
