# Protocols / Modbus / Parameters Mapping — 页面上下文

> 用途：AcuHMI-1-7 网关页面结构上下文，供 AI 将手工用例转换为自动化用例时按需加载。

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 (hash) | `/#/protocols/modbus/paramsMapping` |
| 路由名 | `paramsMapping` |
| 面包屑 | AcuHMI-1-7 / Protocols / Modbus / Parameters Mapping |
| 所属上下文 | 设置侧 (AcuHMI-1-7) |
| 顶级模块 | Protocols |
| 二级 | Modbus |
| 页面标题(H) | Parameters Mapping |

## 2. 进入前置条件

- 已登录（admin）。直接访问该 hash 若会话/协议上下文未就绪会被重定向到 `/#/systemSettings/dateTime`；登录后再次导航可正常进入。
- 页面数据依赖已在系统中添加的物理/虚拟设备（Physical Devices）。

## 3. 页面用途（业务说明）

将本网关已接入的下游设备（如 AcuRev/Acuvim 系列电表）的测量参数，映射为本网关对外提供的 **Modbus Slave** 寄存器地址空间。启用某设备后，为其选择起始地址，系统按参数顺序自动分配连续寄存器地址（Start/End Address, Dec.），上位机即可通过网关的 Slave ID 统一轮询这些设备参数。

## 4. 页面结构总览（区域自上而下）

1. **协议切换标签栏 (menubar)**：Modbus▼ | SNMP | BACnet/IP | MQTT▼ | AWS IoT | Azure IoT
   - `Modbus▼` 悬浮/点击展开二级：Parameters Mapping、Modbus Config、Device List、Device Mirror、Pass Through
   - `MQTT▼` 展开二级：General、User Credential、SSL/TLS、Last Will and Testament、Topic and Parameter Selection
2. **标题**：Parameters Mapping
3. **配置头**：`*Parameters Mapping Enable`（Enable/Disable 单选） + `Slave ID : 100`（只读展示）
4. **设备选择表（上表）**：列 = Enable(表头含全选复选框) / Device Name / Interface / Protocol / Model / Serial Number；每行一个 Enable 复选框
5. **映射控制区**：`Device Name :` 下拉（仅列出上表已勾选设备） + `*Start Address :` 文本框 + `Quick Sort` 按钮 + `Download List` / `Download All` 按钮
6. **参数映射表（下表）**：列 = Enable(表头复选框，未选设备时 disabled) / Device / Parameter / Type / Start Address(Dec.) / End Address(Dec.)；未选设备时显示 `No Data`
7. **Save** 按钮

## 5. 交互元素清单

| 元素 | 类型 | 定位策略 1 (推荐) | 定位策略 2 (备选) | 行为/说明 |
|------|------|-------------------|-------------------|-----------|
| Enable 单选组 | radiogroup | `getByRole('radiogroup', {name:'Parameters Mapping Enable'})` | 文本 "Parameters Mapping Enable" 邻接 radio | Enable(默认选中)/Disable；Disable 关闭整体映射 |
| Enable / Disable radio | radio | `getByRole('radio', {name:'Enable'})` / `name:'Disable'` | radiogroup 内第 1/2 项 | 互斥单选 |
| Slave ID | 只读文本 | `group('Slave ID :')` 内文本 | 文本 "Slave ID :" 后同级 | 展示网关 Slave ID（示例=100），不可编辑 |
| 设备表全选 | checkbox | 上表表头 columnheader "Enable" 内 checkbox | 表头首列 checkbox | 三态(checked=mixed 表示部分选中) |
| 设备行 Enable | checkbox | 行 `getByRole('row',{name:/AcuRev4100_392/}).getByRole('checkbox')` | 行首列 checkbox | 勾选后该设备进入 Device Name 下拉 |
| Device Name 下拉 | 自定义 combobox | `getByRole('combobox')`（Device Name 区域） | 文本 "---Select Device---" 容器 | 展开后 listbox；**仅列出上表已勾选设备** |
| Device Name 选项 | option | `getByRole('option',{name:'AcuRev4100_392'})` | listbox 内项 | 选中后下表加载该设备全部参数 |
| Start Address | textbox | `getByRole('textbox',{name:'Start Address'})` | placeholder "---Enter Start Address---" | 映射起始寄存器地址（Dec.），必填 (*) |
| Quick Sort | button | `getByRole('button',{name:'Quick Sort'})` | Start Address 输入框右侧按钮 | 依据起始地址快速重排/连续分配地址 |
| Download List | button | `getByRole('button',{name:'Download List'})` | 含 download 图标按钮 | 导出当前设备映射列表 |
| Download All | button | `getByRole('button',{name:'Download All'})` | 第 2 个 download 按钮 | 导出全部映射 |
| 参数行 Enable | checkbox | 下表行内 checkbox | 参数行首列 | 逐参数启用/停用映射；默认全选中 |
| 参数表全选 | checkbox | 下表表头 columnheader "Enable" 内 checkbox | 未选设备时 `[disabled]` | 全选/全不选参数 |
| Save | button | `getByRole('button',{name:'Save'})` | 页面底部主按钮 | 保存映射配置 |

## 6. 表单字段与校验规则

- `*Parameters Mapping Enable`：必选，默认 Enable。
- `*Start Address`：必填（星号），十进制寄存器起始地址；配合 Quick Sort 连续分配。
- 参数表地址（Start/End Address Dec.）：系统自动按参数顺序连续分配（示例：System Frequency 3000–3001，下一参数 3002–3003 …，FLOAT 类型占 2 个寄存器）。

## 7. 页面状态与分支

| 状态 | 触发 | 可观察结果 |
|------|------|-----------|
| 初始/未选设备 | 进入页面 | 下表 `No Data`，参数表头全选框 disabled |
| 已勾选设备 | 上表勾选某设备 | 该设备出现在 Device Name 下拉；示例仅 AcuRev4100_392 已勾选 |
| 已选设备+加载参数 | Device Name 选中设备 | 下表加载该设备**全部参数**（AcuRev4100_392 示例 = 1059 行），逐行 Type=FLOAT，地址连续 |
| Disable 整体 | Enable 组选 Disable | 关闭参数映射功能（下方配置不生效） |

## 8. 自动化测试要点

- **典型主流程**：进入页 → 上表勾选目标设备 Enable → Device Name 下拉选中该设备 → 填 Start Address → (可选 Quick Sort) → 参数表逐行/全选 Enable → Save。
- Device Name 下拉是**联动**元素：其选项集合完全由上表勾选状态决定；断言时需先确保设备已勾选。
- 参数表行数可能上千（示例 1059 行），遍历/断言全部行代价高——自动化中建议按"抽样首行/末行/特定参数名"定位（如按 Parameter 文本 "System Frequency" 定位行）。
- 地址自动分配规律：FLOAT 占 2 寄存器，End = Start+1，下一行 Start = 上一行 End+1。可用于地址连续性断言。
- Save 后应捕获成功提示/toast（本页保存反馈需运行时观察）。
- 协议标签栏 Modbus▼/MQTT▼ 为可展开二级菜单，是跨子页导航入口。

## 9. 关联子页（Modbus 二级）

Parameters Mapping、Modbus Config (`modbusConfig`)、Device List (`modbusDeviceList`)、Device Mirror (`logicalParameterMapping`)、Pass Through (`passThrough`)。

## 10. 机器可解析摘要

```json
{
  "route": "/protocols/modbus/paramsMapping",
  "name": "paramsMapping",
  "title": "Parameters Mapping",
  "module": "Protocols/Modbus",
  "context_side": "settings",
  "key_elements": {
    "radiogroup": ["Parameters Mapping Enable(Enable/Disable)"],
    "readonly": ["Slave ID"],
    "tables": ["device_select_table", "parameter_mapping_table"],
    "combobox": ["Device Name (依赖设备表勾选)"],
    "textbox": ["Start Address(*)"],
    "buttons": ["Quick Sort", "Download List", "Download All", "Save"]
  },
  "linkage": "Device Name 下拉选项 = 设备表已勾选行; 选中设备后参数表加载该设备全部参数(示例1059行)",
  "sub_pages": ["paramsMapping","modbusConfig","modbusDeviceList","logicalParameterMapping","passThrough"]
}
```
