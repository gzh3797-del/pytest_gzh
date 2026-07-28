---
name: read-modbus-table
description: |
  解析 Modbus 地址表 Excel 文件，提取所有参数的地址、数据类型、功能码（FC）和缩放系数（scale）。

  当用户提到以下任意情况时调用本 skill：
  - 要求解析/分析/读取 Modbus 地址表、寄存器表、点表 Excel
  - 适配新设备前需要确认寄存器地址、数据类型、FC、scale
  - 任何 "/read-modbus-table"、"读一下这个地址表"、"这个表的 scale 是多少" 类请求
---

分析 Modbus 地址表 Excel 文件，提取所有参数的地址、数据类型、功能码和缩放系数。

## 用法

```
/read-modbus-table <Excel文件路径> [sheet名称（可选，不填则分析所有sheet）]
```

## 分析步骤

收到指令后，运行以下 Python 脚本读取 Excel，然后按规则推断每个参数的映射信息。

### 第一步：列出所有 Sheet

```python
import openpyxl
wb = openpyxl.load_workbook(r"<Excel路径>", data_only=True)
print(wb.sheetnames)
```

### 第二步：读取目标 Sheet

```python
ws = wb["<sheet名>"]
for row in ws.iter_rows(min_row=1, max_row=500, values_only=True):
    if any(v is not None for v in row):
        print(row)
```

### 第三步：逐行提取字段

标准列顺序（AcuRev 系列通用）：

| 列 | 字段 | 说明 |
|---|---|---|
| 0 | Block | 块名/功能码说明，如 `(0x03 Read)` `(01H Read, 05H Write)` |
| 1 | Start(Hex) | 起始地址（十六进制字符串） |
| 2 | End(Hex) | 结束地址 |
| 3 | Start(Dec) | 起始地址（十进制） |
| 4 | End(Dec) | 结束地址 |
| 5 | Description | 参数描述 |
| 6 | Data type | 数据类型 |
| 7 | RW | 读写属性 |
| 8 | Reg | 寄存器数量 |
| 9 | Range | 范围/精度（推断 scale 的关键列） |

### 第四步：推断功能码（FC）

从同一 Block 分组的首行 Block 列读取：

| Block 列内容 | FC |
|---|---|
| 含 `0x03` 或 `03H` | FC03（保持寄存器） |
| 含 `01H` 或 `0x01` | FC01（线圈，可写） |
| 含 `02H` 或 `0x02` | FC02（离散输入，只读） |
| 无标注 | 默认 FC03 |

### 第五步：推断数据类型和寄存器数

| Excel Data type | dtype | count | 说明 |
|---|---|---|---|
| Float / float | float32 | 2 | IEEE-754，无需缩放 |
| Double / float64 | float64 | 4 | IEEE-754，无需缩放 |
| Dword / DWord / uint32 | uint32 | 2 | 需查 Range 推断 scale |
| Word / word / uint16 | uint16 | 1 | 需查 Range 推断 scale |
| bit / Bit | bit | 1 | FC01/FC02，0 或 1 |

### 第六步：从 Range 列推断 scale（重点）

**规则 0 — 显式标注优先（最高优先级，命中则跳过规则 1-4）：**

表格任何列（Range / Default / 备注）出现显式缩放字样时**直接采用**，禁止再走推断：
```
"( scale:0.001)"              → scale = 0.001   （RPP 表风格，写在 Range 列）
"one/two/three decimal"       → 0.1/0.01/0.001  （AcuRev-1320 表风格，常写在 Default 列）
```
显式标注不是经验推断，**不进「待人工确认」清单**，但输出备注里注明标注出处列。
不同厂表标注列位置不固定——提取每行时把 Range 之后的所有列文本都扫一遍再下结论。

**规则 1 — 括号内有小数的 Dword/Word：**
```
"Max 999999999 (99999999.9)"  → 小数点后 1 位 → scale = 0.1
"Max 9999999 (9999.999)"      → 小数点后 3 位 → scale = 0.001
```
通用公式：`scale = 10 ** (-小数位数)`

**规则 2 — 范围上限暗示物理分辨率：**
```
"0~3600"  + 物理含义是角度(°) → 最大 360.0° → scale = 0.1
"0~1000"  + 物理含义是百分比 → 最大 100.0% → scale = 0.1
"0~10000" + 物理含义是百分比 → 最大 100.00% → scale = 0.01
```

**规则 3 — 格式字符串（Float 类）：**
```
"x.xx"  → 2 位小数，float32 本身已是工程值，scale = 1.0
"x.x"   → 1 位小数，同上
```

**规则 4 — Range 为空时，按参数类型和经验判断：**

| 参数类别 | 常见 scale | 备注 |
|---|---|---|
| Float32 实时量（V/A/W/Hz） | 1.0 | 已是工程单位 |
| Dword 能量（kWh/kvarh/kVAh） | 0.1 | 验证：BACnet 值通常比 Modbus 小 10 倍 |
| Word THD/OTHD/ETHD/THFF | 0.01 | 精度 0.01% |
| Word KF（K-Factor） | 0.1 | 精度 0.1 |
| Word CF（Crest Factor） | 0.001 | 精度 0.001 |
| Word 电压不平衡（UNBL_V） | 0.1 | 精度 0.1% |
| Word 电流不平衡（UNBL_I 系统） | 0.1 | 同上 |
| Word 电流不平衡（用户通道） | 0.01 | 精度 0.01% |
| Word 相角 | 0.1 | 0~3600 对应 0°~360° |

**规则 5 — 跨设备差异提示：**
- 🔴 **CHECKPOINT**：不同设备同类参数 scale 可能不同，每台设备都要单独确认；Range 为空、靠规则 4 经验值推断的 scale，须按下条比对验证并经人工确认后再用于下游
- 优先用 Range 列推断；Range 为空时用规则 4 作初始值，跑一次比对报告验证

### 第七步：输出格式

每个 Sheet 输出结构化表格：

```
=== Sheet: Active Energy ===
FC: 03  地址范围: 0x2500 ~ 0x252E

  地址(Hex)  地址(Dec)  描述                    dtype    FC  scale  备注
  0x2500     9472       Epin-A import           uint32   3   0.1    Max 999999999(99999999.9)
  0x2502     9474       Epin-B import           uint32   3   0.1
  ...

⚠️ 需人工确认 scale 的参数（Range 为空的 Word 类型）：
  0x325F  Unbalance I  → 建议 0.1，需比对验证
```

### 第八步：标记跳空地址

若相邻两行地址不连续（gap > 0），标注跳空：
```
  0x2046  EP_EXP    uint32  (↑ 跳过 0x2044~0x2045)
```
这些跳空寄存器在批量读取时不能合并到同一请求。

---

## 异常与格式漂移处理（if-then）

标准列顺序只是 AcuRev 系列的默认假设，遇到以下情况按表处理，禁止硬套列序或静默跳过：

| 触发条件 | 一线处理 | 仍失败兜底 |
|---|---|---|
| openpyxl 缺失 / Excel 打不开（加密、损坏、xls 旧格式） | `pip install openpyxl`；确认路径与扩展名 | 🛑 STOP 请用户另存为 .xlsx 或换副本，终止本次解析 |
| 表头与标准列顺序不符 | 读表头行按**列名**映射（Start/End/Data type/RW/Range/Default） | 列名也对不上 → 🛑 STOP 把前 5 行原样贴给用户确认列含义 |
| sheet 名含尾随空格 / 全角括号 / 笔误 | 用 `wb.sheetnames` 返回的原始字符串引用，禁止手敲 sheet 名 | — |
| Overview 总览与明细 sheet 矛盾（起始地址/长度不一致） | 以**明细 sheet 为准**，差异列入待人工确认清单 | — |
| 行有描述但地址列为空 | 不编造地址，单独列出报给用户 | — |
| 整个 sheet 为空 / 标 reserved | 记录"0 条参数"并注明，不视为解析失败 | — |

## 注意事项

1. **Block 列合并单元格**：Excel 中 Block 列常跨多行合并，读到的非首行为 None，需向上继承当前 Block 值
2. **功能码在 Block 列**：如 `(01H Read, 05H | 0FH Write)` 表示 FC01 读、FC05/FC15 写
3. **地址步长验证**：Dword 步长应为 2，Word 步长为 1，Float 步长为 2，Double 步长为 4
4. **scale 最终以实测数据为准**：Range 列推断是初始值，第一次跑完比对报告后根据 BACnet vs Modbus 实际差值验证

---

## ❌ 反例黑名单（不要做）

以下每条都是真实踩过的坑，任一命中即返工：

| # | 反例 | 后果 / 纠正 |
|---|------|------------|
| 1 | 只读部分 sheet 就推断全量参数 | 漏整块寄存器（团队约定第 7 条）→ 第一步必须列出**全部** sheet 名逐一确认 |
| 2 | Range 为空、靠规则 4 经验值推出的 scale 直接交给下游用 | 比对报告整列偏 10/100 倍 → 必须跑比对验证 + 人工确认后再用 |
| 3 | Block 列读到 None 当作"无功能码" | FC 判错、批量读失败 → 向上继承当前 Block 值 |
| 4 | 跳空地址合并进同一批量读请求 | 读取报 Illegal Data Address → 按第八步标记跳空、拆开请求 |
| 5 | 不做地址步长验证就输出 dtype | Dword 标成 Word、地址错位 → 按注意事项 3 逐块校验步长 |
| 6 | 跨设备沿用同名参数的 scale | 不同设备同类参数 scale 可能不同 → 每台设备单独确认（规则 5） |
| 7 | 表里明写显式 scale 标注仍按经验规则推断或列入待确认 | RPP 表 0x17C7 Range 列明写 `( scale:0.001)` 却被归为"待确认建议 1.0"（2026-07-09 实测踩坑）→ 规则 0 显式标注最高优先级 |
