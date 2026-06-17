# 接线检查算法规格摘要 v1.05

原件：`requirements/raw/接线检测总表_ver1.05.xlsx`（共 10 个 Sheet）
引擎：`test_case/ACM_41_WEB2/wiring_check/core/expected_engine.py`

---

## 适用产品范围

| 产品 | 适用接线方式 |
|------|------------|
| ACM-41-WEB2（4100 多回路） | 3E4WY、2E3W Delta、2E3W Network、2E3W 1Phase、1E2W（共 5 种） |
| AcuHMI-1-7 | 以上 5 种 + 2.5E4WY、3E3W Delta、3E4W Delta（HighLeg）、2E2W Delta（共 9 种） |

> WEB2 Sprint 2 测试脚本仅覆盖 5 种；HMI1-7 适配时需补充剩余 4 种。

---

## 一、3E4WY（3 Element 4 Wire Y）

接入：Ua、Ub、Uc、Ia、Ib、Ic；额定电压用**相电压**

### 电压侧（优先级：缺失 > 反接 > 相位错误；条件 1-10 串行，11-13 独立）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 1 | Van < 0.1·VRATE **且** Vbn < 0.1·VRATE **且** Vcn < 0.1·VRATE | Va & Vb & Vc Wiring Missing |
| 2 | Vab < 0.1·VRATE **且** Vcn ∈ [0.8, 1.2]·VRATE | Va & Vb Wiring Missing |
| 3 | Vbc < 0.1·VRATE **且** Van ∈ [0.8, 1.2]·VRATE | Vb & Vc Wiring Missing |
| 4 | Vca < 0.1·VRATE **且** Vbn ∈ [0.8, 1.2]·VRATE | Va & Vc Wiring Missing |
| 5 | Van ≤ 0.8·VRATE | Va Wiring Missing |
| 6 | Vbn ≤ 0.8·VRATE | Vb Wiring Missing |
| 7 | Vcn ≤ 0.8·VRATE | Vc Wiring Missing |
| 8 | Van ∈ (0.8, 1.2)·VRATE **且** Vbn > 1.3·VRATE **且** Vcn > 1.3·VRATE | Va-Vn Reversed |
| 9 | Vbn ∈ (0.8, 1.2)·VRATE **且** Van > 1.3·VRATE **且** Vcn > 1.3·VRATE | Vb-Vn Reversed |
| 10 | Vcn ∈ (0.8, 1.2)·VRATE **且** Van > 1.3·VRATE **且** Vbn > 1.3·VRATE | Vc-Vn Reversed |
| 11（独立） | ABC: ∠Vb ∉ [240±20°]；ACB: ∠Vb ∉ [120±20°] | Vb Phase Shift |
| 12（独立） | ABC: ∠Vc ∉ [120±20°]；ACB: ∠Vc ∉ [240±20°] | Vc Phase Shift |
| 13（独立） | V_unbalance ≥ 145% | Phase Order Error |

> 1300/2100 不支持相序配置，条件 11-13 不执行。

### 电流侧（多回路 4100；条件 14-16 缺失各自独立，17-22 独立；17/20 互斥，18/21 互斥，19/22 互斥）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 14 | Ian < 0.1A | Ia Wiring Missing（跳过 17/20） |
| 15 | Ibn < 0.1A | Ib Wiring Missing（跳过 18/21） |
| 16 | Icn < 0.1A | Ic Wiring Missing（跳过 19/22） |
| 17 | PF_A ∈ [-1, -0.9] | Ia Polarity Reversed |
| 18 | PF_B ∈ [-1, -0.9] | Ib Polarity Reversed |
| 19 | PF_C ∈ [-1, -0.9] | Ic Polarity Reversed |
| 20 | PF_A ∈ (-0.9, 0.9] | Ia Phase Shift |
| 21 | PF_B ∈ (-0.9, 0.9] | Ib Phase Shift |
| 22 | PF_C ∈ (-0.9, 0.9] | Ic Phase Shift |

---

## 二、2.5E4WY（2.5 Element 4 Wire Y）

接入：Ua、Uc、Ia、Ib、Ic；**电流检查逻辑与 3E4WY 完全相同**

### 电压侧（仅检查 A/C 两相；无相位错误检测）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 1 | Van < 0.1·VRATE **且** Vcn < 0.1·VRATE | Va & Vc Wiring Missing |
| 2 | Van ≤ 0.8·VRATE | Va Wiring Missing |
| 3 | Vcn ≤ 0.8·VRATE | Vc Wiring Missing |
| 4 | Van ∈ (0.8, 1.2)·VRATE **且** Vcn > 1.3·VRATE | Va-Vn Reversed |
| 5 | Vcn ∈ (0.8, 1.2)·VRATE **且** Van > 1.3·VRATE | Vc-Vn Reversed |

> 无多回路表对应此接线方式。

---

## 三、2E3W Network（2 Element 3 Wire Network）

接入：Ua、Uc、Ia、Ic；**多回路 4100/2100 电压/电流检查逻辑与 3E4WY 相同**

### 电压侧（单回路表，仅 A/C 两相；无相序检测）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 1 | Van < 0.1·VRATE **且** Vcn < 0.1·VRATE | Va & Vc Wiring Missing |
| 2 | Van ≤ 0.8·VRATE | Va Wiring Missing |
| 3 | Vcn ≤ 0.8·VRATE | Vc Wiring Missing |
| 4 | Van ∈ (0.8, 1.2)·VRATE **且** Vcn > 1.3·VRATE | Va-Vn Reversed（2100 用 AB） |
| 5 | Vcn ∈ (0.8, 1.2)·VRATE **且** Van > 1.3·VRATE | Vc-Vn Reversed（2100 用 AB） |

### 电流侧（单回路表，仅 A/C 两相）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 7 | Ian < 0.1A | Ia Wiring Missing |
| 9 | Icn < 0.1A | Ic Wiring Missing |
| 10 | PF_A ∈ [-1, -0.9] | Ia Polarity Reversed |
| 12 | PF_C ∈ [-1, -0.9] | Ic Polarity Reversed |

---

## 四、2E3W 1Phase（2 Element 3 Wire 1 Phase）

接入：Ua、Uc、Ia、Ic
设备差异：1310/2100/IIV3 采用 A+B；AcuVim3/4100/1320 采用 A+C

### 电压侧（优先级：缺失 > 反接；无相位错误）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 1 | Vca < 0.1·VRATE | Va & Vc Wiring Missing |
| 2 | Van < 0.8·VRATE | Va Wiring Missing |
| 3 | Vcn < 0.8·VRATE | Vc Wiring Missing |
| 4 | Van ∈ (0.8, 1.2)·VRATE **且** Vcn > 1.3·VRATE | Va-Vn Reversed |
| 5 | Vcn ∈ (0.8, 1.2)·VRATE **且** Van > 1.3·VRATE | Vc-Vn Reversed |

### 电流侧（多回路 4100；条件 7/8 缺失，9-12 独立；9/11 互斥，10/12 互斥）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 7 | Ian < 0.1A | Ia Wiring Missing（跳过 9/11） |
| 8 | Icn < 0.1A | Ic Wiring Missing（跳过 10/12） |
| 9 | PF_A ∈ [-1, -0.9] | Ia Polarity Reversed |
| 10 | PF_C ∈ [-1, -0.9] | Ic Polarity Reversed |
| 11 | PF_A ∈ (-0.9, 0.9] | Ia Phase Shift |
| 12 | PF_C ∈ (-0.9, 0.9] | Ic Phase Shift |

---

## 五、2E3W Delta（2LL Delta）

接入：Ua、Ub、Uc、Ia、Ic（4100 多回路默认，不可更改）；额定电压用**线电压**

### 电压侧（优先级：缺失 > 相序错误；无反接检测）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 1 | Vab < 0.1·VRATE **且** Vbc < 0.1·VRATE **且** Vca < 0.1·VRATE | Va & Vb & Vc Wiring Missing |
| 2 | Vab < 0.1·VRATE | Va & Vb Wiring Missing |
| 3 | Vbc < 0.1·VRATE | Vb & Vc Wiring Missing |
| 4 | Vca < 0.1·VRATE | Va & Vc Wiring Missing |
| 5 | Vbc ∈ (0.8, 1.2)·VRATE **且** (Vab < 0.8·VRATE **或** Vca < 0.8·VRATE) | Va Wiring Missing |
| 6 | Vca ∈ (0.8, 1.2)·VRATE **且** (Vab < 0.8·VRATE **或** Vbc < 0.8·VRATE) | Vb Wiring Missing |
| 7 | Vab ∈ (0.8, 1.2)·VRATE **且** (Vbc < 0.8·VRATE **或** Vca < 0.8·VRATE) | Vc Wiring Missing |
| 8 | ABC: ∠Vbc ∈ [120±20°] **且** ∠Vca ∈ [240±20°]；ACB: 反之 | Phase Order Error |

> **v1.05 变更**：条件 5/6/7 次级判断由 `AND` 改为 `OR`。

### 电流侧（多回路 4100；仅 A/C 两相；条件 9/10 缺失，11-14 独立；11/13 互斥，12/14 互斥）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 9 | Ian < 0.1A | Ia Wiring Missing（跳过 11/13） |
| 10 | Icn < 0.1A | Ic Wiring Missing（跳过 12/14） |
| 11 | ABC: ∠IA ∈ [150±20°]；ACB: ∠IA ∈ [210±20°] | Ia Polarity Reversed |
| 12 | ABC: ∠IC ∈ [270±20°]；ACB: ∠IC ∈ [90±20°] | Ic Polarity Reversed |
| 13 | ABC: ∠IA ∉ [330±20°]；ACB: ∠IA ∉ [30±20°] | Ia Phase Shift |
| 14 | ABC: ∠IC ∉ [90±20°]；ACB: ∠IC ∉ [270±20°] | Ic Phase Shift |

---

## 六、3E3W Delta（3LL Delta）

**电压和电流检查逻辑完全同 2E3W Delta**（见第五节）。

---

## 七、1E2W（1LN）

接入：Ua、Ia

| 检测类型 | 规则 | 结果 |
|---------|------|------|
| 电压缺失 | Van < 0.1·VRATE | Va Wiring Missing |
| 电流缺失 | Ian < 0.1A | Ia Wiring Missing（跳过反接） |
| 电流反接 | PF_A ∈ [-1, -0.9] | Ia Polarity Reversed |

---

## 八、HighLeg Delta（3E4W Delta）

接入：Ua、Ub、Uc、Ia、Ib、Ic；额定电压用**线电压**
**电流检查逻辑完全同 2E3W Delta**（见第五节）

### 特殊额定值约定

```
VRATE         = 线电压额定值（用户配置）
Van_rated     = 0.5 × VRATE          （A/C 低相）
Vbn_rated     = (√3/2) × VRATE ≈ 0.866 × VRATE  （B 高腿）
线电压：Vab = Vbc = Vca = VRATE
```

### 电压侧（优先级：缺失 > 反接（含互换）；条件 1-10 串行，11-13 独立）

| 条件 | 规则 | 结果 |
|-----|------|------|
| 1 | Vab/Vbc/Vca < 0.1·VRATE（全部） | Va & Vb & Vc Wiring Missing |
| 2 | Vab < 0.1·VRATE | Va & Vb Wiring Missing |
| 3 | Vbc < 0.1·VRATE | Vb & Vc Wiring Missing |
| 4 | Vca < 0.1·VRATE | Va & Vc Wiring Missing |
| 5 | Van ≤ 0.4·VRATE（即 0.8·Van_rated） | Va Wiring Missing |
| 6 | Vbn ≤ 0.69·VRATE（即 0.8·Vbn_rated） | Vb Wiring Missing |
| 7 | Vcn ≤ 0.4·VRATE（即 0.8·Vcn_rated） | Vc Wiring Missing |
| 8 | Van ∈ [0.4, 0.6]·VRATE **且** Vbn/Vcn ∈ (0.8, 1.2)·VRATE | Va-Vn Reversed |
| 9 | Vbn ∈ [0.69, 1.04]·VRATE **且** Van/Vcn ∈ (0.8, 1.2)·VRATE | Vb-Vn Reversed |
| 10 | Vcn ∈ [0.4, 0.6]·VRATE **且** Van/Vbn ∈ (0.8, 1.2)·VRATE | Vc-Vn Reversed |
| 11 | Van ∈ [0.8,1.2]·Vbn_rated **且** Vbn ∈ [0.8,1.2]·Van_rated **且** Vcn ∈ [0.8,1.2]·Vcn_rated | Va-Vb Reversed（A/B 互换） |
| 12 | Van ∈ [0.8,1.2]·Van_rated **且** Vbn ∈ [0.8,1.2]·Vcn_rated **且** Vcn ∈ [0.8,1.2]·Vbn_rated | Vb-Vc Reversed（B/C 互换） |
| 13 | Van ∈ [0.8,1.2]·Vcn_rated **且** Vbn ∈ [0.8,1.2]·Vbn_rated **且** Vcn ∈ [0.8,1.2]·Van_rated | Va-Vc Reversed（A/C 互换） |

---

## 九、相序检查算法（Sheet 9）

### 不检测相序的接线方式

2.5E4WY、1E2W、2E3W 1Phase、2E3W Network（单回路三相表）
→ 直接输出 `voltagePhaseOrder = 2`（无相序概念）

### Y 接（3E4WY、2E3W Network 多回路）

检测输入：相电压有效值 + 相角 + 相电压平均值 vLNAvg

| Step | 条件 | 不满足时输出 |
|------|------|------------|
| 1 电压有效性 | Van/Vbn/Vcn > 0.1·VRATE（全部） | voltagePhaseOrder = 2 |
| 2 幅值对称性 | 各相与 vLNAvg 偏差 < 10% | voltagePhaseOrder = 2 |
| 3 ABC 正相序 | ∠Vb ∈ [240±20°] **且** ∠Vc ∈ [120±20°] | → 尝试 ACB 判决 |
| 3 ACB 负相序 | ∠Vb ∈ [120±20°] **且** ∠Vc ∈ [240±20°] | voltagePhaseOrder = 2 |

输出：`0` = ABC；`1` = ACB；`2` = 无法判断

### Δ 接（2E3W Delta、3E3W Delta、3E4W Delta）

检测输入：线电压有效值 + 线电压角度 + 线电压平均值 vLLAvg

| Step | 条件 | 不满足时输出 |
|------|------|------------|
| 1 线电压有效性 | Vab/Vbc/Vca > 0.1·VRATE（全部） | voltagePhaseOrder = 2 |
| 2 幅值对称性 | 各线电压与 vLLAvg 偏差 < 10% | voltagePhaseOrder = 2 |
| 3 ABC 正相序 | ∠Vbc ∈ [240±20°] **且** ∠Vca ∈ [120±20°] | → 尝试 ACB 判决 |
| 3 ACB 负相序 | ∠Vbc ∈ [120±20°] **且** ∠Vca ∈ [240±20°] | voltagePhaseOrder = 2 |

---

## 十、设备支持矩阵（Sheet 8）

| 接线方式 | 接入电压/电流 | 特殊说明 |
|---------|------------|---------|
| 3E4WY | Ua/Ub/Uc/Ia/Ib/Ic | 3E4WY 与 2.5E4WY 电流检查方案相同；额定电压用相电压 |
| 2.5E4WY | Ua/Uc/Ia/Ib/Ic | 无多回路表对应 |
| 2E3W Network | Ua/Uc/Ia/Ic | 2100/4100 多回路电压检查同 3E4WY；不判断相序错误 |
| 2E3W 1Phase | Ua/Uc/Ia/Ic | 1310/2100/IIV3 用 A+B；AcuVim3/4100/1320 用 A+C |
| 2E3W Delta | Ua/Ub/Uc/Ia/Ib/Ic（N接B） | 4100 多回路每两通道为 1 user；1320/4100 只测 Ia/Ic；额定电压用线电压 |
| 3E3W Delta | Ua/Ub/Uc/Ia/Ib/Ic（N浮空） | 电压/电流检查逻辑同 2E3W Delta |
| 3E4W Delta | Ua/Ub/Uc/Ia/Ib/Ic | 电流检查逻辑同 2E3W Delta |
| 1E2W | Ua/Ia | — |

---

## 版本变更记录

| 版本 | 主要变更 |
|------|---------|
| v1.03（2026-04-09） | HMI1-7 Sprint3 基准版本 |
| v1.05（当前） | 2E3W Delta 条件 5/6/7 改为 OR 逻辑；新增 HighLeg Delta（3E4W Delta）完整规格；补充相序算法和设备支持矩阵 |
