# 快速精度测试算法说明文档

> 来源脚本：`acuvimseries_fast_test.py`（AcuRev1320 项目）

---

## 1. 整体流程

```
主入口
  └─ run_precision_measure_script(test_type, wire_type)
       ├─ 切换设备至交流界面（寄存器指令）
       ├─ 归零源档位
       └─ PrecisionMeasure.select_test_case(test_type, wire_type)
            ├─ 按 test_type 选择 CT 类型对应的测点 Excel 文件和 Sheet
            └─ select_wire_type(case_path, sheet_name, wire_type)
                 ├─ 写接线方式到寄存器（0x1042）
                 ├─ 读取测点 Excel 数据
                 └─ fast_precision_measure_by_<wire_type>(file_path, input_list)
                      ├─ 逐行读取测试参数（电压、电流、相角、采样设置）
                      ├─ 检测频率变更 → 触发设备重启
                      ├─ 计算预期功率 P/Q/S
                      ├─ 驱动源输出目标电压/电流/角度
                      ├─ 采样 N 次读取寄存器
                      ├─ 计算精度误差（min/max/avg vs 预期）
                      └─ 写入结果 Excel
```

---

## 2. 参数体系

### 2.1 CT 类型（test_type）

| 值 | 类型 | 说明 |
|---|---|---|
| 0 | mA | 100 mA CT，读取对应 Sheet |
| 1 | None | 无 CT 变换 |
| 2 | mV | 333 mV CT |
| 3 | rct | 比率 CT，读数需 × 10 倍乘子 |

### 2.2 接线方式（wire_type）

| 值 | 代码 | 接线描述 | 寄存器写入值 | 有效量 |
|---|---|---|---|---|
| 0 | 1E2W1P | 1元件 2线 单相 | 0 | Ua, Ia, Pa, Qa, Sa, P_sys |
| 1 | 2E3W1P | 2元件 3线 单相 | 1 | Ua, Uc, Ia, Ic, Pa, Pc, P_sys |
| 2 | 2E3WD | 2元件 3线 角接（Aron法） | 2 | Uab, Ubc, Uca, Ia, Ic, P_sys |
| 3 | 2E3WN | 2元件 3线 网络 | 3 | Uab, Ubc, Uca, Ia, Ib, Ic, P_sys |
| 4 | 3E4WY | 3元件 4线 星接（完整） | 4 | Ua/b/c, Ia/b/c, Pa/b/c, Qa/b/c, Sa/b/c, P_sys |
| 5 | 3E3WD | 3元件 3线 角接 | 5 | Uab, Ubc, Uca, Ia/b/c, P_sys |
| 6 | 3E4WY-P | 3E4WY 放开版（无功/视在放宽） | 4 | 同 3E4WY |

### 2.3 测试输入参数（每行测试点）

每行测点从测点 Excel 中读取，包含：

| 字段 | 含义 |
|---|---|
| case_id | 测试用例编号 |
| ua, ub, uc | 相电压（V） |
| ia, ib, ic | 相电流（A） |
| ua_p, ub_p, uc_p | 电压相角（°） |
| ia_p, ib_p, ic_p | 电流相角（°） |
| freq | 频率（50 / 60 Hz） |
| sample_cnt | 采样次数 |
| sample_interval | 采样间隔（ms） |
| voltage_accuracy | 电压允许误差（如 0.001 = 0.1%） |
| current_accuracy | 电流允许误差 |
| phase_angle_accuracy | 相角允许误差（°） |
| active_power_accuracy | 有功功率允许误差 |

---

## 3. 功率计算算法

### 3.1 单相（Per Phase）

```
Pa = Ua × Ia × cos(θ_Ua − θ_Ia) / 1000    [kW]
Qa = Ua × Ia × sin(θ_Ua − θ_Ia) / 1000    [kvar]
Sa = Ua × Ia / 1000                         [kVA]
```

适用接线：1E2W1P, 2E3W1P, 3E4WY

### 3.2 三相星接系统功率（3E4WY）

```
P_sys = Pa + Pb + Pc
Q_sys = Qa + Qb + Qc
S_sys = Sa + Sb + Sc
```

### 3.3 线电压计算（所有角接接线）

```python
Uab = √(Ua² + Ub² - 2·Ua·Ub·cos(θ_Ua − θ_Ub))
Ubc = √(Ub² + Uc² - 2·Ub·Uc·cos(θ_Ub − θ_Uc))
Uca = √(Uc² + Ua² - 2·Uc·Ua·cos(θ_Uc − θ_Ua))
```

### 3.4 两元件三线角接（2E3WD）— Aron 法

两表法（Aron）：测量 Uab/Ia 和 Ubc(或Ucb)/Ic

```
P_sys = Uab·Ia·cos(θ_Uab − θ_Ia) + Ubc·Ic·cos(θ_Ubc − θ_Ic)   [/1000, kW]
Q_sys = Uab·Ia·sin(θ_Uab − θ_Ia) + Ubc·Ic·sin(θ_Ubc − θ_Ic)   [/1000, kvar]
S_sys = √(P_sys² + Q_sys²)
```

> 注：线电流相角相比相电流相角**滞后 30°**（角接变换），源输出相角调整后，寄存器读取的线电流角度相应偏移。

### 3.5 三元件三线角接（3E3WD）

```python
p_sys, q_sys, s_sys = calculate_3e4wd_power(ua, ub, uc, ua_p, ub_p, uc_p, ia, ib, ic, ia_p, ib_p, ic_p)
```

内部通过线电压向量和线电流向量计算，电流角度同样有 30° 偏移处理。

---

## 4. 精度计算算法

### 4.1 有预期精度（电压、电流、有功）

```python
def get_accuracy_res_by_exp_accuracy(expected, measured_list, accuracy_threshold):
    errors = [(m - expected) / expected for m in measured_list]
    min_val = min(measured_list);  min_acc = min(errors)
    max_val = max(measured_list);  max_acc = max(errors)
    avg_val = mean(measured_list); avg_acc = mean(errors)
    pass = abs(avg_acc) <= accuracy_threshold
    return (min_val, min_acc, max_val, max_acc, avg_val, avg_acc, pass)
```

### 4.2 无预期精度（无功、视在——直接对比偏差）

```python
def get_accuracy_res_by_not_exp_accuracy(expected, measured_list):
    # 只计算 min/max/avg，无合格判断
    return (min_val, max_val, avg_val)
```

### 4.3 相角精度（相角差允许误差）

```python
def get_accuracy_res_by_phase_angle(expected_angle, measured_list, threshold):
    errors = [m - expected_angle for m in measured_list]
    # 处理 360° 跨越（如 359° vs 1°）
    ...
    return (min_val, min_err, max_val, max_err, avg_val, avg_err, pass)
```

---

## 5. 采样流程

```python
for _ in range(sample_cnt):
    # 按采样间隔等待
    time.sleep(sample_interval / 1000)
    # 通过 Modbus 批量读取寄存器
    ua = read_register(0x9309)
    ub = read_register(0x930B)
    ...
    # 追加到各量的样本列表
```

每个物理量返回 `sample_cnt` 个测量值的列表，后续统计 min/max/avg。

---

## 6. 设备控制

### 6.1 频率切换 + 设备重启

当测点要求频率（50/60 Hz）与设备当前设置不同时：

**旧方案（硬件继电器）：**
```python
pow_off_device()   # 断电
time.sleep(5)
pow_on_device()    # 上电
```

**新方案（软重启，本工具采用）：**
```python
# 写接线方式寄存器
modbus_client.write_register(0x1042, wire_type_value)
# 软重启
modbus_client.write_register(0x1147, 0x0001)
time.sleep(5)   # 等待重启完成
```

### 6.2 接线方式设置

通过写 Modbus 寄存器 `0x1042` 设置接线方式（值见 §2.2 表格）。

---

## 7. Modbus 地址表（200 ms 刷新组）

| 物理量 | 寄存器地址 | 备注 |
|---|---|---|
| 频率 | 0x9307 | |
| Ua / Ub / Uc | 0x9309 / 0x930B / 0x930D | 相电压 |
| θ_Ua / θ_Ub / θ_Uc | 0x930F / 0x9311 / 0x9313 | 电压相角 |
| Uab / Ubc / Uca | 0x9317 / 0x9319 / 0x931B | 线电压 |
| Ia / Ib / Ic / In | 0x9325 / 0x9327 / 0x9329 / 0x932B | 相电流 |
| θ_Ia / θ_Ib / θ_Ic | 0x932D / 0x932F / 0x9331 | 电流相角 |
| Pa / Pb / Pc / P_total | 0x9337 / 0x9339 / 0x933B / 0x933D | 有功功率 |
| Qa / Qb / Qc / Q_total | 0x933F / 0x9341 / 0x9343 / 0x9345 | 无功功率 |
| Sa / Sb / Sc / S_total | 0x9347 / 0x9349 / 0x934B / 0x934D | 视在功率 |
| PFa / PFb / PFc / PF_total | 0x934F / 0x9351 / 0x9353 / 0x9355 | 功率因数 |
| **接线方式** | **0x1042** | 写入 |
| **清除电能** | **0x1130** | 写入 |
| **软重启** | **0x1147** | 写 0x0001 重启 |

> 不同电表的地址表不同，GUI 工具中该表由用户配置（支持导入/手动填写）。

---

## 8. 结果 Excel 输出结构

每种接线方式生成独立 Excel 文件，文件名格式：
```
Precision_Measure_<接线类型>_<时间戳>.xlsx
```

每行为一个测试点，列结构：
```
[测试输入参数] | [各物理量: 预期值 | 最小值 | 最大精度 | 最大值 | 最大精度 | 平均值 | 平均精度 | 合格?]
```

---

## 9. 接线方式测量量汇总

| 接线方式 | 电压量 | 电流量 | 功率量 | 线电压 |
|---|---|---|---|---|
| 1E2W1P | Ua | Ia | Pa, P_sys | — |
| 2E3W1P | Ua, Uc | Ia, Ic | Pa, Pc, P_sys | Uca |
| 2E3WD | — | Ia, Ib, Ic | P_sys | Uab, Ubc, Uca |
| 2E3WN | — | Ia, Ib, Ic | P_sys | Uab, Ubc, Uca |
| 3E4WY | Ua, Ub, Uc | Ia, Ib, Ic, In | Pa~Pc, Qa~Qc, Sa~Sc, P/Q/S_sys | Uab, Ubc, Uca |
| 3E3WD | — | Ia, Ib, Ic | P_sys, Q_sys, S_sys | Uab, Ubc, Uca |
| 3E4WY-P | Ua, Ub, Uc | Ia, Ib, Ic, In | Pa~Pc, Qa~Qc, Sa~Sc, P/Q/S_sys | Uab, Ubc, Uca |

---

## 10. 测点文件说明

测点 Excel（位于 `test_point/` 目录）每行为一个测试点，关键列：

| 列 | 说明 |
|---|---|
| A | 测试用例 ID |
| B | 接线方式代码（wire_type） |
| C | Ua / Ub / Uc（V） |
| D | Ia / Ib / Ic（A） |
| E | 电压相角 θ_Ua / θ_Ub / θ_Uc（°） |
| F | 电流相角 θ_Ia / θ_Ib / θ_Ic（°） |
| G | 频率（50 / 60 Hz） |
| H | 采样次数 |
| I | 采样间隔（ms） |
| J | 各精度指标（电压/电流/相角/有功） |

文件：
- `acuvimseries_test_case.xlsx`：AcuVIM 系列测点
- `rev1320_test_case.xlsx`：AcuRev1320 测点
