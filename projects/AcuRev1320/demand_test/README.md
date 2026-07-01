# AcuRev-1320 需量（Demand）测试

电表需量功能的自动化测试模块：由上位机控源（功率源）输出指定的电压/电流/相位/频率，
按设定的需量窗口算法等待电表完成需量累计，再通过 Modbus 回读电表需量寄存器，
与按控源输入值计算出的期望需量做精度比对，逐条用例把实测值与 Pass/Fail 写入结果 Excel。

覆盖 **6 种接线方式 × 2 种需量触发（窗口）方式**，支持固定窗口（Fixed Window）与
滑动窗口（Sliding Window）两种需量算法的功率需量（P/Q/S）与电流需量（Ia/Ib/Ic/In）验证。

---

## 目录结构

| 文件 | 说明 |
|---|---|
| `test_demand.py` | 测试主体：控源 + Modbus 回读 + 需量计算 + 精度比对 + 结果落盘，pytest 入口 `test_demand` |
| `conftest.py` | 注册 `--measure` 命令行选项（测量模式 mv/ma/rct） |
| `demand_table_heading.py` | 结果 Excel 的表头列名定义（`TableTitle`，按接线方式各一套，结构相同） |
| `demand_test_case.xlsx` | 用例输入数据源，含 3 个 sheet：`test_case_mV` / `test_case_mA` / `test_case_rct` |

运行后自动生成结果目录（不纳入版本管理）：

```
<本目录>/<YYYYMMDD>/Demand_<接线方式>/demand_test_res_<时间戳>/
    └── Fixed_Demand_Test_<接线方式>_<时间戳>.xlsx     # 或 Sliding_...
```

---

## 三个测试维度

测试由 `run_demand_test_script(test_type, wire_type, demand_type)` 驱动，三个参数：

### 1. `test_type` —— 测量模式（决定读哪个 sheet）

| 取值 | 对应 sheet | 含义 |
|---|---|---|
| 0 | `test_case_mA` | mA（电流互感器副边电流档） |
| 1 | `test_case_mA` | mA |
| 2 | `test_case_mV` | mV（罗氏线圈 / 电压档） |
| 3 | `test_case_rct` | rct |

> 映射以代码 `select_test_case` 为准（0、1 均指向 mA sheet）。

### 2. `wire_type` —— 接线方式

| 取值 | 接线方式 | Excel 第 6 列筛选标识 |
|---|---|---|
| 0 | 1E2W1P | `1E2w1p` |
| 1 | 2E3W1P | `2E3w1p` |
| 2 | 2E3WD  | `2E3wD` |
| 3 | 2E3WN  | `2E3wN` |
| 4 | 3E4WY  | `3E4wY` |
| 5 | 3E4WD  | `3E4wD` |

### 3. `demand_type` —— 需量触发（窗口）方式

| 取值 | 含义 | 对应方法 |
|---|---|---|
| 0 | Fixed（固定窗口） | `fixed_demand_test_by_*` |
| 1 | Sliding（滑动窗口） | `sliding_demand_test_by_*` |

---

## 用例数据（`demand_test_case.xlsx`）列约定

每个 sheet 第 1 行为表头，从第 2 行起每行一条用例，列含义（0 基索引）：

| 列 | 字段 | 说明 |
|---|---|---|
| 0 | `case_id` | 用例编号 |
| 1 | `voltage` | 电压值（三相同值输出） |
| 2 | `angle` | 电压-电流相位角 |
| 3 | `current` | 电流值（三相同值输出） |
| 4 | `freq` | 频率 Hz |
| 5 | `wire_type` | 接线方式标识（用于按 `wire_type` 过滤，见上表） |
| 6 | `demand_method` | 需量方法（用于按 `demand_type` 过滤） |
| 7 | `interval` | 需量窗口间隔（分钟，1~30） |
| 8 | `update_rate` | 需量更新速率（分钟，1~30） |
| 9 | `demand_trigger` | 需量重置方式：0 重设参数 / 1 时间触发 / 2 清最大需量 |
| 10 | `demand_accuracy` | 比对容差（相对误差，如 0.001） |
| 11 | `sample_cnt` | 抽样次数 |
| 12 | `sample_interval` | 抽样间隔 |

---

## 测试逻辑

### 整体流程（`run_demand_test_script`）

1. 切换电表到交流测量界面、档位切换归零（自动）。
2. 创建 `DemandTest`（内部建立 slave_id=1 的 Modbus 连接，初始化结果目录）。
3. `select_test_case` → 按 `test_type` 选 sheet → 按 `wire_type` 过滤接线方式行
   → 按 `demand_type` 过滤触发方式行 → 设置电表接线方式 → 调用对应
   `fixed_/sliding_demand_test_by_<接线方式>` 执行。
4. 关闭 Modbus 连接、关源、切回默认界面、档位归零（手动），打印总耗时。

### 单条用例验证流程（以 `fixed_demand_test_by_3e4wy` 为例）

对接线方式过滤出的每一行用例：

1. **算期望初始需量**：按控源输入的电压/电流/相位，计算系统有功 P、视在 S，
   再由 `Q = √(S² − P²)` 求无功 Q，电流需量取输入电流值。
2. **控源输出初始值** → `set_demand_para`（设方法/间隔/更新率）→ `sleep(300s)`
   → 校验需量已清零。
3. **触发需量重置**（按 `demand_trigger`：重设参数 / 时间触发 / 清最大需量）→ 再校验清零。
4. **对齐窗口时刻**：`get_wait_seconds` 计算到最近的 `interval` 整数倍时刻需等待的秒数并 `sleep`，
   再 `sleep(interval×60)` 等待第一次需量上报。
5. **第一周期比对**：电压<9.5 或电流=0 时只校验清零；否则按容差比对
   P/Q/S/Ia/Ib/Ic/In（`check_demand_power_current_is_pass`）。
6. **第二周期比对**：等待半个窗口后把电压、电流降为 0.5 倍，按"上半窗 + 下半窗加权"
   重算期望需量，再等半个窗口后比对。
7. **结果落盘**：把输入参数、期望值/实测值成对、Pass/Fail 写入结果 Excel
   （列结构见 `demand_table_heading.py`，含 1st/2nd/3rd 三个周期段）。

固定窗口与滑动窗口的差异在于需量累计算法不同，分别由 `fixed_*` / `sliding_*` 两族方法实现，
比对与落盘框架一致。

> 比对结果写入结果 Excel 的 `is_pass_*` 列；pytest 层在底层抛异常时判该组合失败。

#### 源在线探针（当前仅 `fixed_demand_test_by_3e4wy`）

有压有流的用例加源后，会立即回读表实时 RMS 电压/电流，命令值非零而实测明显偏低即判定
源未开/无输出，抛 `SourceControlError` 快速失败，避免白等整个需量周期。

> ⚠️ **覆盖现状**：该探针目前**只接入了 `fixed_demand_test_by_3e4wy` 这 1 个方法**，
> 其余 11 个 `*_demand_test_by_*` 组合尚未接入——跑那些组合时不会触发探针（源没开仍会空等）。
> 因各接线方式的期望电压/电流分相不同（2E3W/1E2W 部分相为 0），需按接线定制探针的期望列表
> 后再逐个推广。

---

## 执行命令

> ⚠️ 需真实**功率源 + AcuRev-1320 电表**在环。单条接线方式动辄数十分钟（含多次 5 分钟级
> `sleep`），且会持续改变设备状态，用例统一打了 `slow` + `destructive` 标记。

在仓库根目录 `autotest/` 下执行。统一走 pytest，一条命令即可选定**接线 + 触发 + 测量模式**三项，无需环境变量、无需改代码。

### 三项配置与命令参数对应

| 维度 | 配置方式 | 取值 |
|---|---|---|
| `wire_type`（接线） | `-k` 选参数化 id 前半段 | `1e2w1p` / `2e3w1p` / `2e3wd` / `2e3wn` / `3e4wy` / `3e4wd` |
| `demand_type`（触发） | `-k` 选参数化 id 后半段 | `fixed` / `sliding` |
| `measure_mode`（测量模式） | 命令行选项 `--measure` | `mv`（缺省）/ `ma` / `rct`（也接受 `0`/`1`/`2`/`3`） |

用例按"接线方式 × 触发方式"参数化为 12 个独立用例项，id 形如 `3e4wy-fixed`。

### 一条命令配齐三项（推荐）

```bash
# 接线=3E4WY、触发=fixed、测量模式=mV
pytest projects/AcuRev1320/demand_test/test_demand.py -k "3e4wy-fixed" --measure=mv -v

# 接线=2E3WD、触发=sliding、测量模式=mA
pytest projects/AcuRev1320/demand_test/test_demand.py -k "2e3wd-sliding" --measure=ma -v
```

### 其它常用

```bash
# 跑全部 12 个组合（测量模式缺省 mV）
pytest projects/AcuRev1320/demand_test/test_demand.py -v

# 不带 --measure 即默认 mV
pytest projects/AcuRev1320/demand_test/test_demand.py -k "3e4wy-fixed" -v

# 按标记筛选
pytest projects/AcuRev1320/demand_test/test_demand.py -m "destructive" --measure=mv -v

# 仅查看会收集到哪些用例（不真正执行，不需要硬件）
pytest projects/AcuRev1320/demand_test/test_demand.py --collect-only -q
```

> `--measure` 只作用于本次命令，不留任何进程/环境状态，跑完自动回到缺省 mV；跨平台写法一致（无需 PowerShell 的 `$env:` / Linux 的行内变量）。
> 取值非法（如 `--measure=xx`）会在开跑前直接报 `UsageError` 提示，不会带着错配置空跑。


---

## 结果产物

每次运行按接线方式生成独立结果 Excel，路径见上文"目录结构"。
表头列含：用例编号、容差、输入电压/角度/电流/频率、接线方式、需量方法、间隔、更新率、触发方式，
以及 P/Q/S/Ia/Ib/Ic/In 在第 1/2/3 周期的"期望值 + 实测值"成对列与各周期 `is_pass`。
