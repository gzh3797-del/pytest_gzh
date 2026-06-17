# DataCollect — Metering 数据采集与比对

**支持设备：** AcuRev-4100 / AcuRev-4110、AcuRev-2100、Acuvim3、AcuvimIIW

对接 AcuHMI 平台，自动采集设备 Metering 页面数据，与 Modbus TCP 直读值逐参数比对，生成 Excel 报告。

- 脚本：`metering.py`（三步流水线核心）+ `test_metering_collect.py`（pytest 入口）
- 配置入口：`metering.py` 顶部常量（或 `test_metering_collect.py` 中同名常量覆盖）

---

## 1. 使用说明

### 1.1 环境准备

```powershell
# 安装 Python 依赖
pip install playwright pymodbus openpyxl

# 安装 Playwright 浏览器内核
python -m playwright install chromium
```

### 1.2 配置修改

打开 `DataCollect/metering.py`，修改顶部配置块：

```python
# ╔══════════════════════════════════════════════════════╗
# ║              运行前根据实际环境修改                  ║
# ╚══════════════════════════════════════════════════════╝
BASE_URL = "https://192.168.2.9"    # AcuHMI 平台地址
USERNAME = "admin"                   # 登录账号
PASSWORD = "Admin@110001"            # 登录密码
HEADED   = False   # 改 True 可在浏览器中观察操作过程
TOL_REL  = 0.01   # 相对容差 1%
TOL_ABS  = 0.05   # 绝对容差
```

> Modbus 设备的 IP / Port / Unit ID 由脚本自动从页面 **Settings → Connection** 抓取，无需手动填写。
> Status = off 的设备会被自动跳过，无数据（No Data）的通道也会自动跳过。

### 1.3 运行脚本（命令行直接运行）

工作目录统一设为 `Protocols`：

```powershell
cd C:\JrJ\auto\autotest\Protocols

# ① 完整三步（采集 → 匹配 → 比对），推荐
python DataCollect/metering.py

# ② 单步运行（调试用）
python DataCollect/metering.py --collect   # 仅采集 → metering_collect.json
python DataCollect/metering.py --match     # 仅匹配 → metering_register_match.csv
python DataCollect/metering.py --compare   # 仅比对 → metering_compare.xlsx

# ③ 有头模式（浏览器可见）：修改 metering.py 顶部 HEADED = True 后运行
```

### 1.4 运行 pytest 自动化测试

```powershell
cd C:\JrJ\auto\autotest\Protocols

# ① 运行全部 13 条用例
python -m pytest DataCollect -v

# ② 生成 HTML 报告（需 pip install pytest-html）
python -m pytest DataCollect -v --html=reports/metering_report.html --self-contained-html

# ③ 只跑采集阶段用例
python -m pytest DataCollect -k "TestCollect" -v

# ④ 只跑比对阶段用例
python -m pytest DataCollect -k "TestCompare" -v
```

### 1.5 查看报告

测试结束后，以下文件自动生成在 `Protocols/reports/`：

| 文件 | 说明 |
|---|---|
| `metering_collect.json` | 各设备 Metering 页面采集数据（原始） |
| `metering_register_match.csv` | 采集参数 → 寄存器地址映射明细 |
| `metering_register_match.md` | 同数据 Markdown 格式 |
| `metering_compare.xlsx` | 汇总 + 逐参数比对明细 + **用例Mapping** 三个 Sheet |

---

## 2. 脚本逻辑与模块说明

### 2.1 整体工作原理（三步流水线）

```
Step 1 采集（collect）
  Playwright → 登录 AcuHMI → 遍历所有 Status=on 的 Modbus TCP 设备
            → 逐个打开 Metering 视图（Realtime/Demand/Energy/THD/Sequence）
            → 遍历下拉选项 → 跳过 No Data 通道 → 读取参数表格
            → 保存 metering_collect.json

Step 2 匹配（match_registers）
  按设备名自动选择 AcuCloud 模板 xlsx（4 种型号）
  → 解析 blockParams sheet（地址/描述/数据类型/Scale）
  → 逐参数 _param_kw + col_kw 关键词匹配最优寄存器条目
  → 精确/模糊/未匹配分类 → 保存 CSV/MD

Step 3 比对（compare）
  按 Connection 信息建立 Modbus TCP 连接
  → 逐参数从模板地址读寄存器（pymodbus）→ 乘以 Scale 因子
  → 与网页采集值对比（相对/绝对容差）→ 写 metering_compare.xlsx
```

### 2.2 脚本模块结构

```
DataCollect/
├── metering.py                       # 核心脚本
│   ├── 配置层（顶部常量）
│   │   ├── BASE_URL / USERNAME / PASSWORD / HEADED / TOL_REL / TOL_ABS
│   │   └── DEVICE_TEMPLATES          设备名关键词 → 模板 xlsx 映射表
│   │
│   ├── 采集层（Step 1）
│   │   ├── _login()                  Playwright 登录 AcuHMI
│   │   ├── _list_modbus_tcp_devices() 扫描 Physical Devices，过滤 Status=off
│   │   ├── _has_no_data()            检测通道是否显示 No Data
│   │   ├── _read_tables()            解析 el-table 表格内容
│   │   ├── _read_view_all_dropdowns() 遍历下拉选项，跳过 No Data
│   │   └── collect()                 主入口，写 JSON
│   │
│   ├── 匹配层（Step 2）
│   │   ├── _load_template()          zipfile 解析 xlsx blockParams sheet
│   │   ├── _param_kw()               参数名 → 模板关键词（按长度降序避免误匹配）
│   │   ├── _match_entry()            关键词组合匹配 + THD/demand 过滤 + 降级回退
│   │   └── match_registers()         主入口，写 CSV/MD
│   │
│   ├── 比对层（Step 3）
│   │   ├── _modbus_read()            pymodbus 读寄存器，支持 float32/float64/uint32/int16
│   │   ├── _in_tol()                 容差判定（相对+绝对）
│   │   └── compare()                 主入口，写 xlsx（含用例Mapping sheet）
│   │
│   └── main()                        CLI 入口（--collect/--match/--compare 三步可选）
│
├── test_metering_collect.py          # pytest 用例入口
│   ├── apply_config fixture          注入配置（覆盖 metering 模块顶部常量）
│   ├── collected / matched / compared session-scoped fixtures（顺序执行三步）
│   └── TestCollect / TestMatch / TestCompare 三个 class，共 13 条用例
│
└── conftest.py                       # pytest 路径配置（sys.path）
```

### 2.3 关键设计决策

| 设计点 | 说明 |
|---|---|
| 模板解析方式 | 用 `zipfile + xml.etree.ElementTree` 解析 xlsx，避免 openpyxl 对 AutoFilter 的解析报错 |
| 设备跳过策略 | Status=off 的设备在采集阶段即跳过；No Data 通道不生成任何比对行，不计入 FAIL |
| 模板自动选择 | 按设备名关键词（不区分大小写/去空格/去连字符）自动匹配模板，未识别时回退 4100 并告警 |
| 参数名匹配 | PARAM_KW 按键长降序匹配，避免 "active power" 误匹配 "reactive power" |
| 列名降级回退 | user/input channel 先带 col_kw（Positive/Negative/Phase A 等）精确匹配，无结果再降级 |
| Power Factor 符号 | 对比时取 abs()，兼容设备以有符号 PF 上报而网页显示绝对值的情况 |
| Scale 因子 | 从模板 col L 读取，比对前乘以 Scale（如 IIW Power ×1E-3，THD ×0.01）|

### 2.4 模板自动选择规则

| 设备名关键词（不区分大小写） | 使用模板 |
|---|---|
| 含 `4110` 或 `4100` | AcuRev-4100.xlsx |
| 含 `2100` | AcuRev-2100.xlsx |
| 含 `acuvim3` 或 `vim3` | Acuvim3.xlsx |
| 含 `acuvimiiw`、`vimiiw`、`iiw` | AcuvimIIW.xlsx |
| 其他（未识别） | AcuRev-4100.xlsx（默认回退，日志提示警告） |

---

## 3. 用例清单（13 条）

| 用例ID | 用例名称（Class::method） | 测试步骤 | 验证点 |
|---|---|---|---|
| DC-001 | TestCollect::test_json_generated | 采集 | metering_collect.json 正常生成 |
| DC-002 | TestCollect::test_has_devices | 采集 | 至少采集到 1 台设备 |
| DC-003 | TestCollect::test_connection_info | 采集 | 每台设备均抓取到 Connection IP |
| DC-004 | TestCollect::test_metering_views | 采集 | 每台设备至少有 1 个 Metering 视图数据 |
| DC-005 | TestCollect::test_param_count | 采集 | 总采集参数条目 > 0 |
| DC-006 | TestMatch::test_csv_generated | 匹配 | metering_register_match.csv 正常生成 |
| DC-007 | TestMatch::test_match_count | 匹配 | 匹配条目 > 0 |
| DC-008 | TestMatch::test_exact_match_ratio | 匹配 | 精确匹配率 ≥ 80%（排除 N/A） |
| DC-009 | TestMatch::test_no_unmatched_excess | 匹配 | 未匹配条目 ≤ 20 |
| DC-010 | TestCompare::test_xlsx_generated | 比对 | metering_compare.xlsx 正常生成 |
| DC-011 | TestCompare::test_fail_count_zero | 比对 | FAIL = 0（所有值在容差内） |
| DC-012 | TestCompare::test_pass_rate | 比对 | 有效通过率 ≥ 95% |
| DC-013 | TestCompare::test_modbus_read_failure_zero | 比对 | Modbus 读取失败 = 0 |

> 用例Mapping 同步写入 `metering_compare.xlsx` 的「用例Mapping」Sheet。

---

## 4. 工程路径

```
C:\JrJ\auto\autotest\Protocols\DataCollect\
├── metering.py              # 核心脚本（三步流水线）
├── test_metering_collect.py # pytest 用例入口（13 条用例）
├── conftest.py              # pytest 路径配置
└── README.md                # 本文件

模板文件（只读）：C:\JrJ\auto\autotest\Protocols\template\AcuCloud 模板适配\
测试报告：        C:\JrJ\auto\autotest\Protocols\reports\metering_compare.xlsx
```
