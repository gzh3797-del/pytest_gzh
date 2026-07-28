# Pass Through 自动化测试

验证 AcuHMI-1-7 网关的 **Pass Through（透传）** 协议：网关把 Modbus 请求**透明转发**到下游真实电表，
对外用配置的 SlaveID 访问。本套用例覆盖**功能 / 边界 / 业务数据正确性 / 并发**。

- 脚本：`test_pass_through_business.py`（单文件、自包含，9 条用例）
- 配置入口：上级 `../config.py` 的 `HMI_URL / HMI_USERNAME / HMI_PASSWORD / GATEWAY_IP`

---

## 1. 使用说明

### 1.1 环境准备

```powershell
# 安装 Python 依赖
pip install -r C:\JrJ\auto\autotest\requirements.txt

# 安装 Playwright 浏览器内核
python -m playwright install chromium
```

### 1.2 配置修改

打开 `C:\JrJ\auto\autotest\Protocols\config.py`，确认以下字段：

```python
GATEWAY_IP   = "192.168.2.8"          # 网关 Modbus TCP 地址（A 路透传入口）
HMI_URL      = "https://192.168.2.9"  # AcuHMI 网页地址
HMI_USERNAME = "admin"                # 网页登录账号
HMI_PASSWORD = "Admin@110001"         # 网页登录密码
```

同时在 `../register_tables/` 下为各设备型号准备寄存器表（见 1.5 节）。

### 1.3 运行测试

工作目录统一设为 `Protocols`：

```powershell
cd C:\JrJ\auto\autotest\Protocols

# ① 运行全部 9 条 Pass Through 用例
pytest "Pass Through" -v

# ② 生成 HTML 报告（需 pip install pytest-html）
pytest "Pass Through" -v --html=reports/pass_through_report.html --self-contained-html

# ③ 有头模式——浏览器可见，方便调试
$env:HEADED="1"; pytest "Pass Through" -v; Remove-Item Env:HEADED

# ④ 只跑某条用例
pytest "Pass Through" -k "case01" -v

# ⑤ 仅跑数据正确性用例
pytest "Pass Through" -k "pt_001 or pt_002 or pt_003" -v

# ⑥ 调整并发时长 / 透传就绪等待
$env:CONCURRENT_SECONDS="60"; $env:PT_SETTLE_MS="5000"
pytest "Pass Through" -v
```

### 1.4 查看报告

测试结束后，以下文件自动生成在 `Protocols/reports/`：

| 文件 | 说明 |
|---|---|
| `pass_through_compare.xlsx` | 汇总 + 逐寄存器明细 + **用例Mapping** 三个 Sheet |
| `pass_through_compare.json` | 同数据 JSON，供二次处理 |
| `pass_through_connection.json` | 各设备真实 IP / 端口 / ModbusID / 透传 SlaveID |
| `pass_through_missing_tables.json` | 缺寄存器表的型号清单（用于补表） |

### 1.5 寄存器模板（`../template/AcuCloud 模板适配/<Model>.xlsx`）

脚本直接读取 AcuCloud 模板 xlsx（与 DataCollect 共用同一份模板目录），无需单独维护 CSV 文件。

支持的型号及对应模板：

| 设备型号关键词 | 使用模板文件 |
|---|---|
| 4100 / 4110 | AcuRev-4100.xlsx |
| 2100 | AcuRev-2100.xlsx |
| 1300 | AcuRev-1300.xlsx |
| Acuvim3 / VIM3 | Acuvim3.xlsx |
| AcuvimIIW / IIW | AcuvimIIW.xlsx |
| AcuvimIIR / VIMIIR | AcuvimIIR.xlsx |

> 缺少模板的型号会自动 skip 并写入 `pass_through_missing_tables.json`，不影响其他设备继续测试。

---

## 2. 脚本逻辑与模块说明

### 2.1 整体工作原理

```
A 路（透传）：Modbus 主站 → 网关 GATEWAY_IP:502 + Pass Through 配置的 SlaveID → 透传到真实电表
B 路（直读）：Modbus 主站 → 真实电表 IP:502    + 原生 ModbusID                  → 读相同寄存器
                              ↑ IP/ModbusID 由网页 Physical Devices→Settings→Connection 实时抓取
比对 A ↔ B：验证透传是否忠实还原真实电表（逐寄存器，数值量算一致率）
```

关键差异（与 Device Mirror 的区别）：
- 透传读取**必须 Pass Through 已 Enable 且网关透传服务就绪**（脚本自动启用并等待 `PT_SETTLE_MS`）。
- A 路用**页面配置的 SlaveID**（101–247），B 路直连用**原生 ModbusID**。
- 读哪些寄存器由**型号寄存器表** `../register_tables/<Model>.csv` 决定；缺表的设备自动 skip。

### 2.2 脚本模块结构

```
test_pass_through_business.py
├── 配置层（顶部常量）
│   ├── GATEWAY_IP / BASE_URL / USERNAME / PASSWORD   从 config.py 读取
│   ├── REPORT / REG_TABLES_DIR                       报告目录 / 寄存器表目录
│   └── 容差、超时、等待、重试等运行参数
│
├── Modbus 工具层
│   ├── ModbusClient（class）                          pymodbus TCP 客户端，自动重连重试
│   ├── decode_registers()                             支持 float32/float64/uint32/int16/uint16/string
│   ├── load_register_table()                          按设备型号加载寄存器表 CSV
│   └── quantity_class()                               按参数名归类（电压/电流/功率/频率/其他）
│
├── 报告层
│   └── write_xlsx()                                   生成 汇总 / 对比明细 / 用例Mapping 三个 Sheet
│
├── 页面交互层（Playwright）
│   ├── PhysicalDevices（class）                       Physical Devices 列表页
│   │   └── scrape_connection()                        抓取设备真实 IP / Port / ModbusID
│   ├── PassThroughPage（class）                       Pass Through 配置页
│   │   ├── rows() / set_slaveid() / set_enabled()     行操作
│   │   ├── ensure_enabled()                           启用并等待透传服务就绪
│   │   └── save_and_result()                          保存并解析结果
│   └── _login() / page_factory fixture               登录 + 浏览器上下文管理
│
├── 数据流 Fixtures（pytest session-scoped）
│   ├── baseline                                       运行前快照（SlaveID + Enable 状态）→ 运行后还原
│   ├── configured                                     启用 Pass Through + 读取 SlaveID/型号
│   └── pt_comparison                                  A/B 两路读取 + 数值比对 + 写报告
│
└── 用例层（9 条）
    ├── test_pt_001/002/003                             业务数据正确性（功能主路径）
    └── test_pt_case01/02/03/04/06/07                  功能/边界/异常/并发用例
```

### 2.3 关键设计决策

| 设计点 | 说明 |
|---|---|
| 设备状态保护 | `baseline` fixture 快照 Pass Through 各行 SlaveID/勾选 + Pass Through/Device Mirror Enable 状态，运行后整体还原 |
| 透传服务等待 | `ensure_enabled()` 启用后轮询探针寄存器，确认透传服务真正就绪才继续，避免因未就绪导致误判 |
| 并发用例排序 | case07 排在改写 SlaveID 的 case02/03/04 之前，保证多台设备仍为干净基线值 |
| 寄存器表缺失处理 | 缺表型号自动 skip 并写入 `missing_tables.json`，不影响有寄存器表的设备继续测试 |

---

## 3. 配置参数（可用环境变量覆盖）

| 参数 | 来源 / 环境变量 | 默认 | 说明 |
|---|---|---|---|
| 网关地址 | `config.GATEWAY_IP` | 192.168.2.8 | A 路透传网关 Modbus TCP 地址 |
| 网页地址 | `config.HMI_URL` | https://192.168.2.9 | AcuHMI 登录地址 |
| 登录账号 | `config.HMI_USERNAME/PASSWORD` 或 `DM_USERNAME/DM_PASSWORD` | admin / Admin@110001 | 网页登录 |
| 比对容差 | `config.TOLERANCE_PERCENT / TOLERANCE_ABSOLUTE` | 1% / 0.05 | 相对/绝对容差 |
| 一致率阈值 | `COMPARE_MIN_RATE` | 0.6 | 数值量整体一致率下限 |
| 透传就绪等待 | `PT_SETTLE_MS` | 3000 | 启用后等网关透传服务起来（毫秒） |
| 禁用停服等待 | `PT_DISABLE_SETTLE` | 8 | 禁用后确认停服的轮询秒数 |
| 并发时长 | `CONCURRENT_SECONDS` | 20 | case07 多主站并发读秒数 |
| 有头模式 | `HEADED` | 0 | =1 显示浏览器操作过程 |

---

## 4. 用例清单（9 条）

| 用例ID | 用例函数名 | 测试类型 | 验证点 | 关联TC |
|---|---|---|---|---|
| PT-001 | test_pt_001_config | 功能 | 启用Pass Through，Ethernet设备SlaveID在101–247 | （前置检查，无对应手工用例） |
| PT-002 | test_pt_002_data_collected | 功能 | 按型号寄存器表透传读到可比对数据 | （前置检查，无对应手工用例） |
| PT-003 | test_pt_003_passthrough_matches_direct | 数据正确性 | 透传(A) ↔ 直连电表(B) 逐寄存器一致率达标 | TestCase_AcuHMI_008_03_case05 |
| PT-004 | test_pt_case01_enable_disable_toggle | 功能 | Disable后透传读取失败；Enable后恢复 | TestCase_AcuHMI_008_03_case01 |
| PT-005 | test_pt_case07_concurrent_masters | 并发 | 多主站并发读不同SlaveID无串扰 | TestCase_AcuHMI_008_03_case07 |
| PT-006 | test_pt_case02_valid_slaveid_save | 边界 | 有效SlaveID(101–247)保存成功且持久化 | TestCase_AcuHMI_008_03_case02 |
| PT-007 | test_pt_case03_slaveid_boundary | 边界 | 越界(100/248)被拒；边界(101/247)通过 | TestCase_AcuHMI_008_03_case03 |
| PT-008 | test_pt_case04_duplicate_slaveid | 异常 | 重复SlaveID被拒或不同时生效 | TestCase_AcuHMI_008_03_case04 |
| PT-009 | test_pt_case06_disabled_blocks_access | 功能 | 禁用后无法经透传SlaveID访问下游 | TestCase_AcuHMI_008_03_case06 |

> 用例Mapping 同步写入 `pass_through_compare.xlsx` 的「用例Mapping」Sheet。
> PT-001/PT-002 为前置检查（配置生效、透传取到数），无对应手工用例编号；HTML 报告中通过/跳过时不展示这两条，失败时仍会展示以便排查。

---

## 5. 设备状态保护

模块级 `baseline` fixture 快照 Pass Through 各行 SlaveID/勾选 + Pass Through / Device Mirror 两个 Enable 开关，
全部用例后整体还原。运行不会改变设备最终配置。

---

## 6. 工程路径

```
C:\JrJ\auto\autotest\Protocols\Pass Through\
├── test_pass_through_business.py   # 主脚本（pytest 入口）
└── README.md                       # 本文件

依赖配置：C:\JrJ\auto\autotest\Protocols\config.py
寄存器模板：C:\JrJ\auto\autotest\Protocols\template\AcuCloud 模板适配\<Model>.xlsx
测试报告：C:\JrJ\auto\autotest\Protocols\reports\pass_through_compare.xlsx
```
