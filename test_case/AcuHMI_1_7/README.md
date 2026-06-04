# AcuHMI-1-7 自动化测试

AcuHMI-1-7 网关测试脚本目录。

算法规格：`knowledge/gateway/web2/requirements/summaries/wiring_check_v1.05.md`（与 WEB2 共用同一套算法规格）

---

## 目录结构

```
AcuHMI_1_7/
├── bacnet_ui/                 BACnet/IP UI 自动化测试
│   ├── conftest.py            session 级 browser/context fixture，module 级 hmi_page fixture
│   ├── test_bacnet_ui_basic.py  P0 参数列表一致性用例（4 条）
│   ├── test_bacnet_ui_config.py P1/P2 配置页面功能用例（21 条）
│   ├── test_bacnet_ui_protocol.py P1/P2 协议端到端验证用例（11 条，含 BACnet 客户端）
│   ├── helpers/
│   │   ├── template_matcher.py  读取北向模板参数（BACnet description 集合）
│   │   └── hmi_bacnet_client.py BACnet/IP 客户端同步接口（BAC0 封装）
│   └── screenshots/           页面探查截图
└── wiring_check/              接线检查（Wiring Check）自动化测试
    ├── core/                  基础设施模块
    │   ├── config.py          寄存器地址 + 连接参数 + User Channel 配置
    │   ├── meter_modbus.py    Modbus TCP 读写（4100 电表）
    │   ├── signal_driver.py   控源封装（set_ac）
    │   ├── expected_engine.py 5种接线方式算法引擎（推导预期 Wiring Status）
    │   ├── wiring_check_page.py Playwright 页面对象（登录/导航/触发/解析）
    │   └── report.py          HTML 报告生成
    ├── conftest.py            pytest session/module fixtures
    ├── test_3e4wy.py          3 Element 4 Wire Y（31条）
    ├── test_2e3w_delta.py     2 Element 3 Wire Delta（21条）
    ├── test_2e3w_network.py   2 Element 3 Wire Network（31条）
    ├── test_2e3w_1phase.py    2 Element 3 Wire 1 Phase（12条）
    ├── test_1e2w.py           1 Element 2 Wire（4条）
    ├── debug_page.py          页面结构调试脚本
    └── reports/               HTML 报告输出（自动创建）
```

---

---

## BACnet/IP UI 测试

BACnet/IP 模块共 45 条用例，已实现自动化 36 条（80%），不实现 9 条（COV 客户端验证 8 条 + 物理接线依赖 1 条）。

### 用例一览

**test_bacnet_ui_basic.py**（P0 参数列表一致性，4 条）

| 函数 | 用例编号 | 说明 |
|---|---|---|
| test_019_acurev2100_param_list_matches_template | TestCase_AcuHMI-1-7_033_001_019 | AcuRev-2100 参数列表与模板一致（设备未接入自动 skip） |
| test_020_acurev4100_param_list_matches_template | TestCase_AcuHMI-1-7_033_001_020 | AcuRev-4100 参数列表与北向模板一致 |
| test_034_acurev2100_cov_batch_update_matches_template | TestCase_AcuHMI-1-7_033_001_034 | COV Batch Update AcuRev-2100 参数列表与模板一致 |
| test_035_acurev4100_cov_batch_update_matches_template | TestCase_AcuHMI-1-7_033_001_035 | COV Batch Update AcuRev-4100 参数列表与模板一致 |

**test_bacnet_ui_config.py**（P1/P2 配置页面功能，21 条）

| 函数 | 用例编号 | 级别 | 说明 |
|---|---|---|---|
| test_001_page_load_and_default_state | TestCase_AcuHMI-1-7_033_001_001 | LV1 | 页面入口与默认状态 |
| test_004_device_object_name_save | TestCase_AcuHMI-1-7_033_001_004 | LV1 | Device Object Name 配置保存并持久化 |
| test_005_device_instance_save | TestCase_AcuHMI-1-7_033_001_005 | LV1 | Device Instance 配置保存并持久化 |
| test_006_apdu_timeout_valid_boundary | TestCase_AcuHMI-1-7_033_001_006 | LV1 | APDU Timeout 边界选项（3 s / 60 s）保存正常 |
| test_007_apdu_retries_valid_boundary | TestCase_AcuHMI-1-7_033_001_007 | LV1 | APDU Retries 边界选项（0 / 10）保存正常 |
| test_008_foreign_device_toggle | TestCase_AcuHMI-1-7_033_001_008 | LV1 | Enable Foreign Device 开关联动（BBMD 字段显隐） |
| test_009_time_to_live_valid_boundary | TestCase_AcuHMI-1-7_033_001_009 | LV1 | Time To Live 合法边界值（5 / 1440）保存正常 |
| test_021_epics_and_cov_linkage | TestCase_AcuHMI-1-7_033_001_021 | LV1 | COV Enable 与 COV Increment 联动 |
| test_022_cov_increment_valid_range | TestCase_AcuHMI-1-7_033_001_022 | LV1 | COV Increment 合法值（0.000 / 0.123）保存正常 |
| test_032_cov_batch_update_2100_save | TestCase_AcuHMI-1-7_033_001_032 | LV1 | 先使能 Polling Enable，COV Batch Update Select All 批量配置 AcuRev2100 参数并验证持久化 |
| test_033_cov_batch_update_4100_save | TestCase_AcuHMI-1-7_033_001_033 | LV1 | 先使能 Polling Enable，COV Batch Update Select All 批量配置 AcuRev-4100 参数并验证持久化 |
| test_036_batch_update_partial_coverage | TestCase_AcuHMI-1-7_033_001_036 | LV1 | Batch Update 仅覆盖参数 B，验证参数 A 原 COV Increment 不变 |
| test_029_empty_cov_increment_blocks_param_type_switch | TestCase_AcuHMI-1-7_033_001_029 | LV2 | COV Increment 为空时切换 Parameter Type 被阻止 |
| test_030_empty_cov_increment_blocks_page_switch | TestCase_AcuHMI-1-7_033_001_030 | LV2 | COV Increment 为空时切换分页被阻止 |
| test_040_bacnet_port_invalid_values | TestCase_AcuHMI-1-7_033_001_040 | LV2 | Port 非法值（47807 / 49001）被拦截 |
| test_041_network_number_invalid_values | TestCase_AcuHMI-1-7_033_001_041 | LV2 | Network Number 非法值（0 / 65535）被拦截 |
| test_045_cov_increment_invalid_value | TestCase_AcuHMI-1-7_033_001_045 | LV2 | COV Increment 负值（-0.001）被拦截 |
| test_042_apdu_timeout_options | TestCase_AcuHMI-1-7_033_001_042 | LV2 | Advertised APDU Timeout 下拉选项集合校验（3/6/10/20/30/45/60 s） |
| test_043_apdu_retries_options | TestCase_AcuHMI-1-7_033_001_043 | LV2 | Advertised APDU Retries 下拉选项集合校验（0/1/2/3/5/10） |
| test_044_time_to_live_invalid_values | TestCase_AcuHMI-1-7_033_001_044 | LV2 | Time To Live 非法边界外值（4 / 1441）被拦截 |
| test_011_invalid_bbmd_config_rejected | TestCase_AcuHMI-1-7_033_001_011 | LV2 | 非法 BBMD 配置（IP/Port/TTL 同时非法）保存被阻止 |

**test_bacnet_ui_protocol.py**（P1/P2 协议端到端验证，11 条，需 BACnet 客户端）

| 函数 | 用例编号 | 级别 | 说明 |
|---|---|---|---|
| test_002_bacnet_port_valid_boundary | TestCase_AcuHMI-1-7_033_001_002 | LV1 | Port 合法边界值（47808/49000）保存后 BACnet 客户端可通信 |
| test_003_network_number_valid_boundary | TestCase_AcuHMI-1-7_033_001_003 | LV1 | Network Number 合法边界值（1/65534）保存并持久化 |
| test_012_acurev4100_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_012 | LV1 | AcuRev4100 EPICS Enable 后 BACnet 客户端可接收参数（需设备接入） |
| test_013_pxe1_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_013 | LV2 | PXE1 EPICS Enable 后 BACnet 客户端可接收参数（需设备接入） |
| test_014_pxe2_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_014 | LV2 | PXE2 EPICS Enable 后 BACnet 客户端可接收参数（需设备接入） |
| test_015_pxm350_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_015 | LV2 | PXM350 EPICS Enable 后 BACnet 客户端可接收参数（需设备接入） |
| test_016_acuvim3_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_016 | LV2 | AcuVIM3 EPICS Enable 后 BACnet 客户端可接收参数（需设备接入） |
| test_017_acurev2100_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_017 | LV1 | AcuRev-2100 EPICS Enable 后 BACnet 客户端可接收参数（需设备接入） |
| test_018_epics_disable_bacnet_not_readable | TestCase_AcuHMI-1-7_033_001_018 | LV1 | 关闭 EPICS Enable 后 BACnet 对象数量减少（需设备接入） |
| test_037_epics_file_download | TestCase_AcuHMI-1-7_033_001_037 | LV2 | EPICS File Download 触发浏览器下载，文件名非空 |
| test_038_bacnet_disable_client_unreachable | TestCase_AcuHMI-1-7_033_001_038 | LV1 | BACnet Enable = Disable 后客户端无法连接，恢复后可连 |

### 不实现自动化用例（共 9 条）

**物理接线 / 硬件依赖（1 条）：**

| 用例编号 | 说明 | 原因 |
|---|---|---|
| TestCase_AcuHMI-1-7_033_001_010 | 配置 BBMD 后支持跨路由器访问设备 | 需要物理接线（独立子网 + BBMD 转发设备），台架不具备 |

**COV 客户端验证（8 条）：**

| 用例编号 | 说明 |
|---|---|
| TestCase_AcuHMI-1-7_033_001_023 | AcuRev4100 COV Enable 后 BACnet 客户端接收 COV 警告 |
| TestCase_AcuHMI-1-7_033_001_024 | AcuRev2100 COV Enable 后 BACnet 客户端接收 COV 警告 |
| TestCase_AcuHMI-1-7_033_001_025 | PXE1 COV Enable 后 BACnet 客户端接收 COV 警告 |
| TestCase_AcuHMI-1-7_033_001_026 | PXE2 COV Enable 后 BACnet 客户端接收 COV 警告 |
| TestCase_AcuHMI-1-7_033_001_027 | AcuVIM3 COV Enable 后 BACnet 客户端接收 COV 警告 |
| TestCase_AcuHMI-1-7_033_001_028 | PXM350 COV Enable 后 BACnet 客户端接收 COV 警告 |
| TestCase_AcuHMI-1-7_033_001_031 | 关闭 COV Enable 后 BACnet 客户端不再接收 COV 警告 |
| TestCase_AcuHMI-1-7_033_001_039 | 设备离线时配置 COV，客户端超时等待被阻止 |

### BACnet/IP 执行命令（仓库根目录）

`pytest.ini` 已配置 `addopts = --html=pytest_report.html --self-contained-html`，所有命令自动生成报告，无需手动加 `--html` 参数。

```bash
# 全部 BACnet UI 用例（basic 4 + config 21 + protocol 11 = 36 条）
pytest test_case/AcuHMI_1_7/bacnet_ui/ -v

# 只跑 P0 参数列表一致性（纯 UI，无 BACnet 客户端）
pytest test_case/AcuHMI_1_7/bacnet_ui/test_bacnet_ui_basic.py -v

# 只跑 P1/P2 配置页面功能（纯 UI，无 BACnet 客户端）
pytest test_case/AcuHMI_1_7/bacnet_ui/test_bacnet_ui_config.py -v

# 只跑协议端到端验证（需 BACnet 客户端连接，须串行执行勿加 -n）
pytest test_case/AcuHMI_1_7/bacnet_ui/test_bacnet_ui_protocol.py -v

# 单条用例（-k 关键词过滤）
pytest test_case/AcuHMI_1_7/bacnet_ui/ -k "test_040" -v

# 同时跑多条用例（用 or 连接关键词）
pytest test_case/AcuHMI_1_7/bacnet_ui/ -k "test_033 or test_036" -v

# 跳过需要 BACnet 客户端的协议用例（纯 UI 回归）
pytest test_case/AcuHMI_1_7/bacnet_ui/ --ignore=test_case/AcuHMI_1_7/bacnet_ui/test_bacnet_ui_protocol.py -v
```

报告输出到仓库根目录 `pytest_report.html`。

### BACnet/IP 前置条件

- HMI 设备已上电，地址/账号写在 `bacnet_ui/conftest.py` 顶部（`HMI_URL`、`HMI_USERNAME`、`HMI_PASSWORD`）
- 依赖：`pip install pytest playwright` + `playwright install chromium`
- 协议端到端用例（`test_bacnet_ui_protocol.py`）额外依赖 BAC0：`pip install BAC0`
  - BACnet 客户端配置在 `helpers/hmi_bacnet_client.py`（`LOCAL_IP`、`LOCAL_PORT`、`HMI_IP`）
  - `test_002`/`test_038` 会变更全局 BACnet 状态，**protocol.py 必须串行运行**（不加 `-n`）
- `test_021`/`test_022`/`test_032`/`test_033`/`test_036`/`test_045` 及 protocol.py 中 012-018 依赖设备接入，未接入时自动 skip

#### COV Batch Update 用例前置逻辑（test_034 / test_035 / test_033 / test_036）

> **COV Batch Update 的 Parameters 下拉只显示已开启 Polling Enable 的参数。**
> 若某参数 Polling Enable 未开启，该参数不会出现在 Batch Update 下拉中，无法被批量配置。

`_collect_cov_batch_params`（test_034/035 用）和相关 COV 配置用例（test_033/036）的执行逻辑必须包含以下步骤：

1. 打开设备 Parameter Config 弹窗
2. **对当前 Parameter Type 的所有可见行，翻页批量开启 Polling Enable**
3. 开启后打开 COV Batch Update，此时 Parameters 下拉才能显示该 Type 的全量参数
4. 读取 / 操作参数后关闭 Batch Update
5. 切换到下一个 Parameter Type，重复步骤 2-4
6. 所有 Type 遍历完后关闭 Parameter Config

---

## 与 WEB2 脚本的主要差异

| 差异点 | WEB2 | HMI1-7 |
|-------|------|--------|
| 网关 IP | `192.168.3.9` | `192.168.2.8` |
| 页面导航 | 直接 goto URL | Settings 侧边栏 → Diagnostics → Wiring Check tab |
| Confirm 弹窗 | 有 Save 按钮（下发额定电压到表） | **无 Save 按钮**（额定电压只用于计算） |
| 设备选择下拉 | 有（WEB2 下挂多台 4100） | 有（默认 All，可选单台） |
| 结果表结构 | — | **与 WEB2 完全相同** |
| 算法规格 | — | **与 WEB2 完全相同** |

---

## 环境依赖

```bash
pip install playwright openpyxl pymodbus pytest-html
playwright install chromium
```

---

## 运行前配置

**所有可调参数统一在 `test_case/AcuHMI_1_7/config.py` 中修改，各子模块自动从此文件读取，不再需要分别修改。**

| 配置项 | 说明 | 默认值 |
|-------|------|-------|
| `HMI_IP` | HMI 设备 IP | `192.168.2.8` |
| `HMI_USERNAME / HMI_PASSWORD` | Web UI 登录账号 | `q / 1` |
| `HMI_DEVICE_NAME` | HMI 页面 Device 下拉中的被测设备名，留空则选 All | `Acu4100` |
| `BACNET_PORT` | HMI BACnet/IP 服务端口 | `47808` |
| `LOCAL_IP / LOCAL_PORT` | 本机 BAC0 客户端监听地址/端口 | `192.168.2.45 / 47810` |
| `BACNET_CONNECT_WAIT` | BAC0 启动后等待网关响应（s） | `4.0` |
| `BACNET_READ_TIMEOUT` | 单次 BACnet 属性读取超时（s） | `20.0` |
| `BACNET_RESTART_WAIT` | UI 保存后等待 BACnet 服务重启（s） | `6.0` |
| `METER_TCP_IP / METER_TCP_PORT` | AcuRev4100 Modbus TCP 地址 | `192.168.2.29 / 502` |
| `MODBUS_SLAVE` | AcuRev4100 Slave ID | `102` |
| `NOMINAL_VOLTAGE` | 额定电压（V） | `100` |
| `NORMAL_CURRENT` | 电流侧正常幅值（A） | `1.0` |
| `CHECK_TIMEOUT` | 等待接线检查完成的超时（s） | `30` |

---

## 运行方式

### 方式一：pytest + pytest-html（推荐）

```bash
# 全部 99 条用例
pytest test_case/AcuHMI_1_7/wiring_check/ -v \
  --html=test_case/AcuHMI_1_7/wiring_check/reports/pytest_report.html \
  --self-contained-html

# 单种接线方式
pytest test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py -v
pytest test_case/AcuHMI_1_7/wiring_check/test_2e3w_delta.py -v
pytest test_case/AcuHMI_1_7/wiring_check/test_2e3w_network.py -v
pytest test_case/AcuHMI_1_7/wiring_check/test_2e3w_1phase.py -v
pytest test_case/AcuHMI_1_7/wiring_check/test_1e2w.py -v

# 单条用例
pytest test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py -k "V-13-ABC" -v

# 按类型过滤
pytest test_case/AcuHMI_1_7/wiring_check/ -k "V-"    # 只跑电压
pytest test_case/AcuHMI_1_7/wiring_check/ -k "I-"    # 只跑电流
pytest test_case/AcuHMI_1_7/wiring_check/ -k "PASS"  # 只跑基准

# 失败时立即停止
pytest test_case/AcuHMI_1_7/wiring_check/ -x
```

### 方式二：直接运行（生成自定义 HTML 报告，含每通道明细）

```bash
# 全部接线方式一键运行
python test_case/AcuHMI_1_7/wiring_check/run_all.py

# 单种接线方式
python test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py
python test_case/AcuHMI_1_7/wiring_check/test_2e3w_delta.py
python test_case/AcuHMI_1_7/wiring_check/test_2e3w_network.py
python test_case/AcuHMI_1_7/wiring_check/test_2e3w_1phase.py
python test_case/AcuHMI_1_7/wiring_check/test_1e2w.py

# 指定单条或多条用例（传 ID 参数）
python test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py V-13-ABC
python test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py I-04-ABC I-07-ABC
```

报告路径：`wiring_check/reports/wiring_<接线方式>_<时间戳>.html`

---

## 用例数量

| 接线方式 | 合计 | 说明 |
|---------|------|------|
| 3E4WY（ABC+ACB） | 34 | +3 guard 负向用例（条件2/3/4 guard）；+4 ACB 电流用例 |
| 2E3W Delta（ABC+ACB） | 23 | +4 ACB 电流用例；+2 防回归用例（V-05/V-06 REG） |
| 2E3W Network（ABC+ACB） | 34 | 同 3E4WY（共用测试用例列表） |
| 2E3W 1Phase | 12 | |
| 1E2W | 4 | |
| **合计** | **107** | |

---

## 测试逻辑

```
① Modbus 写接线方式（0x1042）+ 相序（0x10DC）+ 额定电压（0x6501/0x651A）至 4100
② 控源 set_ac() 输出信号（等 8s 稳定）
③ expected_engine 根据控源参数 + 算法规格推导预期 Wiring Status
④ Playwright：Settings → Diagnostics → Wiring Check tab → 触发检查 → 等完成 → 解析电压表 + 电流表
⑤ 比对预期 vs 实测 → PASS / FAIL，生成 HTML 报告
```

### HMI 页面导航说明

`wiring_check_page.navigate()` 每次通过菜单三步导航（直接 goto URL 会被 SPA 路由重定向）：

1. `goto /#/systemSettings/dateTime` — 激活 Settings 侧边栏
2. 点击侧边栏 **Diagnostics**（精确文字匹配）
3. 点击页内 **Wiring Check** tab

整个测试会话只导航一次（session scope），后续用例复用同一页面实例。

---

## 关键配置

| 配置项 | 位置 | 当前值 |
|-------|------|-------|
| 被测设备名（页面下拉） | `core/config.py` `HMI_DEVICE_NAME` | `''`（需填写） |
| Meter IP / Port / Slave | `core/config.py` | `192.168.2.29 : 502 : 102` |
| HMI 地址 | `core/config.py` `HMI_IP` | `192.168.2.8` |
| HMI 登录账号 | `core/config.py` `HMI_USER / HMI_PASS` | `q / 1` |
| 额定电压 | `core/config.py` `NOMINAL_VOLTAGE` | `100 V` |
| 信号稳定等待 | `comm/source_control.py` `set_ac()` | `8 s` |

---

## 台架接线要求

**Delta / 1Phase 接线方式须按以下方式连接，否则多通道会得到错误的电流角度：**

| 电流源 | 连接目标 |
|--------|---------|
| Source A | 所有 User Channel 的 Phase A CT |
| Source C | 所有 User Channel 的 Phase C CT |
| Source B | **不接** 任何 Delta / 1Phase CT |

> 3E4WY / Network 正常接三相即可，无特殊要求。

---

## 相关知识库

- 算法规格摘要（v1.05）：`knowledge/gateway/web2/requirements/summaries/wiring_check_v1.05.md`
- 项目 context：`knowledge/gateway/hmi1-7/context.md`
