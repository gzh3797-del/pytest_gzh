# BACnet/IP UI 测试模块（`projects/acuhmi_1_7/tests/bacnet/`）

AcuHMI-1-7 BACnet/IP 配置页面功能 + 协议端到端验证。共 **45 条**用例，已自动化 **36 条**（basic 4 + config 21 + protocol 11），不实现 9 条（COV 客户端验证 8 + 物理接线依赖 1）。

> **报告**：统一用 pytest-html。单跑本模块时加 `--html=<文件> --self-contained-html` 出单文件报告；或用 `python run.py acuhmi_1_7`（跑整项目，报告自动落 `reports/acuhmi_1_7/<时间戳>/html/`）。

---

## 环境依赖

```bash
pip install pytest pytest-html playwright openpyxl pymodbus
playwright install chromium

# 协议端到端用例（test_bacnet_ui_protocol.py）需要 BACnet 客户端
pip install BAC0
```

---

## 执行命令（仓库根目录）

> 报告：加 `--html=reports/bacnet_ui.html --self-contained-html` 出单文件 pytest-html 报告，跑完用浏览器打开（单文件，可直接双击）。

```bash
# 全部 BACnet UI 用例（basic 4 + config 21 + protocol 11 = 36 条）
pytest projects/AcuHMI_1_7/tests/bacnet/ -v --html=reports/bacnet_ui.html --self-contained-html

# 只跑 P0 参数列表一致性（纯 UI，无 BACnet 客户端）
pytest projects/AcuHMI_1_7/tests/bacnet/test_bacnet_ui_basic.py -v --html=reports/bacnet_ui.html --self-contained-html

# 只跑 P1/P2 配置页面功能（纯 UI，无 BACnet 客户端）
pytest projects/AcuHMI_1_7/tests/bacnet/test_bacnet_ui_config.py -v --html=reports/bacnet_ui.html --self-contained-html

# 只跑协议端到端验证（需 BACnet 客户端连接，须串行执行勿加 -n）
pytest projects/AcuHMI_1_7/tests/bacnet/test_bacnet_ui_protocol.py -v --html=reports/bacnet_ui.html --self-contained-html

# 单条用例（-k 关键词过滤）
pytest projects/AcuHMI_1_7/tests/bacnet/ -k "test_013" -v --html=reports/bacnet_ui.html --self-contained-html

# 跳过需要 BACnet 客户端的协议用例（纯 UI 回归）
pytest projects/AcuHMI_1_7/tests/bacnet/ --ignore=projects/AcuHMI_1_7/tests/bacnet/test_bacnet_ui_protocol.py -v --html=reports/bacnet_ui.html --self-contained-html
```

> 查看报告：浏览器打开 `reports/bacnet_ui.html`（`reports/` 已 git 忽略；`--self-contained-html` 为单文件，可直接双击）。

---

## 前置条件

- HMI 设备已上电；地址统一在 **`projects/acuhmi_1_7/config.yaml`**（`hmi_url` / `hmi_ip`），登录账号密码（`WEB_USERNAME` / `WEB_PASSWORD`）放 **`configs/.env`**（模板见 `configs/.env.example`），经 `projects/acuhmi_1_7/settings.py` 适配层暴露（`DEFAULT_USERNAME` / `DEFAULT_PASSWORD`），登录复用统一的 `LoginPage`。
- 依赖：`pip install pytest pytest-html playwright openpyxl pymodbus` + `playwright install chromium`（见「环境依赖」）。
- 协议端到端用例（`test_bacnet_ui_protocol.py`）额外依赖 BAC0：`pip install BAC0`
  - BACnet 客户端参数同样在 `config.yaml`（`local_ip` / `local_port` / `bacnet_port`）。
  - 数值比对（012–017）需该设备 Modbus 连接（`config.yaml` 的 `meter_tcp_ip` / `meter_tcp_port` / `modbus_slave` 等）；未配置时自动跳过数值比对，仅保留范围+可读验证。
  - `test_002` / `test_038` 会变更全局 BACnet 状态，**protocol.py 必须串行运行**（不加 `-n`）。
- `test_021` / `test_022` / `test_032` / `test_033` / `test_036` 及 protocol.py 中 012–018 依赖设备接入，未接入时自动 skip。

> **运行时长提示**：012–017 会启用设备**全部** BACnet 参数并逐一读回（AcuRev4100 约 1869 个、AcuRev2100 约 1225 个），单条用例可能需数分钟。想更快可用 `-k` 只跑小参数量设备（VIM3 ≈154 / AcuRev1300 ≈39）。

---

## 用例一览

**`test_bacnet_ui_basic.py`**（P0 参数列表一致性，4 条）

| 函数 | 用例编号 | 说明 |
|---|---|---|
| test_019_acurev2100_param_list_matches_template | TestCase_AcuHMI-1-7_033_001_019 | AcuRev-2100 参数列表与模板一致（设备未接入自动 skip） |
| test_020_acurev4100_param_list_matches_template | TestCase_AcuHMI-1-7_033_001_020 | AcuRev-4100 参数列表与北向模板一致 |
| test_034_acurev2100_cov_batch_update_matches_template | TestCase_AcuHMI-1-7_033_001_034 | COV Batch Update AcuRev-2100 参数列表与模板一致 |
| test_035_acurev4100_cov_batch_update_matches_template | TestCase_AcuHMI-1-7_033_001_035 | COV Batch Update AcuRev-4100 参数列表与模板一致 |

**`test_bacnet_ui_config.py`**（P1/P2 配置页面功能，21 条）

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

**`test_bacnet_ui_protocol.py`**（P1/P2 协议端到端验证，11 条，需 BACnet 客户端）

| 函数 | 用例编号 | 级别 | 说明 |
|---|---|---|---|
| test_002_bacnet_port_valid_boundary | TestCase_AcuHMI-1-7_033_001_002 | LV1 | Port 合法边界值（47808/49000）保存后 BACnet 客户端可通信 |
| test_003_network_number_valid_boundary | TestCase_AcuHMI-1-7_033_001_003 | LV1 | Network Number 合法边界值（1/65534）保存并持久化 |
| test_012_acurev4100_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_012 | LV1 | AcuRev4100 开启全部参数 Polling Enable 后客户端读回与模板严格一致（全量，需设备接入） |
| test_013_pxe1_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_013 | LV2 | PXE1（AcuvimIIR）全量参数开启后客户端读回与模板严格一致（需设备接入） |
| test_014_pxe2_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_014 | LV2 | PXE2（AcuvimIIW）全量参数开启后客户端读回与模板严格一致（需设备接入） |
| test_015_pxm350_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_015 | LV2 | PXM350（AcuRev1300）全量参数开启后客户端读回与模板严格一致（需设备接入） |
| test_016_acuvim3_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_016 | LV2 | AcuVIM3 全量参数开启后客户端读回与模板严格一致（需设备接入） |
| test_017_acurev2100_epics_enable_bacnet_readable | TestCase_AcuHMI-1-7_033_001_017 | LV1 | AcuRev-2100 全量参数开启后客户端读回与模板严格一致（需设备接入） |
| test_018_epics_disable_bacnet_not_readable | TestCase_AcuHMI-1-7_033_001_018 | LV1 | 关闭 EPICS Enable 后 BACnet 对象数量减少（需设备接入） |
| test_037_epics_file_download | TestCase_AcuHMI-1-7_033_001_037 | LV2 | EPICS File Download 触发浏览器下载，文件名非空 |
| test_038_bacnet_disable_client_unreachable | TestCase_AcuHMI-1-7_033_001_038 | LV1 | BACnet Enable = Disable 后客户端无法连接，恢复后可连 |

## 不实现自动化用例（共 9 条）

**物理接线 / 硬件依赖（1 条）：**

| 用例编号 | 说明 | 原因 |
|---|---|---|
| TestCase_AcuHMI-1-7_033_001_010 | 配置 BBMD 后支持跨路由器访问设备 | 需物理接线（独立子网 + BBMD 转发设备），台架不具备 |

**COV 客户端验证（8 条）：** `..._023` ~ `..._028`、`..._031`、`..._039`（COV Enable/Disable 后客户端收/停收 COV 警告、离线超时等，需 COV 客户端验证，暂不自动化）。

---

## 报告内容（pytest-html）

每条用例在报告中显示 通过/失败/跳过 + 耗时；用例编号写入 JUnit XML property（`record_property`）。协议用例 012–017 的详细信息（上传参数明细、范围结论 缺失=0/多余=0、BACnet `presentValue` vs 直连 Modbus 实时值 ±1%/±0.05 的数值比对结果）通过 `logging` 输出，在 pytest-html 报告的每条用例 **Captured log** 区可见；失败用例附 assert 详情（缺失/多余参数、未读到值对象、超差参数列表）。
