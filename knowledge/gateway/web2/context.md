# web2 — 网关自动化测试项目

## 项目背景
测试 AcuRev-4100-WEB2 网关模块（型号 ACM-41-WEB2）的数据采集、协议转换与云端上传功能。
当前测试阶段为 **Sprint 2（二期）**，在一期基础上新增多项协议与功能，详见下方需求节。

## 被测产品需求

> Sprint 1 汇总 → requirements/summaries/sprint1_requirements.md（V1.0，2026-02-25）
> Sprint 2 汇总 → requirements/summaries/sprint2_requirements.md（V1.01，2026-04-23）+《变更说明书 v1.00》（2026-04-29）
> 接线检查变更 → requirements/summaries/wiring_check_v1.00_20260519.md（v1.00，2026-05-19）
> 接线检查算法规格 → requirements/summaries/wiring_check_v1.05.md（v1.05，原件 requirements/raw/接线检测总表_ver1.05.xlsx）

### 产品定位
WEB2 作为 AcuRev-4100 的网络模块，附加在表本体上直接取电。通过高速 USB（Modbus RTU）与表本体通信，通过以太网（Modbus TCP）最多下挂 3 台非本体 4100 和 8 台 AcuIOM；向上提供 BACnet/IP、Modbus TCP、MQTT、SNMP、Ethernet/IP 等协议接口，并支持 AcuCloud、AWS IoT、Azure IoT 数据上传。

### 网络拓扑（image3 原图）
WEB2 顶层 → Ethernet 2 → 3 台 AcuRev-4100（各自 Eth1 in / Eth2 out）→ 逐级下挂共 8 台 AcuIOM

### 新增/变更功能清单

#### 设备管理
| 功能 | 要点 |
|------|------|
| Virtual Device | 多台 4100 有功能量按公式计算映射；SN 格式 AEVM{5位随机数}；参数≥20个；支持进/出线有功能量（变更） |
| 设备固件更新 | 支持对下挂 4100 和 AcuIOM 升级（MFEA 格式）；禁止固件降级 |
| 模板版本兼容 | 同一网关允许同型号不同固件版本设备共存；模板命名含固件版本号 |
| Checkpoint | Data Log 页面可选历史备份时间点查看 |
| 设备自发现 | mDNS 扫描（**标注暂不做**） |

#### 电表配置
| 功能 | 要点 |
|------|------|
| PT 配置 | PT1/PT2（受铅封保护）；Nominal Current 1~50,000A；修改立即生效，不回算历史 |
| Meter Point | "User Channel"全局改名为"Meter Point"；支持自定义名称（ASCII ≤20字节） |

#### 协议扩展
| 协议 | 状态 | 关键点 |
|------|------|--------|
| BACnet/IP | **新增** | 默认端口47808；每参数可配 COV；支持 EPICS 下载；映射范围含 4100+AcuIOM；**变更：新增谐波参数支持** |
| Ethernet/IP | **新增** | EDS 文件下载；product code 10003；主要对接罗克韦尔 PLC；显式报文端口 44818（44800~44899）；隐式报文 UDP 2222（固定不可改，绑定 Eth1） |
| AWS IoT | **新增** | 断网缓存 24~72h |
| Azure IoT | **新增** | 支持 Device Twin 远程配置 WEB2 |
| Device Mirror | **新增** | 预定义 Slave ID 映射，用户不可配置 |
| Modbus 端口 | **变更** | 范围改为 2000~5999 |

#### Data Log 默认参数（变更文档新增）
- 4100 电表：Realtime（电压/电流/功率/频率）+ Energy（进/出线有功能量）默认选中
- AcuIOM：全部参数默认选中

#### 系统诊断
- **Wiring Check**：五种接线方式检查（WEB2 范围），算法规格见 v1.05 xlsx（该 xlsx 同时覆盖 HMI1-7 额外 4 种：2.5E4WY、3E3W Delta、3E4W Delta、2E2W Delta），电压侧/电流侧分开，支持导出 CSV

#### 安全
- 默认密码改为 Admin@AABBCC（AABBCC=SN后六位）
- 新增忘记密码（每日临时密码，仅 Admin 可用）
- 禁止固件降级；EULA 强制同意；默认禁用 Modbus Pass Through 和 SSH

#### AcuCloud
- Advanced 后门（`?showAdvanced=true`）可修改数据传输 URL、固件更新 URL、Remote Access URL
- Installation/Inspection Record 保存时自动推送至 AcuCloud（仅 4100 设备）

#### Web UI 改版
- About 拆分出 Service 页面（含 Troubleshooting、.a2d 加密诊断文件下载）
- AcuIOM AI/AO/DI 界面重构，新增 Copy/Apply/Reset 批量操作
- Post Channel 测试通道显示详细信息（参考 AXM-WEB2）；Timezone → Time Zone（变更）

## 当前 Bug 状态（2026-05-13，A4WS-135 ~ A4WS-197）

> 完整索引见 bugs/INDEX.md，原始 Jira 导出见 bugs/raw/JIRA.csv

**缺陷共 60 条，全部 Medium 优先级。未关闭 29 条（排除 REJECTED 2 条）。**

| 状态 | 数量 | 主要模块 |
|------|------|---------|
| CLOSED 已关闭 | 29 | BACnet/IP、AWS/Azure IoT 历史问题 |
| IN PROGRESS 修复中 | 10 | MQTT、固件升级、Troubleshooting、AWS IoT 等 |
| CREATED 未分配 | 10 | **接线检查（4条）**、AcuCloud、SNMP、BACnet/IP 性能、Modbus 默认端口 |
| TO BE VERIFIED 待验证 | 8 | 设备配置、BACnet COV、系统稳定性等 |

**重点：接线检查 4 种接线方式结果全错（A4WS-164/167/168/169），建议合并排查同一底层逻辑。**
**BACnet/IP 上传性能问题（A4WS-181）：1869 参数耗时 40 分钟，严重异常。**

## 测试脚本（Protocols/）

| 协议 | 对比方式 | 脚本路径 | 说明 |
|------|---------|---------|------|
| BACnet/IP | **实时对比** | `BACnetIP/comparator.py` | BACnet Present Value vs 实时 Modbus，含范围/单位/数值三段检查 |
| AcuCloud | 快照对比 | `AcuCloud/cloud_comparator.py` | AcuCloud xlsx 快照 vs 实时 Modbus |
| MQTT | 快照对比 | `MQTT/` | MQTT 抓包快照 vs 实时 Modbus |
| Datalog | 快照对比 | `Datalog/` | Datalog JSON 导出 vs 实时 Modbus |
| Ethernet/IP | 快照对比 | `EtherNetIP/enip_comparator.py` | EtherNet/IP 读值 vs 实时 Modbus |
| SNMP | — | 未实现 | — |

**公共模块：** `bacnet_reader.py`（BAC0 封装）、`modbus_reader.py`（FC03 Float32）、`template_reader.py`（blockParams xlsx 解析）、`config.py`（所有可配参数集中入口）

## 运行方式
```bash
# BACnet 比对（默认含范围检查 + 单位检查 + 数值比对）
python comparator.py --device acurev4100
python comparator.py --device acurev4100 --quick    # 只比对前30个参数
python comparator.py --device acurev4100 --no-meta  # 跳过单位检查

# AcuCloud 比对
python cloud_comparator.py --device acurev4100
python cloud_comparator.py --device acurev4100 --row 1  # 指定数据行
```

## 网络配置（当前测试环境）
- 网关 IP：192.168.2.209，BACnet 端口：48000
- 本机 BACnet 监听：192.168.2.45:47808
- 各设备 Modbus 配置见 config.py 的 MODBUS_DEVICE_MAP

## 报告输出
HTML 报告输出至 reports/ 目录，文件名格式：
- compare_<设备名>_<时间戳>.html（BACnet比对）
- cloud_<设备名>_<时间戳>.html（AcuCloud比对）
报告含三段可折叠区块（范围检查 / 单位检查 / 数值比对），summary行显示彩色badge。

## 目录结构
```
Protocols/                   ← 代码根目录（已并入主干）
├── BACnetIP/
│   ├── bacnet_reader.py
│   └── comparator.py
├── AcuCloud/
│   └── cloud_comparator.py
├── EtherNetIP/
│   ├── enip_reader.py
│   └── enip_comparator.py
├── modbus_reader.py
├── template_reader.py
├── config.py
├── devices/          ← Modbus地址表模块（每设备一个.py）
├── template/         ← 参数模板xlsx（blockParams sheet，按设备子目录）
├── Datas/acuclouddatas/  ← AcuCloud导出xlsx快照
└── reports/          ← HTML报告输出（自动创建）
```

## 接线检查自动化测试（Wiring Check）

脚本路径：`test_case/ACM_41_WEB2/wiring_check/`

### 目录结构
```
wiring_check/
├── core/                    基础设施
│   ├── config.py            寄存器地址 + 连接参数（读自 Protocols/config.py）
│   ├── meter_modbus.py      Modbus TCP 读写（192.168.2.242:502）
│   ├── signal_driver.py     控源封装（set_ac）
│   ├── expected_engine.py   5种接线方式算法引擎（推导预期 Wiring Status）
│   ├── wiring_check_page.py Playwright 页面对象（登录/触发/解析）
│   └── report.py            Excel 报告生成
├── conftest.py              pytest session fixtures
├── test_3e4wy.py            3E4WY（27条）
├── test_2e3w_delta.py       2E3W Delta（17条）
├── test_2e3w_network.py     2E3W Network（27条）
├── test_2e3w_1phase.py      2E3W 1Phase（13条）
├── test_1e2w.py             1E2W（4条）
└── reports/                 Excel 报告输出（自动创建）
```

### 运行方式
```bash
# pytest（推荐，全部88条）
pytest test_case/ACM_41_WEB2/wiring_check/ -v
pytest test_case/ACM_41_WEB2/wiring_check/ -k "V-"   # 只跑电压
pytest test_case/ACM_41_WEB2/wiring_check/ -k "I-"   # 只跑电流

# 直接运行（生成 Excel 报告）
python test_case/ACM_41_WEB2/wiring_check/test_3e4wy.py
```

### 用例数量

| 接线方式 | 电压 | 电流 | 合计 |
|---------|------|------|------|
| 3E4WY | 15 | 12 | 27 |
| 2E3W Delta | 11 | 6 | 17 |
| 2E3W Network | 15 | 12 | 27 |
| 2E3W 1Phase | 7 | 6 | 13 |
| 1E2W | 2 | 2 | 4 |
| **合计** | | | **88** |

### 关键配置
- 设备名：`DEVICE_NAME = '41002242'`（test_3e4wy.py 第36行）
- Meter 连接：从 `Protocols/config.py` 的 `MODBUS_DEVICE_MAP['AcuRev4100']` 读取
- WEB2 登录：`core/wiring_check_page.py` 的 `DEFAULT_USER / DEFAULT_PASS`
- 用例摘要：→ testcase/wiring_check_v1.md

## 支持设备
全部10个设备已适配，详见 CLAUDE.md 设备速查表。
IOM-03/04 为已知限制（FC02暂不支持），其余均完整支持。
