# MQTT 测试套件

AcuHMI-1-7 网关 MQTT 功能测试，代码同时兼容 WEB2（AcuRev-4100-WEB2）网关，可通过一行配置切换。

| 模块 | 文件 | 用途 |
|------|------|------|
| **配置页面自动化** | `test_mqtt.py` + `conftest.py` + `mqtt_page.py` | Selenium 操作网关 Web UI，完成 MQTT 参数配置并验证回读；含集成测试（数据收发、SSL、Retained、KeepAlive） |
| **数据比对工具** | `mqtt_comparator.py` | MQTT JSON 数据 vs Modbus 寄存器三段式比对，生成 HTML 报告，支持快照/实时双模式 |
| **证书生成工具** | `gen_certs.py` | 生成 SSL/mTLS 自签名证书套件（CA + Server + Client） |

---

## 目录结构

```
projects/acuhmi_1_7/mqtt/
├── test_mqtt.py          # 配置自动化主文件（72 条用例）
├── conftest.py           # pytest session 级 fixtures（Selenium driver / mqtt_page）
├── mqtt_page.py          # MQTT 配置页面对象（Page Object，兼容 WEB2 / HMI1-7）
├── mqtt_comparator.py    # MQTT vs Modbus 数据比对工具（独立 CLI）
├── gen_certs.py          # SSL/mTLS 证书生成工具
├── certs/                # 证书目录（ca.crt / server.crt / client.crt 等）
└── README.md             # 本文档
```

---

## 一、项目切换（WEB2 ↔ HMI1-7）

测试套件支持两套设备配置，通过 `projects/acuhmi_1_7/settings.py` 顶部的 `PROJECT` 变量一键切换：

```python
# projects/AcuHMI_1_7/settings.py
PROJECT = "HMI17"   # ← 只改这一行即可切换
                    #   "WEB2"  — AcuRev-4100-WEB2 网关
                    #   "HMI17" — AcuHMI-1-7 网关（默认）
```

切换后无需改其他文件，以下参数均自动跟随切换：

| 参数 | WEB2 | HMI17 |
|------|------|-------|
| `GATEWAY_WEB_URL` | `https://192.168.3.56` | 读取 `config.yaml` `hmi_url` |
| `GATEWAY_WEB_USER` | `l` | 读取 `.env` `HMI_USERNAME` |
| `GATEWAY_WEB_PASS` | `1` | 读取 `.env` `HMI_PASSWORD` |
| `DEVICE_NAME` | `Basic Module` | 读取 `config.yaml` `mqtt_default_device` |

> **注意**：切换到 WEB2 时，需确保 `config.yaml` 中 `meter_device_name` 指向 WEB2 下挂设备（其 ip/port/unit 在 `device_modbus` 段维护），而非 HMI17 的 AcuRev4100。

---

## 二、配置文件

HMI17 项目配置分两层：

### 2.1 config.yaml（主配置，按环境修改）

```yaml
# 网关 Web UI
base_url: https://192.168.3.71
hmi_url:  https://192.168.3.71
hmi_ip:   192.168.3.71

# Modbus 直连（用于数值比对）：填 device_modbus 段下的设备名，
# 对应 ip/port/unit 在 device_modbus 段统一维护。
meter_device_name: Acurev4100242

# MQTT
mqtt_broker_address: "www.accu.com"   # 填入设备 MQTT 配置页的 Broker Address
mqtt_broker_port: 1883
mqtt_ssl_port: 8883
mqtt_collect_timeout: 60
mqtt_default_device: "AcuRev4100"    # 对应 MODBUS_DEVICE_MAP 中的 key
```

### 2.2 settings.py（代码层配置，一般不需改）

- 读取 `config.yaml` + `.env` 文件，暴露统一常量
- 顶部 `PROJECT` 变量控制项目切换
- `MODBUS_DEVICE_MAP` 定义各设备的 Modbus TCP 直连参数

---

## 三、配置页面自动化（test_mqtt.py）

### 3.1 快速开始

> 所有命令从**仓库根目录**执行。

```bash
# 全部用例
pytest projects/AcuHMI_1_7/mqtt/test_mqtt.py -v

# 只跑冒烟（lv0，约 2 条）
pytest projects/AcuHMI_1_7/mqtt/test_mqtt.py -v -m lv0

# 跑高优先级（lv0 + lv1，约 40 条）
pytest projects/AcuHMI_1_7/mqtt/test_mqtt.py -v -m "lv0 or lv1"

# 排除需要真实网络的集成测试（离线环境）
pytest projects/AcuHMI_1_7/mqtt/test_mqtt.py -v -m "not integration and not manual"

# 只跑集成测试
pytest projects/AcuHMI_1_7/mqtt/test_mqtt.py -v -m integration
```

### 3.2 用例分布

| 测试类 | 描述 | 用例数 | 标记 |
|--------|------|--------|------|
| `TestMQTT001Entry` | 启停与页面入口（页面跳转、默认状态、URL 直达） | 4 | lv0/lv1/lv3 |
| `TestMQTT002General` | General 参数校验（Broker Address/Port/Client ID/Keep Alive/Timeout/Clean Session） | 18 | lv1/lv2 |
| `TestMQTT003Credential` | User Credential（匿名连接/合法凭据/密码掩码） | 4 | lv1/lv3 |
| `TestMQTT004SSL` | SSL/TLS 配置（Enable 切换/证书上传/格式校验） | 6 | lv1/lv2/lv3 |
| `TestMQTT005LWT` | Last Will and Testament（Enable 联动/Topic/QoS 0-2） | 7 | lv1/lv2 |
| `TestMQTT006Topic` | Topic and Parameter Selection（Base Topic/QoS/Retained/Interval/Payload/设备勾选） | 14 | lv1/lv2 |
| `TestMQTT007SaveRead` | 保存与回读（全 Tab 持久化/未保存切换/分 Tab 互不覆盖/重启后持久） | 4 | lv0/lv3 |
| `TestMQTT008TestConnection` | Test Connection 弹窗（Broker 在线/不可达/凭证错误/SSL 握手失败） | 4 | lv1/lv2 + integration |
| `TestMQTT009DataValidation` | 数据传输验证（Plain/SSL 收数据 + Modbus 三段比对） | 4 | lv1 + integration |
| `TestMQTT010KeepAlive` | Keep Alive 心跳配置（10s/60s/120s 回读；PINGREQ 时序手工） | 4 | lv1/manual |
| `TestMQTT011Retained` | Retained 保留消息（Yes新订阅者立即收/No不收历史/Yes→No清除） | 3 | lv1 + integration |
| **合计** | | **72** | |

### 3.3 用例标记说明

| 标记 | 定义 |
|------|------|
| `lv0` | 冒烟：每个模块核心正向主流程最小闭环 |
| `lv1` | 完整正向验证：合法边界值、完整路径 |
| `lv2` | 异常/负向：非法输入、空值、错误提示 |
| `lv3` | 补充验证：兼容性、模式切换、低频路径 |
| `integration` | 需要真实 Broker + 设备网络联通，内嵌 amqtt Broker 启动 |
| `manual` | 需要 Wireshark 等外部工具，仅手工执行，自动化中 skip |

### 3.4 测试隔离机制

- **`_restore_base()`**：集成测试前后将 MQTT 恢复为最小合法 plain 配置（5 个 Tab 依次单独 Save）
- **模块级 autouse fixture**（`_back_to_mqtt`）：每条用例执行前自动导航至 MQTT 配置页
- **try/finally 保护**：修改 Broker 地址、凭据、SSL 证书的用例均保证恢复基线值

### 3.5 集成测试前置条件

运行 `integration` 标记的用例前需满足：

1. **端口 1883/8883 空闲**（测试自带内嵌 amqtt Broker）：
   ```
   net stop mosquitto
   ```
2. **防火墙开放入站**：
   ```
   netsh advfirewall firewall add rule name="MQTT Test 1883" dir=in action=allow protocol=TCP localport=1883
   netsh advfirewall firewall add rule name="MQTT Test 8883" dir=in action=allow protocol=TCP localport=8883
   ```
3. **设备 MQTT Broker 地址**：须设为测试机可达的 IP 或域名（`config.yaml` 中 `mqtt_broker_address`）
   - HMI17 设备 Broker Address 填域名 `www.accu.com`，通过 hosts 或 DNS 解析到测试机 IP

### 3.6 amqtt 兼容性说明

WEB2/HMI17 设备使用 MQTT 3.1 协议，CONNECT 包保留位可能非零；amqtt 严格执行 MQTT 3.1.1 会拒绝连接。`test_mqtt.py` 在模块顶部已 monkey-patch `ConnectVariableHeader.reserved_flag`，无需额外处理。

---

## 四、数据比对工具（mqtt_comparator.py）

独立 CLI 工具，支持**快照模式**（读本地 JSON）和**实时采集模式**（内嵌 Broker 等待设备连入），生成参数范围 / 单位 / 数值三段式 HTML 比对报告。

### 4.1 运行方式

> 所有命令从**仓库根目录**执行。Windows 须加 `-X utf8` 参数。

```bash
# 实时采集，单设备（默认等待 60s）
python -X utf8 projects/AcuHMI_1_7/mqtt/mqtt_comparator.py --live

# 实时采集，自定义超时
python -X utf8 projects/AcuHMI_1_7/mqtt/mqtt_comparator.py --live --timeout 180

# 只做范围和单位检查，跳过 Modbus 数值比对
python -X utf8 projects/AcuHMI_1_7/mqtt/mqtt_comparator.py --live --no-modbus

# mTLS 加密模式（端口自动切换为 8883）
python -X utf8 projects/AcuHMI_1_7/mqtt/mqtt_comparator.py --live --ssl
```

---

## 五、证书管理（gen_certs.py）

```bash
# 生成证书（含测试机 IP 的 SAN，设备从外部连入时必须指定）
python projects/AcuHMI_1_7/mqtt/gen_certs.py --host 192.168.2.61

# 多网卡
python projects/AcuHMI_1_7/mqtt/gen_certs.py --host 192.168.2.61 192.168.3.9
```

证书输出到 `projects/acuhmi_1_7/mqtt/certs/`：

| 文件 | 用途 |
|------|------|
| `ca.crt` | CA 根证书（分发给设备） |
| `server.crt / server.key` | Broker 服务器证书（测试机） |
| `client.crt / client.key` | 测试客户端证书（paho 订阅） |

设备端上传 `ca.crt` + `client.crt` + `client.key`，Broker Port 改为 `8883`。

---

## 六、前置条件汇总

| 场景 | 前置操作 |
|------|---------|
| UI 自动化（非集成） | Chrome + `pip install selenium` |
| 集成测试（1883） | 停止 Mosquitto：`net stop mosquitto`；防火墙放行 TCP 1883 |
| 集成测试（SSL 8883） | 先执行 `gen_certs.py`，防火墙放行 TCP 8883，设备上传证书 |

依赖库：
```bash
pip install selenium paho-mqtt amqtt pymodbus allure-pytest cryptography python-dotenv
```

---

## 七、已知问题与注意事项

| 问题 | 原因 | 处理方式 |
|------|------|---------|
| 009 连续跑 test_001 失败 | test_002（SSL）结束后设备处于 SSL 模式，plain 1883 连不上 | 单独运行 test_001，或确保上一次测试结束时设备已恢复 plain 模式 |
| `amqtt` 拒绝设备连接（保留位非零） | MQTT 3.1 vs 3.1.1 规范差异 | `test_mqtt.py` 顶部已自动 monkey-patch，无需手动处理 |
| 内嵌 Broker 启动失败 | 端口 1883/8883 已被 Mosquitto 等占用 | `net stop mosquitto` 停止后重试 |
| SSL 集成测试等待超时 | 设备重连需时（退避计时器） | 调大 `config.yaml` 中 `mqtt_collect_timeout` 至 120s |
