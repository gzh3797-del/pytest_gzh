# MQTT 比对测试工具

基于 `mqtt_comparator.py`，支持两种模式：**快照模式**（读取本地 JSON 文件）和**实时采集模式**（内嵌 Broker 等待设备连入）。

---

## 功能概览

| 功能 | 说明 |
|------|------|
| 参数范围检查 | 模板全量 MQTT 参数 vs 设备实际发布参数，检查缺失/多余 |
| 单位检查 | 模板 unit 列 vs JSON unit 字段逐项比对 |
| 数值比对 | JSON value vs 实时 Modbus 寄存器值（支持 RTU / TCP） |
| 多设备合并报告 | 一次采集，自动按模块名匹配模板，生成包含全部设备的单一 HTML 报告 |
| `-nan` 特殊处理 | 设备输出 `-nan`/`nan` 时，参数仍计入范围检查（不视为缺失） |

报告输出到 `tools/Protocols/reports/` 目录，文件名含时间戳。

---

## 运行方式

> 所有命令须从**仓库根目录**执行。

```bash
# 快照模式（自动选最新 JSON，第一个在线模块）
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py

# 快照模式，指定设备
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --device acurev4100

# 快照模式，指定文件
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --file tools/Protocols/MQTT/4100data.json

# 实时采集，单设备（等待 60s）
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live

# 实时采集，单设备，自定义超时
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live --timeout 120

# 实时采集，多设备合并报告（推荐，一次生成全部设备报告）
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live --all-modules --timeout 180

# 只做范围和单位检查，跳过 Modbus 数值比对
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live --all-modules --timeout 180 --no-modbus
```

> **注意**：Windows 控制台默认 GBK 编码，必须加 `-X utf8` 参数，否则日志中的 emoji 会导致崩溃。

---

## 参数说明

| 参数 | 说明 |
|------|------|
| `--live` | 启用实时采集模式，代码自动启动内嵌 MQTT Broker |
| `--all-modules` | 多设备模式，采集所有连入设备并生成合并报告，须与 `--live` 合用 |
| `--timeout <秒>` | 实时采集等待时长，默认 60s，推荐 180s（等待所有设备推送） |
| `--port <端口>` | 内嵌 Broker 监听端口，默认 1883 |
| `--device <名称>` | 指定设备型号（快照/单设备模式），影响模板和 Modbus 配置 |
| `--file <路径>` | 快照模式，指定 JSON 文件路径 |
| `--module <下标>` | 快照/单设备模式，指定模块下标（默认第一个在线模块） |
| `--no-modbus` | 跳过 Modbus 读取和数值比对，仅输出范围/单位检查结果 |
| `--no-meta` | 跳过单位检查 |
| `--keys KEY1 KEY2 ...` | 只比对指定参数 |
| `--ssl` | 启用 mTLS 双向认证，默认端口自动切换为 8883；同时在 1883 开明文端口；证书不存在时自动生成 |
| `--ssl-host <域名或IP> [...]` | 自动生成证书时额外加入 SAN 的域名或 IP（设备用域名连接时必须指定），与 `--ssl` 合用 |
| `--plain-port <端口>` | `--ssl` 模式下同时监听的明文端口，默认 1883；设为 0 表示禁用明文端口 |
| `--cert-dir <目录>` | 指定证书目录（须含 ca.crt / server.crt / server.key / client.crt / client.key），与 `--ssl` 合用 |
| `--ca-cert <路径>` | 单独指定 CA 证书路径，与 `--ssl` 合用 |
| `--server-cert <路径>` | 单独指定服务器证书路径，与 `--ssl` 合用 |
| `--server-key <路径>` | 单独指定服务器私钥路径，与 `--ssl` 合用 |
| `--client-cert <路径>` | 单独指定客户端证书路径，与 `--ssl` 合用 |
| `--client-key <路径>` | 单独指定客户端私钥路径，与 `--ssl` 合用 |

---

## 支持设备

| `--device` 参数值 | 设备型号 |
|------------------|---------|
| `acurev4100` | AcuRev4100 |
| `acurev2100` | AcuRev2100 |
| `acuvimiiw` | AcuvimIIW |
| `acuvimiir` | AcuvimIIR |
| `acuvim3` | AcuVIM3 |
| `pxm350` | PXM350（AcuRev1300） |

`--all-modules` 模式下无需指定设备，脚本根据模块名自动匹配模板。

---

## 配置文件

所有连接参数在 `tools/Protocols/config.py` 中统一配置：

```python
# MQTT 内嵌 Broker
MQTT_BROKER_HOST     = "0.0.0.0"   # 监听地址
MQTT_BROKER_PORT     = 1883         # 监听端口
MQTT_COLLECT_TIMEOUT = 60           # 默认采集超时（秒）
MQTT_TOPIC           = "#"          # 订阅 Topic

# MQTT 数值比对容差
MQTT_TOLERANCE_PERCENT  = 5.0       # ±5%（相对容差）
MQTT_TOLERANCE_ABSOLUTE = 1.0       # ±1.0（绝对容差）
```

设备 Modbus 连接参数（TCP 地址 / RTU 串口）同样在 `config.py` 的 `MODBUS_DEVICE_MAP` 中配置。

---

## 前置条件

1. **端口 1883 空闲**：若系统已运行 Mosquitto 等 MQTT Broker，须先停止（`net stop mosquitto`），否则内嵌 Broker 无法绑定端口。
2. **Windows 防火墙放行**：设备从外部网络连入时，需确保防火墙允许入站 TCP 1883：
   ```
   netsh advfirewall firewall add rule name="MQTT Broker Python" dir=in action=allow protocol=TCP localport=1883
   ```
3. **设备 MQTT 配置**：设备的 MQTT Broker 地址须设为本机 IP（如 `192.168.2.61`），端口 `1883`。

---

## MQTT JSON 格式

设备推送的 JSON 须符合以下结构：

```json
{
  "timestamp": 1778643720,
  "comm_head": { "model": "ACM-41-WEB2", "sn": "..." },
  "modules": [
    {
      "name": "AcuRev4100",
      "model": "AcuRev-4110-mA",
      "sn": "...",
      "online": true,
      "reading": [
        { "param": "FREQ_Hz", "value": "50.000", "unit": "Hz" }
      ]
    }
  ]
}
```

`param` 字段直接对应模板 `paramType` 列，无需额外映射。值为 `-nan`/`nan` 的参数视为正常发布，计入范围检查但不参与数值比对。

---

## SSL / mTLS

内嵌 Broker 支持双向 TLS（mTLS），可验证设备身份并加密传输。

### 1. 生成证书

在**仓库根目录**执行：

```bash
# 仅本地测试（paho 订阅，127.0.0.1 / localhost 已自动包含）
python tools/Protocols/MQTT/gen_certs.py

# 设备从外部网络连入时，加入本机 IP 到服务器证书 SAN
python tools/Protocols/MQTT/gen_certs.py --host 192.168.2.61

# 多网卡 / 多 IP
python tools/Protocols/MQTT/gen_certs.py --host 192.168.2.61 192.168.2.62

# 自定义有效期（默认 3650 天 = 10 年）
python tools/Protocols/MQTT/gen_certs.py --host 192.168.2.61 --days 365
```

> **注意**：若未指定 `--host`，脚本会打印警告提示。如需设备通过本机非回环地址连入，必须在 SAN 中包含该 IP。

证书文件输出到 `tools/Protocols/MQTT/certs/`：

| 文件 | 用途 |
|------|------|
| `ca.key` | CA 私钥（保密，不分发） |
| `ca.crt` | CA 根证书（分发给设备和测试客户端） |
| `server.key` | Broker 服务器私钥（本机使用） |
| `server.crt` | Broker 服务器证书（本机使用） |
| `client.key` | 测试客户端私钥（paho 订阅使用） |
| `client.crt` | 测试客户端证书（paho 订阅使用） |

### 2. 设备端配置

将以下文件导入设备（如 WEB2）的 MQTT SSL 配置页面：

- `ca.crt` — 验证 Broker 身份（CA 根证书）
- `client.crt` — 设备向 Broker 证明自身身份（客户端证书）
- `client.key` — 设备客户端私钥

设备 MQTT Broker 地址设为本机 IP，端口改为 `8883`。

### 3. 启用 SSL 运行

```bash
# mTLS 模式，多设备合并报告（自动使用 8883 端口，证书路径读自 config.py）
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live --all-modules --ssl

# 指定自己的证书目录（适用于多人共用测试台、各自 IP 不同的场景）
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live --all-modules --ssl --cert-dir C:\my_certs

# 显式指定端口
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live --all-modules --ssl --port 8883

# 单设备实时采集 + mTLS
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live --ssl
```

> `--ssl` 标志启用后，若未指定 `--port`，脚本自动使用 `8883`（由 `config.py` 中 `MQTT_SSL_PORT` 控制）。

### 4. 多人环境：各自生成证书

不同测试人员 IP 不同时，每个人**单独生成一套证书**，运行时通过 `--cert-dir` 指定目录，**无需修改 `config.py`**：

```bash
# 第一步：生成含自己 IP 的证书（--out 指定存放目录）
python tools/Protocols/MQTT/gen_certs.py --host 192.168.2.100 --out C:\certs_user1

# 第二步：运行时通过 --cert-dir 传入
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live --all-modules --ssl --cert-dir C:\certs_user1
```

也可逐个指定路径（适合已有现成证书的情况）：

```bash
python -X utf8 tools/Protocols/MQTT/mqtt_comparator.py --live --ssl \
  --ca-cert C:\my_certs\ca.crt \
  --server-cert C:\my_certs\server.crt \
  --server-key C:\my_certs\server.key \
  --client-cert C:\my_certs\client.crt \
  --client-key C:\my_certs\client.key
```

### 5. 证书路径默认值（config.py）

当不传 `--cert-dir` 或单个路径参数时，脚本从 `tools/Protocols/config.py` 读取默认路径：

```python
MQTT_SSL_CERT_DIR    = _os.path.join(_BASE, "MQTT", "certs")
MQTT_SSL_CA_CERT     = _os.path.join(MQTT_SSL_CERT_DIR, "ca.crt")
MQTT_SSL_SERVER_CERT = _os.path.join(MQTT_SSL_CERT_DIR, "server.crt")
MQTT_SSL_SERVER_KEY  = _os.path.join(MQTT_SSL_CERT_DIR, "server.key")
MQTT_SSL_CLIENT_CERT = _os.path.join(MQTT_SSL_CERT_DIR, "client.crt")
MQTT_SSL_CLIENT_KEY  = _os.path.join(MQTT_SSL_CERT_DIR, "client.key")
MQTT_SSL_PORT        = 8883
```

**优先级**：`--ca-cert` 等单个参数 > `--cert-dir` > `config.py` 默认值。
