# Azure IoT 测试套件

AcuHMI-1-7 网关 **Azure IoT Hub** 北向协议自动化测试。  
使用 Playwright 操作 Web UI，通过 azure-eventhub SDK 订阅 Event Hub 兼容端点验证数据上报。

> 本套件为 AWS IoT 测试套件的一对一移植，认证方式由 X.509 证书改为 **Connection String**（共享访问密钥），订阅验证由 paho-mqtt 改为 **azure-eventhub**。

---

## 目录结构

```
azure_iot/
├── README.md                    # 本文件
├── config.yaml                  # 测试配置（Connection String、Event Hub、Modbus 设备）
├── conftest.py                  # pytest fixtures（azure_cfg / azure_page / azure_session_page）
├── pytest.ini                   # 本地 markers 定义（azure_iot / slow / manual）
├── subscribe_messages.py        # 手动订阅 Event Hub 消息脚本（azure-eventhub SDK）
├── certs/                       # 凭据/证书目录
│   ├── README.md                # 说明 Azure IoT 认证方式与证书占位说明
│   └── *.pem / *.key / *.crt   # 如使用 X.509 认证时放置对应证书
├── TestCase_AcuHMI-1-7_AZR_001_001.py   # 配置开启与关闭
├── TestCase_AcuHMI-1-7_AZR_002_001.py   # Connection String 参数校验
├── TestCase_AcuHMI-1-7_AZR_002_002.py   # Interval 11档验证（参数化）
├── TestCase_AcuHMI-1-7_AZR_003_001.py   # 合法 Connection String 连接 + Event Hub 收数据
├── TestCase_AcuHMI-1-7_AZR_003_002.py   # 错误 Connection String 连接失败
├── TestCase_AcuHMI-1-7_AZR_004_001.py   # AcuRev4100 数据三段式验证
├── TestCase_AcuHMI-1-7_AZR_004_002.py   # AcuRev2100 数据三段式验证
├── TestCase_AcuHMI-1-7_AZR_004_003.py   # Acuvim3 数据三段式验证
├── TestCase_AcuHMI-1-7_AZR_004_004.py   # AcuvimIIR 数据三段式验证
├── TestCase_AcuHMI-1-7_AZR_004_005.py   # AcuvimIIW 数据三段式验证
├── TestCase_AcuHMI-1-7_AZR_004_006.py   # AcuRev1300 数据三段式验证
├── TestCase_AcuHMI-1-7_AZR_004_007.py   # Virtual 虚拟设备数据比对
├── TestCase_AcuHMI-1-7_AZR_004_008.py   # 未勾选设备时 Save 被阻止
├── TestCase_AcuHMI-1-7_AZR_004_009.py   # 无参数配置时无数据上报
├── TestCase_AcuHMI-1-7_AZR_005_001.py   # 断网 24h 重传（manual skip）
├── TestCase_AcuHMI-1-7_AZR_005_002.py   # 断网 72h 重传（manual skip）
├── TestCase_AcuHMI-1-7_AZR_005_003.py   # 断网 >72h 边界（manual skip）
├── TestCase_AcuHMI-1-7_AZR_005_005.py   # 补传非阻塞（manual skip）
├── TestCase_AcuHMI-1-7_AZR_006_001.py   # Disable 停止/Re-enable 恢复
├── TestCase_AcuHMI-1-7_AZR_006_002.py   # Azure IoT + AWS IoT 共存
├── TestCase_AcuHMI-1-7_AZR_006_003.py   # 多设备性能（10分钟）
└── TestCase_AcuHMI-1-7_AZR_006_004.py   # 下游设备离线行为（manual）
```

---

## 用例分组

| 分组 | 文件编号 | 测试内容 | 级别 |
|------|---------|---------|------|
| 001 开关 | AZR_001_001 | Enable/Disable 切换，字段显隐 | LV0 |
| 002 参数校验 | AZR_002_001 | Connection String 格式/长度/合法值/空值校验 | LV1 |
| 002 间隔 | AZR_002_002 | 11 档 Interval 实际上报间隔验证（±1s） | LV1 |
| 003 连接 | AZR_003_001 | 合法 Connection String 连接成功 + Event Hub 收到数据 | LV0 |
| 003 连接 | AZR_003_002 | 错误 Key/HostName/DeviceId/乱码 连接失败 | LV2 |
| 004 设备 | AZR_004_001~006 | 各型号设备数据三段式验证（参数完整性/单位/Modbus比对） | LV1 |
| 004 设备 | AZR_004_007 | 虚拟设备 Event Hub 数据 vs Reading 页面比对 | LV3 |
| 004 设备 | AZR_004_008 | 未勾选设备时 Save 被阻止 | LV2 |
| 004 设备 | AZR_004_009 | 无参数配置时连接成功但无上报数据 | LV1 |
| 005 断网 | AZR_005_001~005 | 断网 24h/72h/72h+ 重传（手动执行） | LV1~3 |
| 006 综合 | AZR_006_001 | Disable 停止上报 → Re-enable 恢复 | LV3 |
| 006 综合 | AZR_006_002 | Azure IoT 与 AWS IoT 同时启用互不干扰 | LV3 |
| 006 综合 | AZR_006_003 | 全设备多参数 10 分钟性能测试 | LV3 |
| 006 综合 | AZR_006_004 | 下游设备离线时上报行为（手动） | LV3 |

---

## 前置条件

1. **网关**：AcuHMI-1-7 已上电，Web UI 可访问（`https://192.168.2.8`）
2. **Python 环境**：已安装 `playwright`、`pytest`、`azure-eventhub`、`allure-pytest`
3. **Playwright 浏览器**：`playwright install chromium`
4. **Azure IoT Hub**：已创建 IoT Hub 和设备，获取 Connection String 和 Event Hub 兼容端点连接串
5. **下游设备**：按 `config.yaml` 中 `modbus.devices` 接好设备并在线

---

## 配置文件 config.yaml

填写以下字段后即可运行测试：

```yaml
gateway:
  url: "https://192.168.2.8"
  username: "admin"
  password: "Admin@110002"

azure_iot:
  # 从 Azure Portal → IoT Hub → 设备 → 连接字符串 获取
  primary_conn_str: "HostName=<hub>.azure-devices.net;DeviceId=<device-id>;SharedAccessKey=<key>"
  secondary_conn_str: ""
  interval: "30 seconds"

  # 从 Azure Portal → IoT Hub → 内置终结点（Event Hub 兼容端点）获取
  eventhub_conn_str: "Endpoint=sb://<namespace>.servicebus.windows.net/;SharedAccessKeyName=iothubowner;SharedAccessKey=<key>;EntityPath=<hub>"

  expected_devices: [AcuRev1300, AcuvimIIR, AcuRev4100, AcuRev2100, Acuvim3, AcuvimIIW, test_virtual01]
  virtual_device: "test_virtual01"
  tolerance_pct: 5.0    # 数值比对容差（百分比）
  tolerance_abs: 1.0    # 数值比对绝对容差

modbus:
  devices:
    AcuRev4100: ["192.168.2.242", 502, 1]
    ...
```

---

## 执行命令

```powershell
# 从 azure_iot/ 目录执行
cd C:\work\autotest\autotest\AcuHMI-1-7\tests\protocols\azure_iot

# 全套 Azure IoT 用例
pytest . -m azure_iot -v

# 仅运行快速用例（跳过 slow/manual）
pytest . -m "azure_iot and not slow and not manual" -v

# 单文件执行
pytest TestCase_AcuHMI-1-7_AZR_001_001.py -v
```

也可从 AcuHMI-1-7/ 根目录执行（需指定 `-c`）：

```powershell
cd C:\work\autotest\autotest\AcuHMI-1-7
pytest tests/protocols/azure_iot/ -c tests/protocols/azure_iot/pytest.ini -m "azure_iot and not slow and not manual" -v
```

---

## Allure 报告

**每次 pytest 执行结束后（含单条用例）自动生成报告，无需手动执行任何命令。**

报告保存在 `reports/allure-report-YYYYMMDD_HHMMSS/` 目录下，每次独立存放不覆盖。

打开最新报告：

```powershell
cd C:\work\autotest\autotest\AcuHMI-1-7\tests\protocols\azure_iot
$latest = Get-ChildItem reports/allure-report-* -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
& "C:\work\tools\allure\allure-2.32.0\bin\allure.bat" open $latest.FullName
```

---

## 手动订阅 Event Hub 消息

```powershell
# 持续监听（Ctrl+C 停止）
python tests/protocols/azure_iot/subscribe_messages.py --pretty

# 监听 60 秒后自动退出
python tests/protocols/azure_iot/subscribe_messages.py --timeout 60
```

需要在 `config.yaml` 中填写 `eventhub_conn_str`。

---

## Fixtures 说明

| Fixture | Scope | 说明 |
|---------|-------|------|
| `azure_cfg` | session | 加载 config.yaml，整个会话共享 |
| `azure_page` | function | 每条用例独立导航，用例结束后自动 Disable |
| `azure_session_page` | session | 整个 session 共用同一页面，用于 Interval 参数化测试 |

---

## 与 AWS IoT 的主要差异

| 项目 | AWS IoT | Azure IoT |
|------|---------|-----------|
| 认证方式 | X.509 证书（.pem/.key） | Connection String（SharedAccessKey） |
| 订阅验证 | paho-mqtt + TLS | azure-eventhub SDK |
| 端点配置 | URL + ClientID + Topic | Primary Connection String |
| 本地订阅工具 | `subscribe_messages.sh`（mosquitto_sub） | `subscribe_messages.py`（azure-eventhub） |
| 连接串存放 | certs/ 目录下证书文件 | config.yaml `primary_conn_str` 字段 |
