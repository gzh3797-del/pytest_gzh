# AWS IoT 测试套件

## 功能简介

网关将下挂电表数据通过 **MQTT over TLS** 定期上报至 AWS IoT Core（亚马逊云物联网平台）。  
测试验证：配置参数校验、证书连接、各型号设备数据上报的完整性与准确性、断网重传及多协议共存场景。

---

AcuHMI-1-7 网关 **AWS IoT Core** 北向协议自动化测试。  
使用 Playwright 操作 Web UI，通过 paho-mqtt + TLS 订阅验证数据上报。

---

## 目录结构

```
aws_iot/
├── README.md                   # 本文件
├── config.yaml                 # 测试配置（网关地址、证书路径、Modbus 设备）
├── conftest.py                 # pytest fixtures（aws_cfg / aws_page / aws_session_page）
├── pytest.ini                  # 本地 markers 定义（aws_iot / slow / manual）
├── subscribe_messages.sh       # 手动订阅 MQTT 消息脚本（调用 python subscribe_aws.py）
├── certs/                      # 证书文件目录
│   ├── client.pem              # 网关上传证书（网关→AWS IoT Core）
│   ├── key.pem                 # 网关私钥
│   ├── AmazonRootCA1.pem       # Amazon 根 CA
│   ├── *-certificate.pem.crt  # 本地订阅证书
│   └── *-private.pem.key      # 本地订阅私钥
├── TestCase_AcuHMI-1-7_AWS_001_001.py   # 配置开启与关闭
├── TestCase_AcuHMI-1-7_AWS_002_001.py   # URL 参数校验
├── TestCase_AcuHMI-1-7_AWS_002_003.py   # Interval 11档验证（参数化）
├── TestCase_AcuHMI-1-7_AWS_003_001.py   # 合法证书连接 + MQTT 收数据
├── TestCase_AcuHMI-1-7_AWS_003_002.py   # 非法证书连接失败
├── TestCase_AcuHMI-1-7_AWS_004_001.py   # AcuRev4100 数据三段式验证
├── TestCase_AcuHMI-1-7_AWS_004_002.py   # AcuRev2100 数据三段式验证
├── TestCase_AcuHMI-1-7_AWS_004_003.py   # Acuvim3 数据三段式验证
├── TestCase_AcuHMI-1-7_AWS_004_004.py   # AcuvimIIR 数据三段式验证
├── TestCase_AcuHMI-1-7_AWS_004_005.py   # AcuvimIIW 数据三段式验证
├── TestCase_AcuHMI-1-7_AWS_004_006.py   # AcuRev1300 数据三段式验证
├── TestCase_AcuHMI-1-7_AWS_004_007.py   # Virtual 虚拟设备数据比对
├── TestCase_AcuHMI-1-7_AWS_004_008.py   # 未勾选设备时 Save 被阻止
├── TestCase_AcuHMI-1-7_AWS_004_009.py   # 无参数配置时无数据上报
├── TestCase_AcuHMI-1-7_AWS_005_001.py   # 断网 24h 重传（manual skip）
├── TestCase_AcuHMI-1-7_AWS_005_002.py   # 断网 72h 重传（manual skip）
├── TestCase_AcuHMI-1-7_AWS_005_003.py   # 断网 >72h 边界（manual skip）
├── TestCase_AcuHMI-1-7_AWS_005_005.py   # 补传非阻塞（manual skip）
├── TestCase_AcuHMI-1-7_AWS_006_001.py   # Disable 停止/Re-enable 恢复
├── TestCase_AcuHMI-1-7_AWS_006_002.py   # AWS IoT + Azure IoT 共存
├── TestCase_AcuHMI-1-7_AWS_006_003.py   # 多设备性能（10分钟）
└── TestCase_AcuHMI-1-7_AWS_006_004.py   # 下游设备离线行为（manual）
```

---

## 用例分组

| 分组 | 文件编号 | 测试内容 | 级别 |
|------|---------|---------|------|
| 001 开关 | AWS_001_001 | Enable/Disable 切换，字段显隐 | LV0 |
| 002 参数校验 | AWS_002_001 | URL 格式/长度/合法值/空值校验 | LV1 |
| 002 间隔 | AWS_002_003 | 11 档 Interval 实际上报间隔验证（±1s） | LV1 |
| 003 连接 | AWS_003_001 | 合法证书连接成功 + MQTT 收到数据 | LV0 |
| 003 连接 | AWS_003_002 | 过期/格式错误/无证书连接失败 | LV2 |
| 004 设备 | AWS_004_001~006 | 各型号设备数据三段式验证（参数完整性/单位/Modbus比对） | LV1 |
| 004 设备 | AWS_004_007 | 虚拟设备 MQTT 数据 vs Reading 页面比对 | LV3 |
| 004 设备 | AWS_004_008 | 未勾选设备时 Save 被阻止 | LV2 |
| 004 设备 | AWS_004_009 | 无参数配置时连接成功但无上报数据 | LV1 |
| 005 断网 | AWS_005_001~005 | 断网 24h/72h/72h+ 重传（手动执行） | LV1~3 |
| 006 综合 | AWS_006_001 | Disable 停止上报 → Re-enable 恢复 | LV3 |
| 006 综合 | AWS_006_002 | AWS IoT 与 Azure IoT 同时启用互不干扰 | LV3 |
| 006 综合 | AWS_006_003 | 全设备多参数 10 分钟性能测试 | LV3 |
| 006 综合 | AWS_006_004 | 下游设备离线时上报行为（手动） | LV3 |

---

## 前置条件

1. **网关**：AcuHMI-1-7 已上电，Web UI 可访问（`https://192.168.2.8`）
2. **Python 环境**：`c:\work\tools\python311`，已安装 `playwright`、`pytest`、`paho-mqtt`、`allure-pytest`
3. **Playwright 浏览器**：`playwright install chromium`
4. **AWS IoT Core 账号**：已创建 Thing，下载证书并放入 `certs/`
5. **下游设备**：按 `config.yaml` 中 `modbus.devices` 接好设备并在线

---

## Setup 脚本执行逻辑

每次执行 `pytest` 时，**在第一条用例运行前自动执行一次** `setup_device_config` fixture（定义于 `conftest.py`），完成以下工作：

### 执行步骤

```
pytest 启动
    │
    ▼
① 导航到 AWS IoT 配置页面
    │  使用 config.yaml 中 gateway.device_name 定位右上角导航按钮
    │  点击进入 Device Settings → Protocols → AWS IoT
    │
    ▼
② 临时 Enable（若当前为 Disable 状态）
    │  设备表格仅在 Enable 状态下可见
    │  扫描完毕后自动恢复 Disable
    │
    ▼
③ 扫描 Device Selection 表格，读取所有设备名
    │  逐行读取设备名称
    │  含 "virtual"（大小写不敏感）的行标记为虚拟设备
    │  其余标记为物理设备
    │
    ▼
④ 更新 config.yaml 的 aws_iot.expected_devices
    │  将所有设备名（物理 + 虚拟）写入文件
    │  同步更新内存中的 aws_cfg，本次 session 所有用例立即生效
    │
    ▼
⑤ 校验 modbus.tcp.devices 覆盖
    │  对比物理设备列表与 config.yaml modbus.tcp.devices 的键名
    │  虚拟设备不参与校验
    │
    ├─ 全部覆盖 → 继续执行测试用例
    │
    └─ 存在缺漏 → 终止整个 session，打印提示：
          ================================================================
          [Setup 失败] 以下物理设备未在 config.yaml modbus.tcp.devices 中配置：
            - DeviceNameA
            - DeviceNameB

          请在 config.yaml 的 modbus.tcp.devices 下补充，示例：
            DeviceNameA: {ip: "192.168.x.x", port: 502, unit: 1}
          ================================================================
```

### 处理 Setup 失败

当 Setup 检测到物理设备缺少 Modbus 配置时，**所有用例均不会执行**。  
在 `config.yaml` 的 `modbus.tcp.devices` 下补充对应设备信息后重新运行即可：

```yaml
modbus:
  tcp:
    devices:
      NewDevice: {ip: "192.168.x.x", port: 502, unit: 1}
```

> **注意**：`key` 名称必须与网关 AWS IoT 页面 Device Selection 中显示的设备名**完全一致**（区分大小写）。

### 说明

- Setup 只执行一次（`scope="session"`），不影响单条用例的运行速度
- 若设备表格为空（网关无已接入设备），跳过校验，用例正常执行
- `expected_devices` 每次运行会被**自动覆盖**为当前网关实际接入设备，无需手动维护

---

## 配置文件 config.yaml

```yaml
gateway:
  url: "https://192.168.2.9"       # 网关 Web UI 地址
  username: "admin"
  password: "Admin@110001"

aws_iot:
  url: "<endpoint>.amazonaws.com.cn"  # AWS IoT Core 端点
  client_id: "AHI260110001"           # MQTT 客户端 ID
  topic: "test/topic/aws"             # 上报 Topic
  interval: "30 seconds"              # 上报间隔

  cert_file: "tests/protocols/aws_iot/certs/client.pem"       # 网关上传证书
  key_file:  "tests/protocols/aws_iot/certs/key.pem"          # 网关私钥
  sub_cert_file: "tests/protocols/aws_iot/certs/<hash>-certificate.pem.crt"  # 订阅证书
  sub_key_file:  "tests/protocols/aws_iot/certs/<hash>-private.pem.key"      # 订阅私钥
  ca_file:       "tests/protocols/aws_iot/certs/AmazonRootCA1.pem"

  expected_devices: [AcuRev1300, AcuvimIIR, AcuRev4100, AcuRev2100, Acuvim3, AcuvimIIW, test_virtual01]
  virtual_device: "test_virtual01"
  tolerance_pct: 5.0    # 数值比对容差（百分比）
  tolerance_abs: 1.0    # 数值比对绝对容差

modbus:
  devices:
    AcuRev4100: ["192.168.2.242", 502, 1]
    AcuRev2100: ["192.168.2.64",  502, 101]
    ...
```

---

## 执行命令

```powershell
# 从 aws_iot/ 目录执行
cd C:\work\autotest\autotest\AcuHMI-1-7\tests\protocols\aws_iot

# 全套 AWS IoT 用例
pytest . -m aws_iot -v

# 仅运行快速用例（跳过 slow/manual）
pytest . -m "aws_iot and not slow and not manual" -v

# 单文件执行
pytest TestCase_AcuHMI-1-7_AWS_001_001.py -v
```

也可从 AcuHMI-1-7/ 根目录执行（需指定 `-c`）：

```powershell
cd C:\work\autotest\autotest\AcuHMI-1-7
pytest tests/protocols/aws_iot/ -c tests/protocols/aws_iot/pytest.ini -m "aws_iot and not slow and not manual" -v
```

---

## Allure 报告

**每次 pytest 执行结束后（含单条用例）自动生成报告，无需手动执行任何命令。**

报告保存在 `reports/allure-report-YYYYMMDD_HHMMSS/` 目录下，每次独立存放不覆盖。

> **注意**：不能直接双击 `index.html` 打开——浏览器会因 CORS 策略拦截本地 JSON 数据请求，导致页面空白。必须通过本地 HTTP 服务器访问。

**方式一：右键菜单（推荐）**

首次需双击 `C:\work\autotest\autotest\install_allure_menu.reg` 安装（一次性），之后：

1. 在资源管理器中找到 `reports/allure-report-*/index.html`
2. 右键 → **Open Allure Report (HTTP Server)**

**方式二：双击 open_report.bat**

直接双击本目录下的 `open_report.bat`，自动找到最新报告并打开。

**方式三：PowerShell 命令**

```powershell
cd C:\work\autotest\autotest\AcuHMI-1-7\tests\protocols\aws_iot
$latest = Get-ChildItem reports/allure-report-* -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
C:\work\autotest\autotest\open_report.ps1 $latest.FullName
```

---

## 手动订阅 MQTT 消息

```bash
bash tests/protocols/aws_iot/subscribe_messages.sh
```

脚本调用项目根目录的 `subscribe_aws.py`，确保 Python 环境已安装 `paho-mqtt`。

---

## Fixtures 说明

| Fixture | Scope | 说明 |
|---------|-------|------|
| `aws_cfg` | session | 加载 config.yaml，整个会话共享 |
| `aws_page` | function | 每条用例独立导航，用例结束后自动 Disable |
| `aws_session_page` | session | 整个 session 共用同一页面，用于 Interval 参数化测试 |
