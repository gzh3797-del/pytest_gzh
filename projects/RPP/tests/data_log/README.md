# DataLog 测试套件

AcuHMI-1-7 网关 **Data Log（数据日志）** 功能自动化测试。  
测试网关将 Logger 采集数据通过 FTP / SFTP / HTTP / HTTPS 推送至本地服务器，并验证文件内容与 Modbus 实时读数的一致性。

---

## 目录结构

```
data_log/
├── README.md                        # 本文件
├── config.yaml                      # 测试配置（网关、Modbus 设备、推送服务器地址）
├── config.py                        # 配置加载辅助模块
├── conftest.py                       # pytest fixtures（driver / pool / 服务器启停）
├── pytest.ini                        # 本地 markers 定义（datalog / push / disable / ui）
├── datalog_page.py                   # Page Object：DataLoggers Web UI 操作封装
├── datalog_comparator.py             # 数据比对核心：CSV/JSON 字段 vs Modbus 读数
├── datalog_server_verifier.py        # 服务器接收文件验证（FTP/SFTP/HTTP/HTTPS）
├── helpers.py                        # 通用辅助函数（文件收集、时间解析等）
├── template_reader.py                # 参数模板读取（列名/单位/比例）
├── setup_env.py                      # 测试前环境初始化（服务器启动、网关配置下发）
├── teardown_env.py                   # 测试后环境清理（服务器停止、Logger 重置）
├── servers/                          # 本地服务器实现
│   ├── ftp_server.py                 # FTP 服务器（pyftpdlib，端口 2121）
│   ├── sftp_server.py                # SFTP 服务器（paramiko，端口 2222）
│   ├── http_server.py                # HTTP/HTTPS 服务器（端口 8080/8443）
│   └── __init__.py
├── TestCase_AcuHMI_003_01_case*.py   # Logger1 测试用例
├── TestCase_AcuHMI_003_02_case*.py   # Logger2 测试用例
├── TestCase_AcuHMI_003_03_case*.py   # Logger3 测试用例
├── TestCase_AcuHMI_003_04_case*.py   # Logger4 / 跨Logger 测试用例
└── TestCase_AcuHMI_003_06_case*.py   # Logger 高级场景测试用例
```

---

## 用例分组

### Logger1（003_01，11 个用例）

| 用例 ID | 内容 | 验证协议 |
|---------|------|---------|
| case01 | Logger1 Disable → FTP/SFTP 无文件 | FTP / SFTP |
| case02 | Logger1 Disable → FTP/SFTP/HTTP 无文件 | FTP / SFTP / HTTP |
| case03 | Logger1 Enable → FTP 文件内容验证 | FTP |
| case04 | Logger1 Enable → SFTP 文件内容验证 | SFTP |
| case05 | Logger1 Enable → HTTP 文件内容验证 | HTTP |
| case06 | Logger1 Enable → HTTPS 文件内容验证 | HTTPS |
| case15 | Logger1 文件名格式校验 | FTP |
| case16 | Logger1 字段完整性校验 | FTP |
| case17 | Logger1 数值比对（vs Modbus） | FTP |
| case18 | Logger1 多设备并发推送 | FTP / SFTP |
| case19 | Logger1 Interval 精度验证 | FTP |

### Logger2（003_02，13 个用例）

| 用例 ID | 内容 | 验证协议 |
|---------|------|---------|
| case01 | Logger2 Disable → FTP/SFTP 无文件 | FTP / SFTP |
| case02 | Logger2 Disable → 多协议无文件 | FTP / SFTP / HTTP |
| case03~06 | Logger2 Enable → FTP/SFTP/HTTP/HTTPS 文件内容验证 | 各协议 |
| case07~10 | Logger2 字段完整性/数值比对/Interval | FTP |
| case11 | Logger2 文件推送超时容错 | FTP |
| case12 | Logger2 大批量参数推送 | FTP / SFTP |
| case13 | Logger2 文件格式切换（CSV↔JSON） | FTP |

### Logger3（003_03，13 个用例）

与 Logger2 用例结构相同，测试对象为 Logger3（Post Channel 独立配置）。

### Logger4 / 跨 Logger（003_04，2 个用例）

| 用例 ID | 内容 |
|---------|------|
| case20 | 多 Logger 并发推送性能测试 |
| case21 | Logger 配置切换后行为一致性 |

### 高级场景（003_06，2 个用例）

| 用例 ID | 内容 |
|---------|------|
| case01 | 网络中断后重连补传验证 |
| case02 | 大规模参数全量推送压力测试 |

---

## 前置条件

1. **网关**：AcuHMI-1-7 已上电，Web UI 可访问（`https://192.168.2.8`）
2. **本地服务器**：测试机 IP `192.168.2.149`，端口 2121（FTP）/ 2222（SFTP）/ 8080（HTTP）/ 8443（HTTPS）未被占用
3. **Python 环境**：已安装 `playwright`、`pytest`、`pyftpdlib`、`paramiko`、`pymodbus`
4. **下游设备**：在网关下挂好设备并在线。连接参数（ip/port/unit）运行时由网关 API 自动发现，按型号取第一台在线设备，无需手填；`config.yaml` 的 `modbus.tcp.devices` 仅作发现失败时的兜底（详见下文「设备连接信息的获取」）

---

## 配置文件 config.yaml

```yaml
gateway:
  url: "https://192.168.2.8"
  username: "admin"
  password: "<见 .env>"

modbus:
  # Modbus TCP 设备：连接参数（ip/port/unit）自 2026-06 起由网关 API 动态发现，
  # 以下 tcp.devices 静态表仅作「动态发现失败时的兜底」，正常运行不生效。
  tcp:
    devices:
      AcuRev4100: {ip: "192.168.2.30", port: 502, unit: 102}
      AcuRev4100a: {ip: "192.168.2.29", port: 502, unit: 203}
      AcuRev4100b: {ip: "192.168.3.44", port: 502, unit: 1}
      AcuRev2100: {ip: "192.168.2.64",  port: 502, unit: 101}
      AcuvimIIW:  {ip: "192.168.2.27",  port: 502, unit: 2}
      AcuvimIIR:  {ip: "192.168.2.8",  port: 502, unit: 103}
      AcuRev1300:  {ip: "192.168.2.8",  port: 502, unit: 102}
      
      

  # Modbus RTU 设备：列表形式，每条串口独立配置，挂载各自设备
  # 新增串口线路直接在列表中追加一项即可
  rtu:
    - port:     "COM6"      # 串口号（Windows: COMx，Linux: /dev/ttyUSBx）
      baudrate: 19200
      parity:   "N"         # N=无校验 / E=偶校验 / O=奇校验
      stopbits: 1
      bytesize: 8
      devices:
        AcuvimIIR:  {unit: 1}
        AcuRev1300: {unit: 2}

    - port:     "COM3"
      baudrate: 9600
      parity:   "E"
      stopbits: 1
      bytesize: 8
      devices:
        AcuvimIIW_RTU: {unit: 1}

datalog:
  default_device: "AcuRev4100"
  tolerance_pct: 5.0     # 数值比对百分比容差
  tolerance_abs: 0.05    # 数值比对绝对容差
  push_timeout: 300      # 等待推送最大秒数

server:
  host: "192.168.2.149"
  ftp:
    port: 2121
    user: "datalog"
    password: "datalog123"
  sftp:
    port: 2222
    user: "datalog"
    password: "datalog123"
  http:
    port: 8080
  https:
    port: 8443
```

---

## 执行命令

**从仓库根一条命令即可**：把 data_log 目录作为参数传给 pytest，pytest 会自动用该目录的
`pytest.ini`（rootdir 锁定到此，含 `TestCase_*` 收集规则、markers、Allure 配置），无需 `-c`。
`conftest` 会在本次会话内自动启动所需服务器、登录并下发网关配置，跑完自动停服务器——
**无需** `setup_env.py` / `teardown_env.py`。

```powershell
cd C:\work\autotest\autotest

pytest projects/RPP/tests/data_log/ -k "003_01" -v   # 全部 Logger1 用例
pytest projects/RPP/tests/data_log/ -k "003_02" -v   # 全部 Logger2 用例
pytest projects/RPP/tests/data_log/ -v               # 全部用例
pytest projects/RPP/tests/data_log/TestCase_AcuHMI_003_01_case01.py -v   # 单条用例
```

> 用 `-k` 关键字过滤用例，不要用 `TestCase_*.py` 通配符（PowerShell 不展开 `*`）。
> 若之前跑过 `setup_env.py` 残留了 `.setup_done`，会让 conftest 跳过自动起服务器；
> 先 `python projects/RPP/tests/data_log/teardown_env.py` 清掉再直跑。

### 可选：跨多次运行复用环境（`setup_env.py`）

频繁反复跑时，网关配置下发较慢。可用 `setup_env.py` **预配置一次并常驻**（写 `.setup_done`，
后续 pytest 自动跳过重配、复用其服务器），最后用 `teardown_env.py` 收尾：

```powershell
python projects/RPP/tests/data_log/setup_env.py                 # 另开终端，阻塞运行，Ctrl+C 前保持
pytest projects/RPP/tests/data_log/ -k "003_01" -v    # 可重复多次，秒级进入用例
python projects/RPP/tests/data_log/teardown_env.py              # 停服务器、清 .setup_done
```

---

## 报告（同时产出两份）

**每次 pytest 执行结束后（含单条用例）自动生成报告，无需手动执行任何命令。** 直跑 `pytest .../data_log/`
会**同时**产出以下两份报告：

| 报告 | 位置 | 说明 |
|------|------|------|
| pytest-html | `reports/AcuHMI_1_7/data_log/<时间戳>/report.html` | 单文件（self-contained），与 BACnet/接线检查等其它模块同款，直接双击即可看用例通过/失败/跳过 + 耗时 |
| Allure | `projects/.../data_log/reports/allure-report-<时间戳>/` | 富交互报告（需本地 HTTP 服务器打开，见下） |

> pytest-html 报告由 `conftest.py` 的 `pytest_configure` 自动注入 `--html`（命令行已带 `--html`
> 或经 framework runner 运行时不覆盖，复用其 `REPORT_DIR`）。

### Allure 报告

报告保存在 `reports/allure-report-YYYYMMDD_HHMMSS/` 目录下，每次独立存放不覆盖。

> **注意**：不能直接双击 `index.html` 打开——浏览器会因 CORS 策略拦截本地 JSON 数据请求，导致页面空白。必须通过本地 HTTP 服务器访问。

**方式一：右键菜单（推荐）**

首次需双击 `C:\work\autotest\autotest\install_allure_menu.reg` 安装（一次性），之后：

1. 在资源管理器中找到 `reports/allure-report-*/index.html`
2. 右键 → **Open Allure Report (HTTP Server)**

**方式二：拖拽 open_report.bat**

将 `reports/allure-report-*` 目录拖拽到 `C:\work\autotest\autotest\open_report.bat` 上即可。

**方式三：PowerShell 命令**

```powershell
cd C:\work\autotest\autotest\projects\AcuHMI_1_7\tests\data_log
$latest = Get-ChildItem reports/allure-report-* -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
C:\work\autotest\autotest\open_report.ps1 $latest.FullName
```

---

## 设备连接信息的获取

被测设备的 Modbus 连接参数（ip / port / slave unit）**运行时从网关 REST API 动态发现**，
无需在 `config.yaml` 手动维护（与 BACnet / 接线检查一致）：

1. `conftest.py` 的 `driver` fixture 复用已登录网关的 page，调
   `helpers/physical_devices_reader.discover_modbus_tcp_devices()`，经
   `/api/device/list/modbus` + `/api/device/config/<serial>` 拿到每台下挂 Modbus TCP
   设备的 ip/port/slaveAddress，挂到 `config.DISCOVERED_DEVICES`。
2. 比对时 `datalog_server_verifier.verify_files` 按文件名识别出设备型号，调
   `pick_device_for_template()` **取该型号第一台在线设备**的连接参数；多台同型号
   （如多台 AcuRev4100）只测第一台。
3. **降级与兜底**：发现 API 失败 / 返回空 / 某型号无在线设备（含型号串解析不出，如
   AcuRev1300 deviceModel 待实测）时，自动回退 `config.yaml` 的 `modbus.tcp.devices`
   静态表，行为与改造前一致，不会中断。
4. **RTU 设备**：网关 API 仅发现 Modbus TCP，RTU 设备（`modbus.rtu`）仍走 `config.yaml`
   静态配置。

> 因此 `config.yaml` 的 `tcp.devices` 现仅在动态发现不可用时兜底；正常运行不依赖其 ip/port/unit。

---

## 数据比对说明

验证分三个阶段：

1. **文件接收**：服务器在 `push_timeout` 内收到 `.csv` 或 `.json` 文件
2. **字段完整性**：上报文件中包含参数模板要求的全部字段，单位一致
3. **数值比对**：文件中参数值与 Modbus 实时读数误差在 `tolerance_pct`（5%）或 `tolerance_abs`（0.05）内

比对逻辑见 `datalog_comparator.py`，服务器接收逻辑见 `datalog_server_verifier.py`。

---

## Fixtures 说明

| Fixture | Scope | 说明 |
|---------|-------|------|
| `driver` | session | Playwright 浏览器驱动，完成网关登录、参数配置，并动态发现下挂 Modbus TCP 设备填入 `config.DISCOVERED_DEVICES` |
| `pool` | session | 协议服务器池（FTP/SFTP/HTTP 实例），含接收目录路径 |
