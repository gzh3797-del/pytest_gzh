# DataLog 自动化测试套件

对 AcuHMI-1-7 网关的 **Data Logger 1 / 2 / 3** 进行端到端自动化验证。
全部用例集中在 **`test_datalog_logger.py`** 一个文件，共 **30 条**，按类型参数化。

---

## 目录

- [前提条件](#前提条件)
- [用例执行规则](#用例执行规则)
- [运行命令](#运行命令)
- [用例类型说明](#用例类型说明)
- [用例清单](#用例清单)
- [文件结构](#文件结构)

---

## 前提条件

### 依赖安装

```bash
pip install pytest pytest-html pyftpdlib paramiko selenium
```

### config.py 配置

在仓库根目录 `config.py` 中填写以下项，确保网关能访问本机服务器：

| 配置项 | 说明 |
|---|---|
| `DATALOG_SERVER_HOST` | 本机 IP（网关须能访问，如 `192.168.2.149`） |
| `DATALOG_FTP_PORT` / `DATALOG_FTP_USER` / `DATALOG_FTP_PASS` | FTP 端口和账号（默认 2121 / datalog / datalog123） |
| `DATALOG_SFTP_PORT` / `DATALOG_SFTP_USER` / `DATALOG_SFTP_PASS` | SFTP 端口和账号（默认 2222 / datalog / datalog123） |
| `DATALOG_HTTP_PORT` | HTTP 端口（默认 8080） |
| `DATALOG_HTTPS_PORT` | HTTPS 端口（默认 8443） |
| `DATALOG_SSL_CERT` / `DATALOG_SSL_KEY` | HTTPS 自签名证书路径（留空则不启用 HTTPS） |
| `GATEWAY_WEB_URL` | 网关 Web UI 地址（如 `https://192.168.2.9/#/login`） |
| `GATEWAY_WEB_USER` / `GATEWAY_WEB_PASS` | 网关登录账号密码 |
| `MODBUS_DEVICE_MAP` | 各设备 Modbus IP / Port / Unit ID（C 类完整三段验证用） |

### 运行目录

**所有命令必须从仓库根目录执行**：

```bash
# 正确
pytest Protocols/Datalog/tests/test_datalog_logger.py -v

# 错误（sys.path 解析失败）
cd Protocols/Datalog/tests && pytest .
```

---

## 用例执行规则

### pytest 标记（mark）

| 标记 | 覆盖用例 | 典型执行场景 |
|------|---------|------------|
| `lv1` | A 类（6 条）+ B 类（6 条）+ E 类（3 条）= **15 条** | 功能回归必跑 |
| `lv4` | D 类（3 条）| 版本验收时执行，验证超长 Prefix 错误提示 |

> **C 类推送用例（12 条）未打 mark**，通过 `-k` 关键字或直接运行整个文件来执行（详见运行命令）。

### 执行策略

```
C 类冒烟  →  -k "case04 or case05 or case06"（Logger1 FTP/SFTP/HTTP，≈20 分钟）
功能回归  →  -m "lv1" + C 类（整文件或 -k 过滤）
版本验收  →  整个文件（全 30 条）
```

### 用例隔离机制

`conftest.py` 中的 `clear_dirs` fixture（`autouse=True`）在**每条用例开始前**自动清空
FTP / SFTP / HTTP / HTTPS 数据目录中的所有 `.json` / `.csv` 文件，保证用例间互不干扰。

### 未自动化的用例

| 用例编号 | 原因 |
|---------|------|
| TestCase_AcuHMI_003_02_case12 | Logger2 网络中断恢复，需手动拔插网线 |
| TestCase_AcuHMI_003_03_case12 | Logger3 网络中断恢复，需手动拔插网线 |

---

## 运行命令

```bash
# ── 按标记运行 ──────────────────────────────────────────────

# 功能回归（A + B + E 类，15 条，约 60 分钟）
pytest Protocols/Datalog/tests/test_datalog_logger.py -m lv1 -v

# 负向用例（D 类，3 条）
pytest Protocols/Datalog/tests/test_datalog_logger.py -m lv4 -v

# ── C 类推送用例 ─────────────────────────────────────────────

# Logger1 冒烟（FTP/SFTP/HTTP 各一条，约 20 分钟）
pytest Protocols/Datalog/tests/test_datalog_logger.py -k "case04 or case05 or case06" -v

# Logger1 全部 C 类
pytest Protocols/Datalog/tests/test_datalog_logger.py -k "003_01" -v

# Logger2 推送用例
pytest Protocols/Datalog/tests/test_datalog_logger.py -k "003_02_case0" -v

# ── 单条用例 ─────────────────────────────────────────────────

pytest Protocols/Datalog/tests/test_datalog_logger.py -k "TestCase_AcuHMI_003_01_case04" -v

# ── 全量（30 条）────────────────────────────────────────────

pytest Protocols/Datalog/tests/test_datalog_logger.py -v

# ── 生成 HTML 报告 ───────────────────────────────────────────

pytest Protocols/Datalog/tests/test_datalog_logger.py -v \
  --html=Protocols/Datalog/tests/reports/datalog.html \
  --self-contained-html

# ── 常用选项 ─────────────────────────────────────────────────

# 遇到第一条失败即停止
pytest Protocols/Datalog/tests/test_datalog_logger.py -v -x

# 简短失败原因
pytest Protocols/Datalog/tests/test_datalog_logger.py -v --tb=short
```

---

## 用例类型说明

### A 类 — Logger 开关（Disable）`test_disable_no_push`

Logger Enable 置为 off 后等待 90 秒，验证指定协议目录下**无文件**产生。

**验证点：** 文件不存在  
**标记：** `lv1`（6 条，Logger 1/2/3 × FTP-SFTP / HTTP-HTTPS）

---

### B 类 — Post Channel = None `test_none_channel_no_push`

Logger 启用但 Post Channel 选 None，等待 90 秒，验证所有协议目录下**无文件**。

**验证点：** 文件不存在  
**标记：** `lv1`（6 条，Logger 1/2/3 × CSV / JSON）

---

### C 类 — 推送验证（核心）`test_push_verify`

配置 Logger 推送至指定协议服务器，等待文件到达后立即禁用 Logger，逐项验证：

| # | 验证项 | 依据 |
|---|--------|------|
| 1 | 文件已推送到目标目录 | 超时内收到 |
| 2 | 文件扩展名与格式匹配 | Log File Format（csv / json） |
| 3 | 文件名前缀正确 | Log File Name Prefix |
| 4 | 文件名时间格式正确 | Log File Name Format（UTC Timestamp / Time Interval Format） |
| 5 | 内容时间戳格式匹配 | Timestamp Format（Local Time String / UTC Seconds / ISO8601） |
| 6 | 相邻行时间间隔正确 | Log Interval（允差 ±10%，最小 ±30s） |
| 7 | 文件行数与时间跨度正确 | Log File Length / Log Interval |
| 8 | 参数范围 + Modbus 数值比对 | 仅 case04（Logger1 FTP）执行完整三段验证 |

**标记：** 无（通过 `-k` 按 case 编号过滤）  
**用例数：** 12 条（Logger1×6、Logger2×4、Logger3×2）

---

### D 类 — Prefix 边界 `test_prefix_too_long_save_fail`

填入超长 Log File Name Prefix 后点击保存，验证页面出现错误提示。

**验证点：** 页面有错误文字  
**标记：** `lv4`（3 条，Logger 1/2/3 各一条）

---

### E 类 — Length→Interval 联动 `test_interval_linkage`

遍历 Log File Length 的 11 种选项，读取每种 Length 对应的 Log Interval 可选项，
与期望规则（`EXPECTED_INTERVALS`）对比。

**验证点：** Interval 可选项无缺失、无多余  
**标记：** `lv1`（3 条，Logger 1/2/3 各一条）

---

## 用例清单

### A 类 — Disable（6 条，`lv1`）

| 用例编号 | Logger | 检查协议 |
|---------|:------:|---------|
| TestCase_AcuHMI_003_06_case03 | 1 | FTP / SFTP |
| TestCase_AcuHMI_003_01_case02 | 1 | HTTP / HTTPS |
| TestCase_AcuHMI_003_02_case01 | 2 | FTP / SFTP |
| TestCase_AcuHMI_003_02_case02 | 2 | HTTP / HTTPS |
| TestCase_AcuHMI_003_03_case01 | 3 | FTP / SFTP |
| TestCase_AcuHMI_003_03_case02 | 3 | HTTP / HTTPS |

### B 类 — None 通道（6 条，`lv1`）

| 用例编号 | Logger | 格式 |
|---------|:------:|------|
| TestCase_AcuHMI_003_01_case03 | 1 | CSV |
| TestCase_AcuHMI_003_01_case15 | 1 | JSON |
| TestCase_AcuHMI_003_02_case03 | 2 | CSV |
| TestCase_AcuHMI_003_02_case07 | 2 | JSON |
| TestCase_AcuHMI_003_03_case03 | 3 | CSV |
| TestCase_AcuHMI_003_03_case07 | 3 | JSON |

### C 类 — 推送验证（12 条，无 mark）

| 用例编号 | Logger | 协议 | 格式 | File Length | 说明 |
|---------|:------:|------|------|------------|------|
| TestCase_AcuHMI_003_01_case04 | 1 | FTP  | CSV  | 1 min  | **完整三段验证**（范围+Modbus数值） |
| TestCase_AcuHMI_003_01_case05 | 1 | SFTP | CSV  | 5 min  | |
| TestCase_AcuHMI_003_01_case06 | 1 | HTTP | CSV  | 10 min | |
| TestCase_AcuHMI_003_01_case16 | 1 | FTP  | JSON | 1 min  | |
| TestCase_AcuHMI_003_01_case17 | 1 | SFTP | JSON | 5 min  | |
| TestCase_AcuHMI_003_01_case18 | 1 | HTTP | JSON | 10 min | |
| TestCase_AcuHMI_003_02_case04 | 2 | FTP  | CSV  | 1 min  | |
| TestCase_AcuHMI_003_02_case05 | 2 | SFTP | CSV  | 5 min  | |
| TestCase_AcuHMI_003_02_case06 | 2 | HTTP | CSV  | 10 min | |
| TestCase_AcuHMI_003_02_case08 | 2 | FTP  | JSON | 1 min  | |
| TestCase_AcuHMI_003_03_case04 | 3 | FTP  | CSV  | 1 min  | |
| TestCase_AcuHMI_003_03_case05 | 3 | SFTP | CSV  | 5 min  | |

> **未收录用例说明：**  
> Logger1 case07~14（15min～1month 文件时长）未收录，因等待时间过长（最长约 30 天），
> 需要时可手动向 `_PUSH_CASES` 追加对应 `PushCase` 行后执行。

### D 类 — Prefix 超长（3 条，`lv4`）

| 用例编号 | Logger |
|---------|:------:|
| TestCase_AcuHMI_003_01_case19 | 1 |
| TestCase_AcuHMI_003_02_case11 | 2 |
| TestCase_AcuHMI_003_03_case11 | 3 |

### E 类 — 联动验证（3 条，`lv1`）

| 用例编号 | Logger |
|---------|:------:|
| TestCase_AcuHMI_003_04_case21 | 1 |
| TestCase_AcuHMI_003_02_case13 | 2 |
| TestCase_AcuHMI_003_03_case13 | 3 |

---

## 文件结构

```
Protocols/Datalog/tests/
├── conftest.py              # pytest fixtures
├── test_datalog_logger.py   # 全部 30 条用例
├── helpers.py               # 共享工具函数（备用）
├── README.md                # 本文件
└── reports/                 # HTML 报告输出目录
```

### conftest.py — fixtures

| fixture | scope | 说明 |
|---------|-------|------|
| `pool` | session | 构建协议池（FTP / SFTP / HTTP / HTTPS `ServerInfo`） |
| `servers` | session | 启动本地协议服务器，测试结束后自动停止 |
| `driver` | session | Selenium 登录网关，配置 Post Channel 1=FTP / 2=SFTP / 3=HTTP |
| `clear_dirs` | function（autouse） | 每条用例前清空所有协议目录的 `.json` / `.csv` 文件 |
