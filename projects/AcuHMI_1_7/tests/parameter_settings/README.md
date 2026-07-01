# 参数下发自动化测试（parameter_settings）

通过 Playwright 操作网关 Web 页面下发参数，下发后用 Modbus TCP 读回寄存器值进行核验。三台电表共 **98 条**用例，一条命令跑完。

---

## 一、工作原理

```
pytest
  └─ Playwright 打开网关 Web UI（https://<GATEWAY_IP>）
       └─ 登录 → 导航至 Physical Devices → <设备> → Settings → General
            └─ 填写参数值 → 点击 Save（参数下发）
                 └─ pymodbus 直连设备 Modbus TCP 读回寄存器
                      └─ 断言 寄存器值 == 预期编码值
```

每条用例独立运行，失败时自动截图（保存到 `logs/FAIL_<用例名>.png` 并附加到 Allure 报告）。

---

## 二、目录结构

```
Protocols/parameter_settings/
├── conftest.py              # 顶层：session 级浏览器 + 登录 + 失败截图
├── README.md                # 本文档
│
├── acurev4100/              # AcuRev4100（35 条）
│   ├── conftest.py          # Allure feature/story 标签
│   ├── helpers_4100.py      # 页面操作 + Modbus 验证函数
│   ├── _src_general_settings.py   # General Settings 操作函数
│   ├── _src_event_waveform.py     # Event & Waveform 操作函数
│   └── test_Testcase_AcuHMI_AcuRev4100_*.py   # 35 个测试文件
│
├── acuvim_iiw/              # AcuvimIIW（54 条）
│   ├── conftest.py
│   ├── helpers_iiw.py
│   ├── _src_general_settings.py
│   └── test_Testcase_AcuHMI_AcuvimIIW_*.py    # 54 个测试文件
│
├── acurev2100/              # AcuRev2100（9 条）
│   ├── helpers_2100.py
│   ├── _src_general_settings.py
│   └── test_Testcase_AcuHMI_AcuRev2100_*.py   # 9 个测试文件
│
├── acurev1300/              # 占位（待实现）
├── acuvim3/                 # 占位（待实现）
└── acuvim_iir/              # 占位（待实现）
```

---

## 三、测试覆盖

### 3.1 AcuRev4100（35 条）

| 分组 | 参数 |
|------|------|
| 基本配置 | Freq / PT / NomCur / PhaseOrder / Backlight / Password |
| 能量设置 | EnergyFmt / EnergyPulse |
| 需量设置 | DmndInt / DmndRate / Demand |
| 电能质量 | VARPF / VoltInt_Hys / VoltInt_Thr |
| 电压跌落 | VoltSag_Thr / VoltSag_Hys / VoltSag_DO / VoltSag_RO |
| 电压突升 | VoltSwl_Thr / VoltSwl_Hys |
| LED 脉冲 | LedPulse / LedWidth |
| 最大需量 | MdMode / MdDate |
| 波形录制 | WfSmpl / WfPre / WfPost |
| 数字输入 | DI1 / DI2 / DI3 / DI4 / ManTrig |
| RPM | RPM |

### 3.2 AcuvimIIW（54 条）

| 分组 | 参数 |
|------|------|
| 基本配置 | CT1 / CT2 / CT41 / CT42 / I4 / PT / SrvCfg / AvgInterval / SubInterval |
| Class S 额定值 | RatedVoltage / RatedCurrent |
| Class S 电压突升 | VoltSwl_Enable / Thr / Hys / DO / RO / WF |
| Class S 电压跌落 | VoltDip_Enable / Thr / Hys / DO / RO / WF |
| Class S 电压中断 | VoltIntr_Enable / Thr / Hys / DO / RO / WF |
| Class S 电流突升 | CurrSwl_Enable / Thr / Hys / DO / RO / WF |
| Class S 不平衡电流 | UnbalCurr_Enable / Thr / Hys / DO / RO / WF |
| Class S 不平衡电压 | UnbalVolt_Enable / Thr / Hys / DO / RO / WF |
| Class S 波形采样 | WfSampleRate |
| 其他 | DmndMethod / EnergyReading / EnergyType / IDir / VarCalc / VarPF |

### 3.3 AcuRev2100（9 条）

AveragingInterval / Ct2Type / CtChannels / DemandMethod / RatedVoltage / SubInterval / VarMethod / VarPf / Wiring

---

## 四、快速开始

### 前置条件

```
pip install playwright pytest pytest-playwright allure-pytest pymodbus
playwright install chrome
```

### 运行全部 98 条（从仓库根目录执行）

```bash
pytest Protocols/parameter_settings/ -v
```

### 只跑某台设备

```bash
# AcuRev4100（35 条）
pytest Protocols/parameter_settings/acurev4100/ -v

# AcuvimIIW（54 条）
pytest Protocols/parameter_settings/acuvim_iiw/ -v

# AcuRev2100（9 条）
pytest Protocols/parameter_settings/acurev2100/ -v
```

### 生成 Allure 报告

```bash
pytest Protocols/parameter_settings/ -v --alluredir=allure-results
allure serve allure-results
```

---

## 五、环境配置

### 5.1 当前方式：环境变量（推荐，无需改代码）

`conftest.py` 优先读取以下环境变量，未设置时回退 `config.py`：

| 环境变量 | 默认值 | 说明 |
|---------|--------|------|
| `BASE_URL` | `https://192.168.3.51` | 网关 Web UI 地址 |
| `WEB_USERNAME` | `l` | 登录用户名 |
| `WEB_PASSWORD` | `1` | 登录密码 |
| `HEADLESS` | `false` | `true` = 无头模式（后台运行） |
| `SLOW_MO` | `300` | Playwright 操作间隔（毫秒），调试时可调大 |

**临时切换网关（单次运行）：**

```bash
# Windows PowerShell
$env:BASE_URL="https://192.168.3.47"; pytest Protocols/parameter_settings/ -v

# Linux / macOS
BASE_URL=https://192.168.3.47 pytest Protocols/parameter_settings/ -v
```

### 5.2 config.py（持久化配置）

修改 `Protocols/config.py` 中的 `GATEWAY_IP`：

```python
GATEWAY_IP = "192.168.3.51"    # HMI 1-7 网关 IP（parameter_settings 使用此值）
```

每台电表的 Modbus 地址独立硬编码在各自的 `helpers_*.py` 中：

| 设备 | helpers 文件 | Modbus IP | Port | Slave ID |
|------|-------------|-----------|------|----------|
| AcuRev4100 | `acurev4100/helpers_4100.py` | 192.168.2.242 | 502 | 1 |
| AcuvimIIW | `acuvim_iiw/helpers_iiw.py` | 192.168.2.27 | 502 | 2 |
| AcuRev2100 | `acurev2100/helpers_2100.py` | 192.168.2.64 | 502 | 101 |

### 5.3 YAML 配置文件（多项目/多环境支持）

如需在 **WEB2 / HMI1-7 / 其他网关** 之间切换，可使用 YAML 文件替代环境变量，实现配置与代码完全分离。

**在 `Protocols/parameter_settings/` 下创建 `config_env.yaml`：**

```yaml
# config_env.yaml — 测试环境配置，按需切换
gateway:
  url: "https://192.168.3.51"
  username: "l"
  password: "1"

browser:
  headless: false
  slow_mo: 300

devices:
  acurev4100:
    modbus_host: "192.168.2.242"
    modbus_port: 502
    slave_id: 1
  acuvim_iiw:
    modbus_host: "192.168.2.27"
    modbus_port: 502
    slave_id: 2
  acurev2100:
    modbus_host: "192.168.2.64"
    modbus_port: 502
    slave_id: 101
```

**在 `conftest.py` 中加载（在现有 `import config` 之前）：**

```python
import yaml
from pathlib import Path

_YAML = Path(__file__).parent / "config_env.yaml"
if _YAML.exists():
    with open(_YAML, encoding="utf-8") as f:
        _yaml_cfg = yaml.safe_load(f)
    _BASE_URL = _yaml_cfg["gateway"]["url"]
    _USERNAME = _yaml_cfg["gateway"]["username"]
    _PASSWORD = _yaml_cfg["gateway"]["password"]
    _SLOW_MO  = _yaml_cfg["browser"]["slow_mo"]
    _HEADLESS = _yaml_cfg["browser"]["headless"]
```

不同项目各自存放一份 YAML 文件（不提交到 git，加入 `.gitignore`），无需修改任何测试代码即可切换环境。

---

## 六、调试技巧

### 只跑某一条用例（按文件名）

```bash
pytest Protocols/parameter_settings/acurev4100/test_Testcase_AcuHMI_AcuRev4100_Freq_013.py -v -s
```

### 可见浏览器 + 慢速（方便观察页面操作）

```bash
$env:SLOW_MO=800; pytest Protocols/parameter_settings/acurev4100/ -v -s
```

### Modbus 失败重试

`helpers_4100.py` 中的 `modbus_read()` 内置 3 次重试 + 3s 间隔，适配 HMI 网关持续轮询导致连接槽被占用的场景。如遇频繁 Modbus 超时，可调整 `_RETRIES` 和 `_RETRY_DELAY`。

### 截图位置

失败截图自动保存到 `logs/FAIL_<用例名>.png`，同时附加到 Allure 报告。

---

## 七、新增设备

1. 在 `parameter_settings/` 下新建设备目录（如 `acuvim3/`）
2. 创建 `helpers_<设备>.py`，参照 `helpers_4100.py` 实现 `nav_to_general()` / `modbus_read()` 等
3. 逐参数新建 `test_Testcase_AcuHMI_<设备>_<参数>_<编号>.py`
4. 在 `conftest.py` 同级目录添加 `conftest.py`（Allure 标签）

---

## 八、已知问题

| 设备 | 现象 | 原因 | 状态 |
|------|------|------|------|
| AcuRev4100 | 偶发 Modbus `no response` | HMI 网关持续轮询占用连接槽 | ✅ 已修复（`helpers_4100.py` 加重试） |
| AcuvimIIW _034/_037 | VoltSwl DO/RO 下拉无 `No Output` 选项 | 疑似固件 Bug | 🔴 待确认，需提 Jira |
| AcuvimIIW _009 | RatedVoltage 690V 回读 200V | 硬件波形点位限制，设备自动回退 | 🟡 已知约束，非 Bug |
| AcuvimIIW/_046 / AcuRev2100/_009 | 长时间运行后导航超时 | 网关 session 负载积累，15s 超时不足 | 🟡 建议改 25000ms |
