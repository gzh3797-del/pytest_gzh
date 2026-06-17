# AcuHMI-1-7 UI 自动化测试 — 执行命令索引

Playwright + pytest 的 Web UI 测试。**所有命令在仓库根目录 `autotest/` 下执行**（用例用绝对导入 `from projects.acuhmi_1_7...`，不能在子目录里跑），用 `pytest` 启动器（不要用 `python -m pytest`，避免 IDE 解释器路径问题）。

## 通用说明

- **目录结构**：`projects/acuhmi_1_7/` 下 `settings.py`（常量）/ `config.yaml`（非敏感配置）/ `conftest.py`（fixture + 失败截图）/ `pages` / `helpers` / `tests/ui/<子模块>/`。
- **配置来源**：账号密码、SMTP 等机密在仓库根 `configs/.env`（已 gitignore，模板见 `configs/.env.example`）；非敏感项在 `config.yaml`。
- **测试报告**：`pytest.ini` 已配 `--html`，运行后自动生成到 `reports/acuhmi_1_7/<子模块>/<时间戳>/report.html`（命令里无需再写 `--html`；按 `tests/` 下子模块名分目录，失败截图在同级 `screenshots/`）。
- **常用命令**：

```bash
# 全部 UI 用例
pytest projects/AcuHMI_1_7/tests/ui/ -v

# 登录冒烟
pytest projects/AcuHMI_1_7/tests/ui/test_login.py -v

# 显示浏览器窗口（调试）
pytest projects/AcuHMI_1_7/tests/ui/<子模块>/ -v --headed

# 只收集不执行
pytest projects/AcuHMI_1_7/tests/ui/<子模块>/ --collect-only -q
```

## ⚠️ 高危操作（重启 / 恢复出厂）—— 运行前务必确认

部分用例会**重启被测设备**（恢复出厂同样触发重启），单条耗时约 2-3 分钟、会中断设备服务。**运行前先确认设备可用、无人占用**。已知高危用例集中在系统设置（见下方该子模块小节）。恢复出厂类已带 `@pytest.mark.destructive`，可用 `-m "not destructive"` 排除。

## 子模块索引

| 子模块 | 目录 | 用例(文件)数 | 对应测试用例 / 编号 |
|--------|------|:-----------:|---------------------|
| 系统设置 | `systemsettings/` | 58 | 005_*（+Remote Access） |
| 用户管理 | `usermanagement/` | 147 | 007_* |
| 设备数据协议转换 | `protocols/` | 110 | BACnet/MQTT/SNMP/EtherNetIP… |
| 接入设备日志管理 | `datalog/` | 39 | DataLog |
| 系统诊断 | `diagnostics/` | 27 | — |
| 模板管理 | `templates/` | 20 | — |
| 接入设备参数设置 / 设备管理 | `physicaldevices/` | 18 | — |
| About | `about/` | 11 | 009_* |
| Virtual Device | `virtualdevice/` | 11 | — |
| 安全性测试 | `security/` | 6 | — |
| Web Devices | `webdevices/` | 5 | — |
| Maintenance | `maintenance/` | 4 | — |
| 兼容性测试 | `compatibility/` | 1 | — |
| UI 界面测试 | `ui/` | 1 | — |
| 登录 | `test_login.py` | 1 | — |

---

# 各子模块执行命令

> 约定：每个子模块至少给「全跑」命令；需要细分时用 `-k "<编号片段>"` 按文件名过滤。
> 各子模块负责人可在自己小节补充常用的细分命令。

## 系统设置 systemsettings（58）

```bash
# ── 全部 ──
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/ -v
# 排除恢复出厂用例（避免重启设备）
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/ -m "not destructive" -v

# ── 按子模块（-k 按文件名通配；注释为 已落地/用例总数）──
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/general -k "005_01" -v    # Date & Time               10/11
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/general -k "005_02" -v    # Network                    2/8
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/general -k "005_04" -v    # Email                     26/29
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/general -k "005_05" -v    # Alarm Notification         7/11
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/general -k "005_06" -v    # Certificate/FactoryReset   9/19
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/general -k "005_07" -v    # Whitelist                  3/10         
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/general -k "009_006" -v   # Remote Access              1/3
# 合计 58/99
# 注：Email 发信类用例（005_04_case01/02/03、005_05 Test Email 的 case02_01/03/03_1/03_2）
#     依赖真实可用 SMTP 服务器，已不做自动化（手动验证）；005_04/005_05 仅保留"配置保存"类用例。
# 注：005_08 配置导入类（case01_01 非法文件、case01_07_1 跨版本导入）均不做自动化——真机确认
#     系统对任意 .an 文件 Import 都先弹"will reboot device"确认框、不先校验内容/版本（页面
#     Caution 仅展示文字、不拦截），点 OK 会重启设备，无法安全自动化；改人工验证。

# ── 单用例（把 <编号> 换成具体用例，如 005_01_case01）──
pytest projects/AcuHMI_1_7/tests/ui/systemsettings/general/test_TestCase_AcuHMI_<编号>.py -v -s

# ⚠️ 高危（会重启设备，跑前确认）：005_01_case01（重启）、005_06_case06/07/08/09/11/12/13/16（恢复出厂，case09 跑两次）
```

## 用户管理 usermanagement（147）

```bash
pytest projects/AcuHMI_1_7/tests/ui/usermanagement/ -v
# 细分示例（按需）：pytest projects/AcuHMI_1_7/tests/ui/usermanagement/ -k "007_01" -v
```

## 设备数据协议转换 protocols（110）

```bash
pytest projects/AcuHMI_1_7/tests/ui/protocols/ -v
```

## 接入设备日志管理 datalog（39）

```bash
pytest projects/AcuHMI_1_7/tests/ui/datalog/ -v
```

## 系统诊断 diagnostics（27）

```bash
pytest projects/AcuHMI_1_7/tests/ui/diagnostics/ -v
```

## 模板管理 templates（20）

```bash
pytest projects/AcuHMI_1_7/tests/ui/templates/ -v
```

## 接入设备参数设置 / 设备管理 physicaldevices（18）

```bash
pytest projects/AcuHMI_1_7/tests/ui/physicaldevices/ -v
```

## About about（11）

```bash
pytest projects/AcuHMI_1_7/tests/ui/about/ -v
```

## Virtual Device virtualdevice（11）

```bash
pytest projects/AcuHMI_1_7/tests/ui/virtualdevice/ -v
```

## 安全性测试 security（6）

```bash
pytest projects/AcuHMI_1_7/tests/ui/security/ -v
```

## Web Devices webdevices（5）

```bash
pytest projects/AcuHMI_1_7/tests/ui/webdevices/ -v
```

## Maintenance maintenance（4）

```bash
pytest projects/AcuHMI_1_7/tests/ui/maintenance/ -v
```

## 兼容性测试 compatibility（1）

```bash
pytest projects/AcuHMI_1_7/tests/ui/compatibility/ -v
```

## UI 界面测试 ui（1）

```bash
pytest projects/AcuHMI_1_7/tests/ui/ui/ -v
```

## 登录 test_login（1）

```bash
pytest projects/AcuHMI_1_7/tests/ui/test_login.py -v
```

---

## 问题排查（QA）

- **命令秒退、无输出**：`python` 指向了 Windows 应用商店占位符。用 `pytest` 启动器，或确认 IDE/终端的解释器指向当前虚拟环境（不是失效的旧 `.venv` 路径）。
- **报告没生成**：确认在仓库根目录执行；`pytest.ini` 的 `addopts` 已含 `--html`，报告在 `reports/acuhmi_1_7/<子模块>/<时间戳>/`。
- **自定义 marker**：`destructive`（改设备状态/恢复出厂）、`slow`（重启/同步等耗时）已在根 `pytest.ini` 注册，可用 `-m "not destructive"` 过滤。
