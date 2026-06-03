# AcuHMI-1-7 自动化测试

AcuHMI-1-7 网关测试脚本目录。

算法规格：`knowledge/gateway/web2/requirements/summaries/wiring_check_v1.05.md`（与 WEB2 共用同一套算法规格）

---

## 目录结构

```
AcuHMI_1_7/
└── wiring_check/              接线检查（Wiring Check）自动化测试
    ├── core/                  基础设施模块
    │   ├── config.py          寄存器地址 + 连接参数 + User Channel 配置
    │   ├── meter_modbus.py    Modbus TCP 读写（4100 电表）
    │   ├── signal_driver.py   控源封装（set_ac）
    │   ├── expected_engine.py 5种接线方式算法引擎（推导预期 Wiring Status）
    │   ├── wiring_check_page.py Playwright 页面对象（登录/导航/触发/解析）
    │   └── report.py          HTML 报告生成
    ├── conftest.py            pytest session/module fixtures
    ├── test_3e4wy.py          3 Element 4 Wire Y（31条）
    ├── test_2e3w_delta.py     2 Element 3 Wire Delta（21条）
    ├── test_2e3w_network.py   2 Element 3 Wire Network（31条）
    ├── test_2e3w_1phase.py    2 Element 3 Wire 1 Phase（12条）
    ├── test_1e2w.py           1 Element 2 Wire（4条）
    ├── debug_page.py          页面结构调试脚本
    └── reports/               HTML 报告输出（自动创建）
```

---

## 与 WEB2 脚本的主要差异

| 差异点 | WEB2 | HMI1-7 |
|-------|------|--------|
| 网关 IP | `192.168.3.9` | `192.168.2.8` |
| 页面导航 | 直接 goto URL | Settings 侧边栏 → Diagnostics → Wiring Check tab |
| Confirm 弹窗 | 有 Save 按钮（下发额定电压到表） | **无 Save 按钮**（额定电压只用于计算） |
| 设备选择下拉 | 有（WEB2 下挂多台 4100） | 有（默认 All，可选单台） |
| 结果表结构 | — | **与 WEB2 完全相同** |
| 算法规格 | — | **与 WEB2 完全相同** |

---

## 环境依赖

```bash
pip install playwright openpyxl pymodbus pytest-html
playwright install chromium
```

---

## 运行前配置

编辑 `core/config.py`：

| 配置项 | 说明 |
|-------|------|
| `HMI_IP` | HMI 地址，默认 `192.168.2.8` |
| `HMI_USER / HMI_PASS` | 登录账号，默认 `q / 1` |
| `HMI_DEVICE_NAME` | HMI Device 下拉中显示的被测 4100 设备名（如 `'4100242'`）；留空则选 All |
| `METER_TCP_IP` | 4100 Modbus TCP IP，默认 `192.168.2.29` |
| `MODBUS_SLAVE` | 4100 Slave ID，默认 `102` |
| `NOMINAL_VOLTAGE` | 额定电压（V），默认 `100` |

---

## 运行方式

### 方式一：pytest + pytest-html（推荐）

```bash
# 全部 99 条用例
pytest test_case/AcuHMI_1_7/wiring_check/ -v \
  --html=test_case/AcuHMI_1_7/wiring_check/reports/pytest_report.html \
  --self-contained-html

# 单种接线方式
pytest test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py -v
pytest test_case/AcuHMI_1_7/wiring_check/test_2e3w_delta.py -v
pytest test_case/AcuHMI_1_7/wiring_check/test_2e3w_network.py -v
pytest test_case/AcuHMI_1_7/wiring_check/test_2e3w_1phase.py -v
pytest test_case/AcuHMI_1_7/wiring_check/test_1e2w.py -v

# 单条用例
pytest test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py -k "V-13-ABC" -v

# 按类型过滤
pytest test_case/AcuHMI_1_7/wiring_check/ -k "V-"    # 只跑电压
pytest test_case/AcuHMI_1_7/wiring_check/ -k "I-"    # 只跑电流
pytest test_case/AcuHMI_1_7/wiring_check/ -k "PASS"  # 只跑基准

# 失败时立即停止
pytest test_case/AcuHMI_1_7/wiring_check/ -x
```

### 方式二：直接运行（生成自定义 HTML 报告，含每通道明细）

```bash
# 全部接线方式一键运行
python test_case/AcuHMI_1_7/wiring_check/run_all.py

# 单种接线方式
python test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py
python test_case/AcuHMI_1_7/wiring_check/test_2e3w_delta.py
python test_case/AcuHMI_1_7/wiring_check/test_2e3w_network.py
python test_case/AcuHMI_1_7/wiring_check/test_2e3w_1phase.py
python test_case/AcuHMI_1_7/wiring_check/test_1e2w.py

# 指定单条或多条用例（传 ID 参数）
python test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py V-13-ABC
python test_case/AcuHMI_1_7/wiring_check/test_3e4wy.py I-04-ABC I-07-ABC
```

报告路径：`wiring_check/reports/wiring_<接线方式>_<时间戳>.html`

---

## 用例数量

| 接线方式 | 合计 | 说明 |
|---------|------|------|
| 3E4WY（ABC+ACB） | 34 | +3 guard 负向用例（条件2/3/4 guard）；+4 ACB 电流用例 |
| 2E3W Delta（ABC+ACB） | 23 | +4 ACB 电流用例；+2 防回归用例（V-05/V-06 REG） |
| 2E3W Network（ABC+ACB） | 34 | 同 3E4WY（共用测试用例列表） |
| 2E3W 1Phase | 12 | |
| 1E2W | 4 | |
| **合计** | **107** | |

---

## 测试逻辑

```
① Modbus 写接线方式（0x1042）+ 相序（0x10DC）+ 额定电压（0x6501/0x651A）至 4100
② 控源 set_ac() 输出信号（等 8s 稳定）
③ expected_engine 根据控源参数 + 算法规格推导预期 Wiring Status
④ Playwright：Settings → Diagnostics → Wiring Check tab → 触发检查 → 等完成 → 解析电压表 + 电流表
⑤ 比对预期 vs 实测 → PASS / FAIL，生成 HTML 报告
```

### HMI 页面导航说明

`wiring_check_page.navigate()` 每次通过菜单三步导航（直接 goto URL 会被 SPA 路由重定向）：

1. `goto /#/systemSettings/dateTime` — 激活 Settings 侧边栏
2. 点击侧边栏 **Diagnostics**（精确文字匹配）
3. 点击页内 **Wiring Check** tab

整个测试会话只导航一次（session scope），后续用例复用同一页面实例。

---

## 关键配置

| 配置项 | 位置 | 当前值 |
|-------|------|-------|
| 被测设备名（页面下拉） | `core/config.py` `HMI_DEVICE_NAME` | `''`（需填写） |
| Meter IP / Port / Slave | `core/config.py` | `192.168.2.29 : 502 : 102` |
| HMI 地址 | `core/config.py` `HMI_IP` | `192.168.2.8` |
| HMI 登录账号 | `core/config.py` `HMI_USER / HMI_PASS` | `q / 1` |
| 额定电压 | `core/config.py` `NOMINAL_VOLTAGE` | `100 V` |
| 信号稳定等待 | `comm/source_control.py` `set_ac()` | `8 s` |

---

## 台架接线要求

**Delta / 1Phase 接线方式须按以下方式连接，否则多通道会得到错误的电流角度：**

| 电流源 | 连接目标 |
|--------|---------|
| Source A | 所有 User Channel 的 Phase A CT |
| Source C | 所有 User Channel 的 Phase C CT |
| Source B | **不接** 任何 Delta / 1Phase CT |

> 3E4WY / Network 正常接三相即可，无特殊要求。

---

## 相关知识库

- 算法规格摘要（v1.05）：`knowledge/gateway/web2/requirements/summaries/wiring_check_v1.05.md`
- 项目 context：`knowledge/gateway/hmi1-7/context.md`
