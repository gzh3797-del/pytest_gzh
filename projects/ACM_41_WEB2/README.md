# ACM-41-WEB2 自动化测试

AcuRev-4100-WEB2 网关模块测试脚本目录。

算法规格：`knowledge/gateway/web2/requirements/summaries/wiring_check_v1.05.md`

---

## 目录结构

```
ACM_41_WEB2/
└── wiring_check/              接线检查（Wiring Check）自动化测试
    ├── core/                  基础设施模块
    │   ├── config.py          寄存器地址 + 连接参数 + User Channel 配置
    │   ├── meter_modbus.py    Modbus TCP 读写（192.168.2.242:502）
    │   ├── signal_driver.py   控源封装（set_ac，稳定等待 8s）
    │   ├── expected_engine.py 5种接线方式算法引擎（推导预期 Wiring Status）
    │   ├── wiring_check_page.py Playwright 页面对象（登录/触发/解析）
    │   └── report.py          HTML 报告生成
    ├── conftest.py            pytest session/module fixtures
    ├── test_3e4wy.py          3 Element 4 Wire Y（34条）
    ├── test_2e3w_delta.py     2 Element 3 Wire Delta（21条）
    ├── test_2e3w_network.py   2 Element 3 Wire Network（34条）
    ├── test_2e3w_1phase.py    2 Element 3 Wire 1 Phase（12条）
    ├── test_1e2w.py           1 Element 2 Wire（4条）
    └── reports/               HTML 报告输出（自动创建）
```

---

## 环境依赖

```bash
pip install playwright openpyxl pymodbus pytest-html
playwright install chromium
```

---

## 运行方式

### 方式一：pytest + pytest-html（推荐）

```bash
# 全部 105 条用例，生成 pytest 标准 HTML 报告
pytest test_case/ACM_41_WEB2/Wiring_check/ -v \
  --html=test_case/ACM_41_WEB2/Wiring_check/reports/pytest_report.html \
  --self-contained-html

# 单种接线方式
pytest test_case/ACM_41_WEB2/Wiring_check/test_3e4wy.py -v          # 3E4WY
pytest test_case/ACM_41_WEB2/Wiring_check/test_2e3w_delta.py -v     # 2E3W Delta
pytest test_case/ACM_41_WEB2/Wiring_check/test_2e3w_network.py -v   # 2E3W Network
pytest test_case/ACM_41_WEB2/Wiring_check/test_2e3w_1phase.py -v    # 2E3W 1Phase
pytest test_case/ACM_41_WEB2/Wiring_check/test_1e2w.py -v           # 1E2W

# 单条用例（精确 ID）
pytest test_case/ACM_41_WEB2/Wiring_check/test_3e4wy.py -k "V-13-ABC" -v

# 多条（用 or 连接）
pytest test_case/ACM_41_WEB2/Wiring_check/test_3e4wy.py -k "I-04-ABC or I-07-ABC" -v

# 按类型过滤
pytest test_case/ACM_41_WEB2/Wiring_check/ -k "V-"    # 只跑电压
pytest test_case/ACM_41_WEB2/Wiring_check/ -k "I-"    # 只跑电流
pytest test_case/ACM_41_WEB2/Wiring_check/ -k "PASS"  # 只跑基准

# 失败时立即停止
pytest test_case/ACM_41_WEB2/Wiring_check/ -x
```

> pytest-html 报告包含每条用例的 PASS/FAIL 和错误信息，适合 CI/持续集成。

### 方式二：直接运行（生成自定义 HTML 报告，含每通道明细）

```bash
# 全部接线方式一键运行（逐个执行，每种生成独立报告）
python test_case/ACM_41_WEB2/Wiring_check/run_all.py

# 单种接线方式（全部用例）
python test_case/ACM_41_WEB2/Wiring_check/test_3e4wy.py
python test_case/ACM_41_WEB2/Wiring_check/test_2e3w_delta.py
python test_case/ACM_41_WEB2/Wiring_check/test_2e3w_network.py
python test_case/ACM_41_WEB2/Wiring_check/test_2e3w_1phase.py
python test_case/ACM_41_WEB2/Wiring_check/test_1e2w.py

# 指定单条或多条用例（传 ID 参数，生成自定义 HTML 报告）
python test_case/ACM_41_WEB2/Wiring_check/test_3e4wy.py V-13-ABC
python test_case/ACM_41_WEB2/Wiring_check/test_3e4wy.py I-04-ABC I-07-ABC
python test_case/ACM_41_WEB2/Wiring_check/test_2e3w_delta.py V-05-REG V-06-REG
```

> 传入不存在的 ID 时脚本会列出所有可用 ID 并退出。

报告路径：`wiring_check/reports/wiring_<接线方式>_<时间戳>.html`

> 自定义报告含完整表格（每 User Channel 每相的预期 vs 实测），颜色编码高亮异常，适合人工核查。

### 报告方式对比

| | pytest-html | 自定义 HTML |
|-|------------|------------|
| 触发方式 | `pytest --html=...` | `python test_*.py` |
| 每通道明细 | ✗ | ✅ |
| 颜色编码 | 基础 | 详细（5种状态颜色） |
| CI 集成 | ✅ | ✗ |

---

## 报告结构（HTML）

每种接线方式生成独立报告，列结构根据接线方式动态调整：

| 固定列 | 电压列 | 电流列 |
|-------|-------|-------|
| 电表型号 / 电表IP / 接线方式 / 用例编号 / 输入信号 | 每相：预期 \| 实测 | 每 User Channel 每相：预期 \| 实测 |

**各接线方式列数：**

| 接线方式 | 电压列 | 每通道电流列 | 通道数 | 总电流列 | 备注 |
|---------|-------|------------|-------|---------|------|
| 3E4WY | A / B / C | A / B / C | 8 | 72 | |
| 2E3W Network | A / B / C | A / B / C | 12 | 108 | 每通道实际分配 AB/CA/BC 循环，比对时自动跳过未分配相 |
| 2E3W Delta | A / B / C | A / C | 12 | 48 | |
| 2E3W 1Phase | **A / C** | A / C | 12 | 48 | 无 B 相电压 |
| **1E2W** | **A** | **A** | **24** | **48** | |

**颜色编码（实测列）：**

> 实测列使用两种着色模式：

- **有预期值时**（正常比对模式）：
  - 绿色（`#d4edda`）：实测值与预期一致
  - 亮红（`#ff4444`）：实测值与预期不符（FAIL）
  - 灰色横线 `—`：无实测值（N/A 或空）

- **无预期值时**（类型回退模式）：
  - 绿色（`#d4edda`）：Pass
  - 黄色（`#fff3cd`）：Wiring Missing
  - 浅红/粉红（`#f8d7da`）：Polarity Reversed
  - 浅橙（`#fce8d5`）：Phase Shift / Phase Order Error
  - 白色：其他值

---

## 用例数量

| 接线方式 | 合计 | 说明 |
|---------|------|------|
| 3E4WY（ABC+ACB） | 34 | +3 guard 负向用例（条件2/3/4 guard：Vcn/Van/Vbn 偏低时双相缺失条件不触发）；+4 ACB 电流用例 |
| 2E3W Delta（ABC+ACB） | 21 | +4 ACB 电流用例 |
| 2E3W Network（ABC+ACB） | 34 | 同 3E4WY（共用测试用例列表） |
| 2E3W 1Phase | 12 | 删除 V-06 VC相移（v1.05 已移除该条件） |
| 1E2W | 4 | |
| **合计** | **105** | |

---

## 关键配置

| 配置项 | 位置 | 当前值 |
|-------|------|-------|
| 被测设备名（页面下拉） | `core/config.py` `WEB2_DEVICE_NAME` | `4100229` |
| Meter IP / Port / Slave | `core/config.py` | `192.168.2.29 : 502 : 101` |
| WEB2 地址 | `core/config.py` `WEB2_IP` | `192.168.3.9` |
| WEB2 登录账号 | `core/config.py` `WEB2_USER / WEB2_PASS` | `q / 1` |
| 额定电压 | `core/config.py` `NOMINAL_VOLTAGE` | `100 V` |
| 信号稳定等待 | `comm/source_control.py` `set_ac()` | `8 s` |

---

## 测试逻辑

```
① Modbus 写接线方式（0x1042）+ 相序（0x10DC）+ 额定电压（0x6501/0x651A）
② 控源 set_ac() 输出信号（等 8s 稳定）
③ expected_engine 根据控源参数 + 算法规格推导预期 Wiring Status
④ Playwright 触发页面 Wiring Check → 等完成（+1s 渲染等待）→ 解析电压表 + 电流表
⑤ 比对预期 vs 实测 → PASS / FAIL，生成 HTML 报告
```

### 特殊设计说明

**2E3W Delta 电流角度**：控源输入为绝对角（= 规格相对角 + Vab偏移），引擎内部归一化后比对。
- ABC（Vab@30°）：规格角 + 30°，例如 Pass 绝对角 Ia@0°、Ic@120°
- ACB（Vab@330°）：规格角 + 330°，例如 Pass 绝对角 Ia@0°、Ic@240°

**2E3W Network channel 分配**：12 个 User Channel 按 AB/CA/BC 循环分配，比对时自动跳过未分配相。

**报告规格参考列**：每条用例显示对应 `接线检测总表_ver1.05.xlsx` 的 Sheet 名和行号，便于追溯规格。

---

## 台架接线要求

**Delta / 1Phase 接线方式必须按以下方式连接，否则多通道会得到错误的电流角度：**

| 电流源 | 连接目标 |
|--------|---------|
| Source A | 所有 User Channel 的 Phase A CT（共 12 个） |
| Source C | 所有 User Channel 的 Phase C CT（共 12 个） |
| Source B | **不接** 任何 Delta / 1Phase CT |

> 3E4WY / Network 正常接三相即可，无特殊要求。

---

## 相关知识库

- 算法规格摘要（v1.05）：`knowledge/gateway/web2/requirements/summaries/wiring_check_v1.05.md`
- 算法规格原件：`test_case/ACM_41_WEB2/wiring_check/core/接线检测总表_ver1.05.xlsx`
- 用例摘要：`knowledge/gateway/web2/testcase/wiring_check_v1.md`
- 项目 context：`knowledge/gateway/web2/context.md`
