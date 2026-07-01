# 接线检查（Wiring Check）自动化用例执行说明

AcuHMI-1-7 接线检查测试，覆盖 5 种接线方式：`3E4WY`、`2E3W Delta`、`2E3W Network`、`2E3W 1Phase`、`1E2W`。
对比逻辑：预期由 `core/expected_engine.py` 推导，实测由 Playwright 读 HMI 页面 Wiring Status。

> 运行前提：信号源、被测表 Modbus TCP、HMI Web 页面均可达；环境参数在 `projects/AcuHMI_1_7/settings.py`。
> 所有命令从**仓库根目录**（`autotest/`）执行；`-s` 用于实时输出推导/实测日志（浏览器非 headless）。

| 接线方式 | 测试文件 |
|---------|---------|
| 3E4WY | `test_3e4wy.py` |
| 2E3W Delta | `test_2e3w_delta.py` |
| 2E3W Network | `test_2e3w_network.py` |
| 2E3W 1Phase | `test_2e3w_1phase.py` |
| 1E2W | `test_1e2w.py` |

---

## 全部执行

```bash
pytest projects/AcuHMI_1_7/tests/Wiring_check/ -v -s
```

## 分接线方式执行

```bash
pytest projects/AcuHMI_1_7/tests/Wiring_check/test_3e4wy.py        -v -s   # 3E4WY
pytest projects/AcuHMI_1_7/tests/Wiring_check/test_2e3w_delta.py   -v -s   # 2E3W Delta
pytest projects/AcuHMI_1_7/tests/Wiring_check/test_2e3w_network.py -v -s   # 2E3W Network
pytest projects/AcuHMI_1_7/tests/Wiring_check/test_2e3w_1phase.py  -v -s   # 2E3W 1Phase
pytest projects/AcuHMI_1_7/tests/Wiring_check/test_1e2w.py         -v -s   # 1E2W
```

## 单条执行（`-k` 按用例 ID 过滤）

节点 ID 形如 `用例编号-脚本ID`（如 `TestCase_AcuHMI-1-7_034_001_002-V-01-ABC`），`-k` 匹配脚本 ID 即可：

```bash
# 单条
pytest projects/AcuHMI_1_7/tests/Wiring_check/test_3e4wy.py -k "V-01-ABC" -s

# 多条（或关系）
pytest projects/AcuHMI_1_7/tests/Wiring_check/test_3e4wy.py -k "V-01-ABC or I-04-ABC" -s
```

各文件可用的脚本 ID：

| 文件 | 脚本 ID |
|------|--------|
| `test_3e4wy.py` / `test_2e3w_network.py` | `PASS-ABC`、`V-01-ABC`…`V-13-ABC`、`I-01-ABC`…`I-09-ABC`；ACB 相序：`PASS-ACB`、`V-11-ACB`、`V-12-ACB`、`V-13-ACB`、`I-05-ACB`、`I-06-ACB`、`I-08-ACB`、`I-09-ACB` |
| `test_2e3w_delta.py` / `test_2e3w_1phase.py` | `PASS-ABC`、`V-01`、`V-02`…、`I-01`… |
| `test_1e2w.py` | `PASS`、`V-01`、`I-01`、`I-02` |

## 按类别批量执行（`-k` 关键字）

```bash
pytest .../test_3e4wy.py -k "V-"   -s   # 所有电压侧用例
pytest .../test_3e4wy.py -k "I-"   -s   # 所有电流侧用例
pytest .../test_3e4wy.py -k "ACB"  -s   # 所有 ACB 相序用例
pytest .../test_3e4wy.py -k "PASS" -s   # Pass 基准用例
```

## 列出全部用例（不执行）

```bash
pytest projects/AcuHMI_1_7/tests/Wiring_check/test_3e4wy.py --collect-only -q
```

---

> 脚本直跑方式（`python test_xxx.py [ID...]`，生成 `reports/*.html`）保留可用，仅作调试备选。

---

## 被测表连接参数来源

`config.yaml` 的 `meter_device_name` 仍用于**选择**测哪台表（同时是 HMI 下拉显示名）。
该表的 Modbus 连接参数（IP/端口/Modbus ID）在测试运行时由 `meter_connection` fixture 从网关
Physical Devices REST API 动态发现，不再由 `config.yaml` 维护。独立运行（`python test_*.py`）
时由 `cfg.ensure_meter_connection()` 自动完成发现。目标表未在网关在线设备中时立即报错。
