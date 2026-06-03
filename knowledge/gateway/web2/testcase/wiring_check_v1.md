# 接线检查用例摘要 v1（Sprint 3）

来源：接线检查AcuRev-4100 Sprint 3v1.1 20260520.xlsx（前5个 Sheet）
脚本：test_case/ACM_41_WEB2/wiring_check/
用例文件：接线检查_测试用例_WEB2_v1.xlsx（本目录）

## 覆盖范围

| 接线方式 | ABC | ACB | 合计 | 脚本 | 用例编号范围 |
|---------|-----|-----|------|------|------------|
| 3E4WY | 23 | 8 | 31 | test_3e4wy.py | TestCase_ACUREV4100WEB2_WRI_3E4_001~031 |
| 2E3W Delta | 15 | 6 | 21 | test_2e3w_delta.py | TestCase_ACUREV4100WEB2_WRI_DLT_001~021 |
| 2E3W Network | 23 | 8 | 31 | test_2e3w_network.py | TestCase_ACUREV4100WEB2_WRI_NET_001~031 |
| 2E3W 1Phase | 12 | — | 12 | test_2e3w_1phase.py | TestCase_ACUREV4100WEB2_WRI_1PH_001~012 |
| 1E2W | 4 | — | 4 | test_1e2w.py | TestCase_ACUREV4100WEB2_WRI_1E2_001~004 |
| **合计** | | | **99** | | |

## 用例 ID 命名规则（脚本内部）

- `PASS-ABC` / `PASS-ACB`：正常接线基准
- `V-01` ~ `V-13`：电压侧条件（缺失/反接/相位错误）
- `I-01` ~ `I-09`：电流侧条件（缺失/反接/相位错误）
- `-ABC` / `-ACB` 后缀：仅相位相关条件区分相序

## 用例编号对应关系

脚本 ID 与 Excel 用例编号的映射维护在：
`test_case/ACM_41_WEB2/wiring_check/core/tc_map.py`

HTML 报告包含「用例编号」列（TestCase_ACUREV4100WEB2_WRI_xxx_nnn）和「脚本ID」列，
可直接对照 Excel 用例文件回填测试结果。

## 检测类型覆盖

**电压侧（所有接线方式）**
- 接线缺失：单相/双相/三相全缺失
- 反接：Va-Vn / Vb-Vn / Vc-Vn（Delta 无此检测）
- 相位错误：Phase Shift B/C、Phase Order Error

**电流侧（各方式按通道数）**
- 接线缺失：Ian/Ibn/Icn < 0.1A
- 极性反接：PF ∈ [-1, -0.9]（3E4WY/Network/1Phase）；角度判断（Delta）
- 相位错误：PF ∈ (-0.9, 0.9]（3E4WY/Network/1Phase）；角度判断（Delta）

## 对比逻辑

```
控源输出 set_ac()
    ↓
expected_engine（算法规格逐条推导预期 Wiring Status）
    ↓
Playwright 触发页面 Wiring Check → 读取结果表
    ↓
比对：预期 vs 实测 → PASS / FAIL
```
