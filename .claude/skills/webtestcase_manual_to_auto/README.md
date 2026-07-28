# webtestcase_manual_to_auto — 使用说明

Web UI 手工用例（Excel）→ Playwright + pytest 自动化的**端到端**单一 skill。
由旧 `webtestcase-analyze_debug` + `webtestcase_manual_to_auto` 合并而成。

## 一图流程

```
① 分析（绿/红记根因/橙，严禁编造） → ② 人工澄清 → ③ 批量生成（先出约束表）
   → ④ 覆盖门禁自检（可追溯+步骤/预期零漏项） → ⑤ pytest 运行 → ⑥ 自动调试(≤3轮)
   → ⑦ 回填结果列 → ⑧ 增量沉淀 context
```

## 目录

```
.claude/skills/webtestcase_manual_to_auto/
├── SKILL.md            # 执行流程（权威，Claude 按此执行）
├── README.md           # 本文件
├── excel_writer.py     # Excel 读写（真实 schema + 3 个 Claude 列）
├── GENERATION_RULES.md # 脚本生成规范（import/fixture/定位/Element Plus）
├── COVERAGE_GATE.md    # 生成后覆盖门禁（约束表可追溯 + 步骤/预期零漏项自检）
└── DEBUG_PATTERNS.md   # 自动调试错误模式 → 修复策略
```

## 快速开始

```
# 看各模块分析状态
/webtestcase_manual_to_auto

# 分析 + 生成 + 调试 + 回填（单模块）
/webtestcase_manual_to_auto --module 用户管理

# 只分析染色，人工澄清后再生成
/webtestcase_manual_to_auto --module 系统设置 --analyze-only

# 单条闭环
/webtestcase_manual_to_auto --case TestCase_AcuHMI_009_08_case01

# 换用例集 / 换项目
/webtestcase_manual_to_auto --all --file <其他.xlsx> --project <项目名>
```

## 关键参数

| 参数 | 说明 | 默认 |
|------|------|------|
| `--file` | 用例集 Excel | `knowledge/gateway/hmi1-7/testcase/【AcuHMI-1-7】项目软件全功能测试用例2026060701.xlsx` |
| `--project` | 目标项目（项目映射） | `AcuHMI_1_7` |
| `--module` / `--all` / `--case` | 处理范围 | — |
| `--analyze-only` | 只分析染色 | 否 |
| `--debug-only` | 跳过生成，跑已有脚本 | 否 |
| `--overwrite` | 覆盖已存在脚本 | 否 |

## 路径规则

- 目标脚本落盘：`projects/<project>/tests/<Excel「自动化类型」列>/test_<用例编号>.py`
- 页面 context 目录：`<knowledge_root>/requirements/context/`（先读 `_INDEX_context.md`）
- 过滤条件：仅处理 `自动化(是/否)="是"` 的行

## Excel 写回

只**追加** 3 个 Claude 列，不动任何原始列：
- `需补充信息(claude识别回填)`（绿/红/橙 + 字体色）
- `用户答复(澄清信息)`（用户填，二次运行读）
- `自动化脚本调试结果`（`是` / `跳过|原因` / `调试失败|原因`）

## 环境依赖

```
python -m pip install openpyxl
```
Playwright/pytest 依赖沿用各项目 `projects/<项目>/`（`conftest.py`、`pages/login_page.py`、`settings.py`）。

## 常见问题

**Q：分析全红？** context 目录缺该页面沉淀 → skill 会调 `analyze-web-page` 补探；或用例描述模糊 → 在「用户答复」列澄清后二次运行。

**Q：会不会编造选择器？** 不会。无 context 依据的步骤一律标红记根因；生成后 Gate A 可追溯性门禁再兜底——每个 locator/断言值都必须回溯到探查约束表，回溯不到即判失败（COVERAGE_GATE.md）。

**Q：脚本 pytest 跑过了就算数吗？** 不算。Step 4.5 覆盖门禁额外校验：用例每条步骤都有对应操作（Gate B 零漏项）、每条预期都有对应断言（Gate C 零漏项，`A or B or C`/只验第一行等弱断言判 partial）。跑得过但漏断言的用例会被门禁拦下补齐或转红。

**Q：会改我的用例集吗？** 只追加 3 个 Claude 列并写这 3 列；`用例编号` 等原始列逐字符不动（约定 #12）。

**Q：换个产品用？** 在 SKILL.md「项目映射」表加一行（code_root / knowledge_root），`--project` 指定即可。

**Q：缺页探查会残留脚本吗？** 不会。探查全走 `analyze-web-page` skill，任务结束零残留（CLAUDE.md 最高优先级禁止项）。
