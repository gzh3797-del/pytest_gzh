# analyze-web-page 技能

## 功能
使用 Playwright MCP 穷尽式探索一个站点，**为每个可路由子页产出一份「页面上下文」文档 `<Prefix_SubPage>_context.md`**，落到项目知识库 `requirements/context/`（约定 #11「选择器沉淀」），供 `webtestcase_manual_to_auto` 把手工用例转自动化、以及后续 UI 用例直接复用选择器。

> 产物契约、命名、格式全部对齐仓库现有 `_context.md`（样例：`knowledge/gateway/hmi1-7/requirements/context/`）。**不再**产出旧的整页 10 章大报告或 generate-pytest 数据。

## 原理
Playwright MCP 提供浏览器操作工具集（`browser_navigate`、`browser_snapshot` 等），Claude 直接调用获取页面信息，无需中间脚本（唯一脚本 `audit_snapshots.py` 只做门禁的文件计数）。

## 用法
```
/analyze-web-page <URL> <项目知识库根> [登录凭据]
```
- `<项目知识库根>` = 项目一览中该项目 `context.md` 所在目录（如 `knowledge/gateway/hmi1-7`）；交付物写入其下 `requirements/context/`。
- `[登录凭据]` 可选（账号/密码/自签证书提示），缺省时运行时询问。
- 开工前先读该目录 `_INDEX_context.md`（若存在），本次为**增量合并**而非覆盖。

示例：
```
/analyze-web-page https://192.168.3.71 knowledge/gateway/hmi1-7 admin/Admin@080066
```

## 产物
| 路径 | 内容 |
|------|------|
| `<项目知识库根>/requirements/context/<Prefix_SubPage>_context.md` | 每个可路由子页一份页面上下文（交付物） |
| `<项目知识库根>/requirements/context/_INDEX_context.md` | 页面上下文索引（增量维护） |
| `.analyze_scratch/`（工作区临时目录） | 阶段1–4 中间快照，**任务末清理零残留**，不进 git |

## 6 阶段流程（含审计门禁）

| 阶段 | 操作 | 门禁 |
|------|------|------|
| ① 初始扫描 | `browser_navigate` + `browser_snapshot` + 采集 API/框架（**JS DOM 计数区分 Ant/El Plus**）→ `expected_ops.json` | — |
| ② 穷尽探索 | 展开/切换/触发/遍历 + **每个可路由子页跟随** + **SPA 导航实测（路由守卫/同路由 goto）** | 原子操作规则 |
| ③ 表单深度 | 每项下拉/radio/checkbox双态/级联/搜索/提交/重置 + 校验时机·**依赖字段多状态**·跨页传播·结果反馈·高危实测 | 缺省否定清单 |
| ③.4 反转质询 | 自我批评 → ❌ 回退阶段2 | 零成本门禁 |
| ③.5 独立审计 | 独立 Agent adversarial + `audit_snapshots.py` 计数 | FAIL 回退阶段2 |
| ④ 最终确认 | 全页快照对比 + 快照计数 + 页面类型判定 | 文件系统计数 |
| ⑤ 沉淀上下文 | 逐子页产出 `_context.md` → `requirements/context/` + 维护 `_INDEX_context.md` + 清理临时目录 | 产物计数对齐 + §六 质量自检 |

## 依赖
- Playwright MCP（已在项目配置中启用）

## 文件说明
| 文件 | 类型 | 说明 |
|------|------|------|
| `SKILL.md` | 指令规范 | **核心**。执行流程、门禁、`_context.md` 产物契约，AI 直接按此执行 |
| `EXHAUSTIVE.md` | 探索标准 | 穷尽式页面探索 9 大维度 |
| `PITFALLS.md` | 坑清单 | 五类实战工程坑（SPA 路由守卫 / 同路由 goto / Ant·El Plus JS 计数区分 / 依赖字段多状态 / 跨页传播·表格行按钮），含检测方法与 `_context.md` 记录字段 |
| `OUTPUT_STANDARD.md` | 输出标准 | `_context.md` / `_INDEX_context.md` 命名、格式、字段、`getByRole` 定位规范 |
| `scripts/audit_snapshots.py` | 门禁脚本 | 跨平台统计中间快照数并与 `expected_ops.json` 逐维度对比；可选 `ctx_dir` 参数信息性报告已交付 `_context.md` 数 |
