# testcase-analyze Skill 使用说明

## 功能简介

`testcase-analyze` 是一个 Claude Code 自定义 Skill，用于自动分析手工测试用例 Excel，
判断每条用例是否可以实现自动化，并将分析结果以颜色标注回填到 Excel 中。

| 颜色 | 含义 |
|------|------|
| 深绿色 | 已理解，可直接编写自动化脚本 |
| 红色 | 有疑问，需用户补充说明 |
| 橙色 | 半自动化，部分步骤需人工介入 |

回填列：Excel 第 13 列「需补充信息(cloude识别回填)」
用户答复列：Excel 第 14 列「用户答复(基于需补充信息,澄清信息)」

---

## 目录结构

```
auto_test_skills/testcase-analyze/
├── testcase-analyze.md          # Skill 主体内容（执行流程和判断规则）
├── testcase-analyze-readme.md   # 本说明文档
└── excel_writer.py          # Python 辅助工具（Excel 读写/染色）

.claude/commands/
└── testcase-analyze.md          # Skill 调用入口（与上面内容相同）

product_structure_testcase_regulation/
├── autotest_generativerule.md   # 通用自动化规则
├── usermanagement_struct.md     # 用户管理模块页面结构文档（已有）
└── <模块英文名>_struct.md       # 其他模块（按需补充）
```

---

## 一、分享给同事（如何部署）

### 方式 A：通过 Git（推荐）

项目已将 Skill 文件纳入 Git 管理，同事执行以下命令即可获取：

```bash
git pull
```

拉取后确认以下文件存在：
```
.claude\commands\testcase-analyze.md
auto_test_skills\testcase-analyze\excel_writer.py
```

### 方式 B：手动复制

如果不使用 Git，将以下两个路径下的内容复制到同事的项目目录：

```
# 复制 1：Skill 调用入口（必须）
.claude\commands\testcase-analyze.md

# 复制 2：Python 辅助工具（必须）
auto_test_skills\testcase-analyze\excel_writer.py
```

### 环境依赖（首次使用前执行一次）

确认 Python 已安装 openpyxl：

```bash
C:\Users\<用户名>\AppData\Local\Programs\Python\Python312\python.exe -m pip install openpyxl
```

---

## 二、使用方法

> **重要**：Skill 必须用斜杠命令触发，普通提问（"帮我分析用例"）不会启动 Skill。

### 前提条件

1. 已安装 Claude Code（桌面版或 VS Code 插件）
2. 在 Claude Code 中打开项目目录 `C:\autotest_local\autotest\gw_project_web_outo`

### 调用方式

**第一步：列出所有模块（不知道模块名时使用）**

在 Claude Code 对话框输入：
```
/testcase-analyze
```

Claude 会列出 Excel 中所有模块及用例数量，例如：
```
请指定要分析的模块：
  1. 接入设备日志管理     183 条
  2. 设备数据协议转换     149 条
  3. 用户管理             149 条
  4. 系统设置             100 条
  ...
```

**第二步：选择分析范围**

| 场景 | 命令 |
|------|------|
| 分析单个模块 | `/testcase-analyze --module 用户管理` |
| 分析多个模块 | `/testcase-analyze --module 用户管理 系统设置 About` |
| 分析全部模块 | `/testcase-analyze --all` |
| 指定 Excel 文件 | `/testcase-analyze --all --file path/to/其他.xlsx` |

Claude 会自动完成以下操作：
1. 确定目标模块列表
2. 检查各模块是否有对应的 struct.md 文档，分为两组分别处理
3. 确认 Excel 第 13、14 列（需补充信息/用户答复）存在
4. **无文档模块**：快速批量标红，写入缺少文档说明
5. **有文档模块**：逐条精细分析，标注深绿/红/橙
6. 每个模块处理完后立即写入 Excel
7. 在对话中输出各模块汇总统计表

---

## 三、页面结构文档（struct.md）说明

Skill 分析用例时依赖各模块的页面结构文档，用于判断 UI 元素定位方式和导航路径。

### 当前已有文档

| 文档文件 | 覆盖模块 |
|---------|---------|
| `usermanagement_struct.md` | 用户管理 |

### 分析无 struct.md 的模块时

若该模块尚无 struct.md，Skill 会：
1. 输出警告，说明缺少文档
2. 继续分析（不中断）
3. **将所有自动化用例标为红色**，并在回填内容中注明缺少文档依据

示例输出：
```
⚠️  警告：未找到「系统设置」的页面结构文档
分析将继续，但所有用例将标注为红色（原因：缺少文档依据）。
```

### 如何补充新模块的 struct.md

1. 参考格式：`product_structure_testcase_regulation/usermanagement_struct.md`
2. 新建文件：`product_structure_testcase_regulation/<模块英文名>_struct.md`
3. 内容需包含：
   - 导航路径（如何从首页进入该模块）
   - 页面 Tab / 子页面结构
   - 关键 UI 元素的 Playwright 定位方式
   - 表单字段名称及说明
   - 弹框/提示文案
4. 补充完成后重新调用 `/testcase-analyze --module <模块名>`，Claude 会更新颜色

---

## 四、工作流（完整步骤）

```
【准备阶段】
补充或确认各模块的 struct.md 文档（已有则跳过）
    ↓
【第1次调用】/testcase-analyze --all
    ↓
  ┌─ 无 struct.md 的模块 ───────────────────────────────────────┐
  │  快速批量标红 + 输出缺失文档列表                             │
  │  → 补充 struct.md 后用 --module 重新分析该模块              │
  └─────────────────────────────────────────────────────────────┘
  ┌─ 有 struct.md 的模块 ───────────────────────────────────────┐
  │  逐条精细分析（深绿=理解，红=疑问，橙=半自动化）             │
  └─────────────────────────────────────────────────────────────┘
    ↓
输出各模块汇总统计表
    ↓
用户查看红色行，在第14列「用户答复」填写说明
    ↓
【第2次调用】/testcase-analyze --module <有疑问的模块>
    ↓
Claude 读取用户答复，重新判断：
  - 理解了 → 第13列改为深绿色
  - 仍有疑问 → 保持红色，补充追问
    ↓
重复直到所有行变为深绿色
    ↓
调用 /auto-generate 生成自动化脚本（独立 Skill，待开发）
```

---

## 五、跨产品复用说明

Skill 的模块列表**不是硬编码的**，而是从 Excel「模块」列动态读取，因此换一个产品的 Excel 文件，模块列表自动变为该文件的内容。

### 现已支持的跨产品用法

```
/testcase-analyze --module XX模块 --file path/to/其他产品.xlsx
```

`--file` 指向目标产品的 Excel 文件，模块名即该文件「模块」列的实际内容。

### 复用的前提条件

Excel 文件必须包含以下列名（`excel_writer.py` 按列名匹配，与列顺序无关）：

| 列名 | 说明 |
|------|------|
| `模块` | 模块名称 |
| `子模块` | 子模块名称 |
| `用例编号` | 用例编号 |
| `用例标题` | 用例标题 |
| `预置条件` | 前置条件 |
| `测试步骤` | 测试步骤 |
| `预期结果` | 预期结果 |
| `用例级别` | 用例级别 |
| `用户识别半自动化` | 是否半自动化（"是"/空） |
| `用户识别自动化` | 是否做自动化（"是"/空） |
| `需补充信息(cloude识别回填)` | Claude 回填分析结果（第 13 列） |
| `用户答复(基于需补充信息,澄清信息)` | 用户澄清内容（第 14 列） |

若其他产品 Excel 列名不同，`excel_writer.py` 将读不到数据（静默返回空）。

### 两种应对策略

**A. 统一模板（推荐）**：团队约定所有产品使用相同的 Excel 列名，`--file` 指向各自文件即可，无需任何改动。

**B. 列名映射（按需扩展）**：如果现有 Excel 格式无法统一，可为 `excel_writer.py` 增加 `--col-map` 参数支持列名自定义，但调用会更复杂。

---

## 六、注意事项（本项目）

- 只写入第 13、14 列，**不修改**原始用例数据（第 1-12 列）
- 已标注深绿色的行，二次运行时**不会被覆盖**（幂等保护）
- 「用户识别自动化」列（第10列）值不为"是"的用例**跳过不分析**
- Excel 文件路径含中文时，命令行中需用双引号包裹路径
- 分析结果依赖 struct.md 文档质量；文档越详细，深绿色判断越准确

---

## 七、常见问题

**Q：输入 `/testcase-analyze` 没有反应？**
A：确认 `.claude/commands/testcase-analyze.md` 文件存在，且当前工作目录是项目根目录。

**Q：Excel 没有被写入？**
A：运行以下命令确认 excel_writer.py 可以正常工作：
```bash
python auto_test_skills/testcase-analyze/excel_writer.py --list-modules "Manual_testcase/【AcuHMI-1-7】测试用例_foAI.xlsx"
```
若报错，检查 openpyxl 是否已安装。

**Q：分析结果全部是红色？**
A：两种可能：
- 该模块没有对应的 struct.md（输出中会有 ⚠️ 警告）→ 按第三节说明补充文档后重新调用
- struct.md 存在但该模块用例描述模糊 → 在第14列填写用户答复，重新调用

**Q：模块名包含空格或特殊字符怎么办？**
A：用引号包裹模块名：
```
/testcase-analyze --module "接入设备日志管理"
```

**Q：新模块如何快速开始？**
A：直接运行 `/testcase-analyze --all`，无 struct.md 的模块会自动标红并列出需补充的文档清单，参考 `usermanagement_struct.md` 补充后，用 `--module <模块名>` 重新分析即可更新颜色。

**Q：`--all` 和逐个 `--module` 有什么区别？**
A：效果相同，区别在于效率。`--all` 一次性处理所有模块，输出汇总统计表；`--module` 适合只关注特定模块或做二次澄清时使用。

---

## 八、相关文件

| 文件 | 说明 |
|------|------|
| `product_structure_testcase_regulation/autotest_generativerule.md` | 自动化用例生成规则 |
| `product_structure_testcase_regulation/usermanagement_struct.md` | 用户管理模块页面结构文档 |
| `product_structure_testcase_regulation/*_struct.md` | 其他模块页面结构文档（按需补充） |
| `Manual_testcase/【AcuHMI-1-7】测试用例_foAI.xlsx` | 手工测试用例（分析对象） |
