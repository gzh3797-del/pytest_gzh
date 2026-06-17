---
name: strategy-design
description: |
  将功能需求转化为测试策略脑图并导出为 XMind 文件（兼容 XMind 26）。
  适用于任意项目、任意技术栈的测试策略输出任务。

  当用户提到以下任意情况时，必须调用本 skill，不得跳过：
  - 要求输出测试点、测试策略、测试脑图、XMind
  - 要求分析功能需求并生成测试覆盖
  - 要求补充测试点或整理测试维度
  - 要求将需求文档转化为测试策略
  - 任何"帮我写测试点""生成脑图""导出 xmind"类请求

  本 skill 包含完整的：
  1. 测试策略内容规范（分组顺序、叶子节点句式、颗粒度、语言风格）
  2. XMind 文件生成方法（兼容 XMind 26 的 XML 格式，逻辑图布局）

  不得跳过本 skill 直接输出测试点或生成文件。
---

# XMind 测试策略 Skill

## 概述

本 skill 将功能需求转化为规范的测试策略脑图，并生成可直接用 XMind 26 打开的 `.xmind` 文件。

**适用场景：** 任意项目的协议类、配置类、UI 类、API 类功能测试策略输出。

**执行顺序（每次必须按序执行）：**
1. 阅读本文档全部内容规范（第零步至第六步）
2. 阅读共享方法学文档 `knowledge/shared/test_design_methods.md`（方法分类、风险扫描、领域场景库），用于第三步的方法选择与第四步的补充
3. 按下方「生成前必做：读取当前项目知识库背景」章节确定目标项目，读取其 `context.md` 并扫 `requirements/summaries/` 清单
4. 阅读下方「句式模板与语言规范」章节
5. 按规范起草测试点，补充需求未覆盖的场景
6. 阅读下方「XMind 文件生成规范」章节，生成 XMind 文件并提供下载
7. **生成文件后自动调用 `coverage-check` skill 做查漏复核**：传入「原始需求文档路径 + 刚生成的 XMind 产物 + 模式标识`脑图`」，由其独立重导并多基准回扫，在对话里给三档分级查漏结论；本步不生成文件、不改动脑图，漏点是否回炉由工程师决定

---

## 分析隔离原则（每次必须遵守）

**每次处理新模块时，必须重新读取当前文档原始内容，严禁携带历史分析信息。**

- 每次处理特定模块的需求或测试点前，**必须从当前文档/文件重新读取该模块的原始内容**，不得依赖对话记忆中的旧版内容
- **严禁**将本次对话中对其他模块的分析结论带入当前任务
- **严禁**引用历史对话中对同一功能的旧版分析（需求文档可能已更新）
- 若当前读取内容与历史记忆存在任何不一致，**以当前读取内容为唯一依据**，不做融合
- **知识库是补充背景，不是权威来源**：项目 `context.md` / `requirements/summaries/` 仅用于术语与模块/设备范围对齐、以及补充跨模块联动等场景；待测功能仍以会话喂入的当前需求文档为唯一权威依据，**知识库与当前文档冲突时一律以当前文档为准**（知识库可能过时）

---

## 生成前必做：读取当前项目知识库背景

起草测试点前，先定位当前项目并读取其知识库背景，避免漏掉跨模块联动、设备兼容等需要项目全局视野才能想到的场景。

**1. 确定目标项目**
- 从需求文档路径/内容、输出路径或会话上下文推断当前项目，对照 `CLAUDE.md`「项目一览」表校对
- 推断不出或有歧义时，**反问用户确认，不臆测**

**2. 读取 `context.md`**
- 路径取自「项目一览」表的「详情」列（如 `knowledge/gateway/web2/context.md`、`knowledge/cloud/context.md`）
- 用途：掌握现有模块结构、项目约定、下挂设备清单，保证模块命名/术语/设备范围与项目一致

**3. 扫 `requirements/summaries/` 清单，按需读相关摘要**
- 列出该项目 `requirements/summaries/` 目录下的文件清单（无 INDEX，按文件名判断相关性）
- 挑出与当前功能相关的摘要按需读取，作为补充测试点（跨模块联动、功能共存、设备兼容）的来源；**不全读**

**4. 兜底**
- 该项目无 `context.md`，或 `requirements/summaries/` 为空/不存在时，跳过并一句话说明，照常用会话喂入的文档生成

> 知识库的角色见上方「分析隔离原则」最后一条：仅作背景补充，与当前文档冲突时以当前文档为准。

---

## 第零步：功能概述节点

在根节点下、所有测试分组之前，插入一个**固定子节点**，标题为 `功能概述`，作为背景说明而非测试点。

**子节点内容（各写一条叶子节点）：**

| 子节点标题 | 内容要求 |
|------------|----------|
| 功能定位 | 一句话说明该功能是什么、解决什么问题，面向评审者 |
| 工作流程 | 数据/操作从哪里来 → 经过什么处理 → 到哪里去，突出链路 |
| 关键约束 | 前提条件、协议依赖、只读限制、硬件要求等不可忽略的边界 |

**写法规则：**
- 用陈述句描述事实，不写测试动作
- 术语保留英文原始写法
- 若需求文档未提供某项信息，写 `（需确认）` 占位，不得省略该子节点
- 功能概述节点**不计入用例数量预估**

**XMind 生成方式：**（`group` / `leaves` 辅助函数见下方「XMind 文件生成规范」生成模板）
```python
overview = group("功能概述", leaves(
    "功能定位：<一句话描述>",
    "工作流程：<链路描述，如 A → B → C>",
    "关键约束：<约束条件，多条用分号分隔>",
))
# 插入位置：根节点第一个子节点，排在所有测试分组之前
```

---

## 第一步：章节与分组规范

### 章节根节点命名
- 格式：`序号 + 空格 + 功能名 + 空格 + 功能需求`
- 示例：`3.3.2 AWS IoT 功能需求`、`4.1 用户登录功能需求`
- 若需求无章节编号，直接用功能名：`支付功能需求`
- 英文术语、协议名、参数名保留原始写法

### 二级分组顺序
按下列顺序组织，有则写，无则跳过，不强制全部出现：

1. **启停与页面入口** — 功能开关、页面跳转路径、默认状态
2. **基础参数校验** — 各参数的默认值、范围、格式、唯一性
3. **附属配置** — 子功能配置，如证书、模板、附加选项
4. **对象选择与核心操作** — 设备/用户/文件的选择与主流程操作
5. **关键联动规则** — 开关依赖、前置条件、级联生效
6. **批量操作与确认机制** — 批量选择、一键操作、二次确认
7. **导入导出或文件交互** — 文件上传下载、格式校验、内容比对
8. **兼容性与异常** — 禁用效果、多功能共存、设备离线、性能

### 分组命名规则
- 优先用测试维度命名，不搬运需求原文标题
- 标题简短抽象可复用：`基础参数校验`、`兼容性与异常`
- 含特定术语时保留英文：`SSL 与证书配置`、`Device Twin 联动`

### 待确认分组高亮规则

若某个二级分组对应的**需求原文包含「待确认」字样**，该分组节点高亮：用 `PENDING_STYLE`——**橙色填充方框（`#FF9F00` + roundedRect）**。在整图"下划线、无边框"的节点中，橙色方框格外醒目。

**实现方式：** 统一样式已在下方「XMind 文件生成规范」的生成模板中定义。普通分组用 `group(title, children)`（下划线、加粗）；含「待确认」的分组改传 `group(title, children, pending=True)` 即套用 `PENDING_STYLE`。

> 若需求中**无「待确认」**内容，不会用到 `PENDING_STYLE`，其样式定义保留在 styles.xml 中无副作用。

---

## 第二步：叶子节点写法

叶子节点是最终测试点，**不再继续细分**。每个叶子节点必须：

- 写成完整一句，包含「操作对象 + 校验条件 + 预期结果 + 覆盖范围」四类信息
- 优先使用「输入/操作 + 结果」复合句，不写「系统支持 XX」纯能力描述
- 默认值、取值范围、开关依赖、客户端/服务端结果写全，不省略关键条件
- 句末标注「X个用例」当该测试点代表后续多条用例

**❌ 禁止写法：**
- `系统支持数据上报` — 纯需求描述，无操作无结果
- `功能正常` / `配置正确` — 抽象结论，缺对象和条件
- `Interval 配置` — 只有名词，无校验动作

**✅ 正确写法：**
- `Interval 默认为空，范围 10~600 seconds，边界值（10s / 600s）配置后保存展示正确，数据按对应间隔上报；超界 / 非数字 / 特殊符 / 负数 等非法值均阻止保存并提示，4个用例`
- `Enable SSL 默认关闭；启用后 Certificate / Key 字段可上传。未启用时无法上传，1个用例`
- `用户名为空时阻止提交并在输入框下提示"用户名不能为空"，1个用例`

---

## 第三步：颗粒度规则

- 策略脑图颗粒度介于需求点与测试用例之间，阅读后可直接继续拆用例
- 参数级基础校验按**一参数一测试点**组织
- 多个强关联校验维度可合并为一个复合测试点
- 设备/对象遍历类测试点写明覆盖范围和预期用例数
- 异常输入、非法配置、错误提示文案写成**独立测试点**

### 方法选择（对照共享方法学文档）

等价类 + 边界值是默认基线。起草前对照 `knowledge/shared/test_design_methods.md` 第一节判断是否叠加以下方法，命中即用，并在脑图中体现：

- **决策表**：功能由 ≥2 个独立条件组合决定结果（如 Wiring Check、分级/折扣规则、报警组合）→ 把有意义的条件组合各列为一个测试点
- **状态转换**：功能存在明确状态机（启停、主备切换、连接断开→重连→在线、注册/订阅生命周期）→ 覆盖各状态迁移，并补关键非法迁移
- **正交 / pairwise**：多因素多水平组合会爆炸（如 多协议 × 多型号 × 多参数 兼容性）→ 用正交压缩，脑图写明覆盖范围与压缩后用例数

---

## 第四步：补充测试点原则

**先做风险扫描，再补充。** 起草补充点前，按 `knowledge/shared/test_design_methods.md` 第二节的五个维度（潜在缺陷区、用户易错操作、系统薄弱环节、依赖与前后置、异常恢复）过一遍，把扫出的项作为补充来源；再用第三节「领域专属场景库」按功能类型取用成熟场景。此外，结合「生成前必做」读取的项目 `context.md` / 相关摘要，补充跨模块联动、功能共存、设备兼容类场景（与当前需求文档冲突时以当前文档为准）。

在需求文档覆盖基础上，主动补充以下类型，句首加 `[补充]` 前缀：

- **边界与非法输入**：字段为空、格式非法、超出范围；**数值字段的非法输入必须覆盖 超界 / 负数 / 非数字字符 / 特殊符 / 空 全部等价类（见方法学文档「按字段类型的强制非法等价类」），且同类字段划分保持一致**
- **开关依赖反向**：未启用时字段不可见/不可配置
- **多功能共存**：多个功能同时启用时互不干扰
- **禁用后效果**：关闭功能后不再对外提供服务
- **异常恢复**：断网重连、服务重启、**掉电/重启后配置持久化**后行为是否正常
- **双栈接入**：IPv4/IPv6 双栈下核心读写/上报流程行为一致（需求未提 IPv6 时补充并注明"待确认是否支持"）
- **性能边界**：大量数据/设备时系统性能是否正常
- **错误推测**：基于工程经验/历史 bug 模式推测的易错点（如单位换算、字节序、时间戳/时区）

---

## 第五步：测试范围外声明

在 XMind 中创建一个**自由主题**，标题为 `测试范围外`，用于集中呈现所有不纳入本轮测试的内容。

**自由主题下包含两类子节点：**

| 子分组 | 来源 | 内容 |
|--------|------|------|
| 已删除需求 | 需求文档中带删除线的内容 | 原文描述，逐条列出 |
| 本轮不覆盖 | 主动声明超出本次测试范围的项 | 如：第三方服务端行为、性能测试、暂不开发功能 |

**识别"已删除需求"：**
- 解析 docx 时，段落 Run 若有 `run.font.strike == True`，标记为删除线内容
- 删除线内容**不生成任何测试点**，不计入用例数量预估

**声明"本轮不覆盖"：**
- 标注「待定」「后续迭代」「暂不实现」的功能点
- 涉及第三方系统内部行为、硬件层验证、超出项目范围的联动场景
- 每条写明原因，格式：`功能描述 — 原因`

**XMind 自由主题生成方式：**

```python
out_of_scope = group("测试范围外",
    group("已删除需求", leaves(
        "删除线内容描述1",
    )) +
    group("本轮不覆盖", leaves(
        "功能描述A — 原因",
    ))
)
# 两个子分组均无内容时，对应子分组省略；若均为空，则不创建自由主题
```

---

## 第六步：用例数量预估标注

| 节点层级 | 标注格式 | 示例 |
|----------|----------|------|
| 根节点 | `章节名（共预估 X 条）` | `3.3.2 AWS IoT 功能需求（共预估 36 条）` |
| 二级分组 | `分组名（预估 X 条）` | `参数校验与合法性（预估 17 条）` |
| 叶子节点 | 句末已有「X个用例」，不重复标注 | — |

**计算规则：**
- 叶子节点句末已标注「X个用例」的，取该数值；未标注的默认计 1 条
- 二级分组 = 其下所有叶子节点用例数之和
- 根节点 = 所有二级分组预估数之和

---

## 生成检查清单

输出前逐项确认：
- [ ] 已确定目标项目并读取其 `context.md`；已扫 `requirements/summaries/` 清单并参考相关摘要补充跨模块/共存/设备兼容场景（项目无对应内容时已说明跳过）
- [ ] 章节根节点格式为「编号 + 功能名 + 功能需求」
- [ ] 二级分组按规定顺序排列，使用测试维度命名
- [ ] 基础参数按「一参数一测试点」，默认值/范围/结果写全
- [ ] 入口、启停、开关依赖、字段可见性、保存行为均已覆盖
- [ ] 已对照共享方法学文档判断方法叠加：多条件组合→决策表、状态机→状态转换、多因素组合→正交
- [ ] 每个可输入非法值的字段，非法等价类已按字段类型补全（数值=超界/负数/非数字字符/特殊符/空），且同类字段划分一致
- [ ] 配置/通信类功能已评估 异常恢复（断网/掉电/重启）与 IPv4/IPv6 双栈，并补充对应测试点或显式说明不适用
- [ ] 已按风险扫描五维度过一遍，扫出的易错/薄弱/异常恢复点已补充
- [ ] 遍历类测试点写明覆盖范围和用例数
- [ ] 多条用例的测试点句末标注「X个用例」
- [ ] 异常输入、非法配置、提示文案写成独立测试点
- [ ] 补充测试点加 `[补充]` 前缀
- [ ] 需求含「待确认」的二级分组已用 `group(..., pending=True)` 生成（橙色方框 `#FF9F00`）
- [ ] 已创建自由主题「测试范围外」；两者均为空时未创建
- [ ] 删除线内容归入「已删除需求」，未计入测试点和用例数
- [ ] 每个二级分组标题已追加「预估 X 条」
- [ ] 根节点标题已追加「共预估 X 条」
- [ ] XMind 文件包含 `META-INF/manifest.xml`
- [ ] 所有节点含 `structure-class="org.xmind.ui.logic.right"`
- [ ] 节点已套样式：根/分组/叶子=下划线无边框（ROOT/GROUP/LEAF_STYLE），待确认=橙色方框（PENDING_STYLE），分支为曲线

---

# 句式模板与语言规范

## 句式模板

### 1. 入口与默认状态
```
通过 <页面路径> 正常进入 <页面名>，<功能名> 默认 <状态>
```
**示例：**
- `通过设置 -> 通信 -> AWS IoT 正常进入配置页面，AWS IoT 默认 Disable`

---

### 2. 参数校验
```
<参数名> 默认 <默认值>，范围 <最小值> ~ <最大值>，参数配置后 <预期结果>
<参数名> 为空时阻止保存并提示，1个用例
```
**示例：**
- `超时时间 默认 30s，范围 1~300s，超出范围时阻止保存并提示，参数配置保存展示正确`
- `端口号 默认 8080，范围 1024~65535，参数配置后服务可正常访问`

---

### 3. 开关依赖与字段可见性
```
<开关名> 启用后 <字段列表> 字段可配置；未启用时无法配置，1个用例
```
**示例：**
- `Enable SSL 启用后 Certificate / Key 字段可上传；未启用时无法上传，1个用例`

---

### 4. 对象选择与核心操作（遍历类）
```
选择 <对象类型>，<操作步骤>，查看 <预期结果>。遍历 <覆盖范围>，<用例数>个用例
```
**示例：**
- `选择 Modbus RTU 设备，配置合法参数并启用，查看云端能否按配置间隔收到对应数据，1个用例`

---

### 5. 禁用效果
```
禁用 <功能名> 后不再对外提供 <服务类型>；重新启用后 <预期恢复状态>，1个用例
```
**示例：**
- `禁用 AWS IoT 后不再向云端发布数据；重新启用后客户端可以正常接收数据，1个用例`

---

### 6. 联动规则
```
<前置开关> 启用后才可启用 <后置开关>；<后置开关> 启用后才可编辑 <参数名>，1个用例
```
**示例：**
- `主连接启用后才可配置备用连接；主连接失效时自动切换至备用连接，1个用例`

---

### 7. 非法输入与错误提示
```
<参数名> 为非法值（如 <举例>）时，阻止保存并提示"<提示文案>"，<用例数>个用例
```
**示例：**
- `IP 地址输入非法格式（如 999.999.0.1 / 纯字母）时阻止保存并提示，2个用例`

---

### 8. 批量操作
```
在 <功能页面> 中选择 <范围> 批量 <操作>，操作结果正确；未修改项保持不变，<用例数>个用例
```

---

### 9. 文件导入导出
```
点击 <功能名> 导出 <文件类型>，比对 <参考对象> 校验文件内容正确
导入 <来源> 的 <文件类型>，查看系统数据是否与文件内容一致
```

---

### 10. 兼容性与性能
```
<功能A> 与 <功能B> 同时启用，查看两者能否各自独立正常运行，互不干扰，1个用例
```
**示例：**
- `AWS IoT 与 Azure IoT 同时启用，查看两侧云端能否各自独立正常接收数据，互不干扰，1个用例`

---

## 语言规范

### 优先使用的判断词
```
正常进入 / 默认 / 范围 / 可配置 / 保存展示正确
启用后 / 未启用时无法配置 / 配置保存后生效
阻止保存并提示 / 阻止提交并提示
查看能否收到 / 无法访问 / 保持不变
不再对外提供服务 / 客户端可以正常接收数据
```

### 格式规范
- 页面路径用箭头：`设置 -> 通信 -> AWS IoT`
- 多字段并列用斜线：`Certificate / Key / Interval`
- 遍历范围显式写出，不用"相关""部分"等模糊表述
- 中文为主，技术术语、参数名、按钮名保留英文

### 禁止写法
| ❌ 错误 | ✅ 正确 |
|--------|--------|
| 系统支持上传文件 | 上传合法格式文件后系统正确解析并展示，1个用例 |
| 功能正常 | 配置保存后客户端可以正常接收数据 |
| 遍历相关设备 | 遍历 Modbus RTU / TCP / BACnet IP 设备，3个用例 |

---

# XMind 文件生成规范

## 文件格式说明

XMind 26 使用 XMind 8 XML 格式，本质是一个 ZIP 包，包含：

```
MyFile.xmind
├── content.xml          # 脑图内容（必须）
├── styles.xml           # 样式（必须，可为空）
└── META-INF/
    └── manifest.xml     # 文件清单（必须，缺失则报"无法打开文件"）
```

> ⚠️ `META-INF/manifest.xml` 是最常见的缺失项，缺少它 XMind 26 会直接报错拒绝打开。

---

## 布局与节点样式（固定）

整图采用**逻辑图（向右）+ 下划线节点（无边框）+ 曲线分支**，三级字号区分：

| 节点 | 形状 | 字重/字号 | 分支线 |
|------|------|-----------|--------|
| 根节点 | 下划线 underline（无边框） | 加粗 20pt | 曲线 curve 3pt |
| 二级分组 | 下划线 underline（无边框） | 加粗 14pt | 曲线 curve 2pt |
| 叶子节点 | 下划线 underline（无边框） | 常规 12pt | 曲线 curve 1pt |
| 待确认分组 | 圆角方框 roundedRect + 橙色填充 `#FF9F00` | 加粗 14pt | 曲线 curve |

- 所有 `<topic>` 必带 `structure-class="org.xmind.ui.logic.right"`。
- 文字坐在分支线上方（下划线样式），不套方框；唯有「待确认」分组用橙色方框，在下划线节点中醒目。

---

## 完整 Python 生成模板

```python
import zipfile, uuid, time, os

def uid():
    return uuid.uuid4().hex[:26]

ts = str(int(time.time() * 1000))

def escape(s):
    return (s.replace('&', '&amp;')
             .replace('<', '&lt;')
             .replace('>', '&gt;')
             .replace('"', '&quot;'))

LOGIC = "org.xmind.ui.logic.right"

# 四个样式 ID（脚本顶部生成一次，全局复用）
ROOT_STYLE    = uid()
GROUP_STYLE   = uid()
LEAF_STYLE    = uid()
PENDING_STYLE = uid()   # 仅含「待确认」的二级分组使用

def topic(title, children_xml="", style_id=None):
    t = uid()
    sid = f' style-id="{style_id}"' if style_id else ''
    inner = f'<title>{escape(title)}</title>'
    if children_xml:
        inner += f'<children><topics type="attached">{children_xml}</topics></children>'
    return f'<topic id="{t}" timestamp="{ts}" structure-class="{LOGIC}"{sid}>{inner}</topic>'

def leaves(*titles):
    """叶子节点：下划线、常规字重"""
    return ''.join(topic(t, style_id=LEAF_STYLE) for t in titles)

def group(title, children_xml="", pending=False):
    """二级分组：下划线、加粗；含「待确认」时传 pending=True → 橙色方框"""
    return topic(title, children_xml, style_id=PENDING_STYLE if pending else GROUP_STYLE)

# ---- 构建脑图内容（示例）----
overview = group("功能概述", leaves(
    "功能定位：<一句话描述>",
    "工作流程：<A → B → C>",
    "关键约束：<约束，分号分隔>",
))

g1 = group("基础参数校验（预估 X 条）", leaves(
    "URL 默认为空，字符长度范围 20~128，配置保存后展示正确……，1个用例",
    "Topic 默认为空，字符长度范围 1~128……，1个用例",
))
# 含「待确认」分组示例：group("Meter Point 配置 待确认（预估 X 条）", leaves(...), pending=True)

root = topic("3.3.X 功能需求（共预估 X 条）", overview + g1, style_id=ROOT_STYLE)

sheet_id = uid()
content_xml = f'''<?xml version="1.0" encoding="utf-8"?>
<xmap-content xmlns="urn:xmind:xmap:xmlns:content:2.0"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:xhtml="http://www.w3.org/1999/xhtml"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:svg="http://www.w3.org/2000/svg"
  timestamp="{ts}" version="2.0">
  <sheet id="{sheet_id}" timestamp="{ts}">
    {root}
    <title>测试策略</title>
  </sheet>
</xmap-content>'''

# 统一样式：根/分组/叶子=下划线·无边框·曲线分支；待确认分组=橙色方框
styles_xml = f'''<?xml version="1.0" encoding="utf-8"?>
<xmap-styles xmlns="urn:xmind:xmap:xmlns:style:2.0"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:svg="http://www.w3.org/2000/svg" version="2.0">
  <styles>
    <style id="{ROOT_STYLE}" type="topic">
      <topic-properties shape-class="org.xmind.topicShape.underline" border-line-width="0pt"
        line-class="org.xmind.branchConnection.curve" line-width="3pt" line-color="#000000"
        fo:font-weight="bold" fo:font-size="20pt" fo:color="#000000"/>
    </style>
    <style id="{GROUP_STYLE}" type="topic">
      <topic-properties shape-class="org.xmind.topicShape.underline" border-line-width="0pt"
        line-class="org.xmind.branchConnection.curve" line-width="2pt" line-color="#000000"
        fo:font-weight="bold" fo:font-size="14pt" fo:color="#000000"/>
    </style>
    <style id="{LEAF_STYLE}" type="topic">
      <topic-properties shape-class="org.xmind.topicShape.underline" border-line-width="0pt"
        line-class="org.xmind.branchConnection.curve" line-width="1pt" line-color="#000000"
        fo:font-weight="normal" fo:font-size="12pt" fo:color="#000000"/>
    </style>
    <style id="{PENDING_STYLE}" type="topic">
      <topic-properties shape-class="org.xmind.topicShape.roundedRect" svg:fill="#FF9F00"
        border-line-width="1pt" border-line-color="#FF9F00"
        line-class="org.xmind.branchConnection.curve" line-width="1pt"
        fo:font-weight="bold" fo:font-size="14pt" fo:color="#000000"/>
    </style>
  </styles>
</xmap-styles>'''

manifest_xml = '''<?xml version="1.0" encoding="utf-8"?>
<manifest xmlns="urn:xmind:xmap:xmlns:manifest:1.0">
  <file-entry full-path="content.xml" media-type="text/xml"/>
  <file-entry full-path="styles.xml" media-type="text/xml"/>
  <file-entry full-path="META-INF/" media-type=""/>
  <file-entry full-path="META-INF/manifest.xml" media-type="text/xml"/>
</manifest>'''

out = os.path.join(os.getcwd(), '输出文件名.xmind')
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.writestr('content.xml', content_xml.encode('utf-8'))
    zf.writestr('styles.xml', styles_xml.encode('utf-8'))
    zf.writestr('META-INF/manifest.xml', manifest_xml.encode('utf-8'))

print(f'已生成：{os.path.abspath(out)}')
```

---

## 关键注意事项

1. **`META-INF/manifest.xml` 必须存在**，否则 XMind 26 报"无法打开文件"
2. **所有节点必须带 `structure-class`** 才能呈现为逻辑图
3. **节点样式必须套用 style-id**：根/分组/叶子分别用 `ROOT_STYLE`/`GROUP_STYLE`/`LEAF_STYLE`（下划线、无边框、曲线分支），待确认分组用 `PENDING_STYLE`（橙色方框）。不套 style-id 会退回默认方框样式，不符合规范
4. **样式属性勿写错**：形状 `shape-class="org.xmind.topicShape.underline"`、分支 `line-class="org.xmind.branchConnection.curve"`、填充 `svg:fill`、字重 `fo:font-weight`、字号 `fo:font-size`
5. **特殊字符必须 escape**：`&` `<` `>` `"` → 用 `escape()` 函数处理
6. **每个 topic id 必须唯一**，用 `uuid.uuid4().hex[:26]` 生成
7. **输出路径**：`os.path.join(os.getcwd(), '文件名.xmind')`，输出到当前工作目录，全平台通用；生成后 `print` 输出完整绝对路径
8. **文件命名规范**：`{项目}_{章节}_{描述}.xmind`
