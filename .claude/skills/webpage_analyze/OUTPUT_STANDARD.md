# 页面上下文文档（`_context.md`）输出规范

## 目标

每个**可路由子页**产出一份 `<Prefix_SubPage>_context.md`，作为：
- **选择器沉淀**（约定 #11）：页面结构/选择器一次探明、后续 UI 用例直接复用；
- **手工用例转自动化的输入**：供 `webtestcase_manual_to_auto` 消费。

格式**对齐仓库现有文件**：`knowledge/<项目路径>/requirements/context/*_context.md`（权威样例见 `Devices_DataLog_AcuCloud_context.md`、`SystemSettings_Network_context.md`、`Templates_TemplateList_context.md`）。本规范即从这些文件抽象而来——**产出必须与它们同构**。

---

## 一、文件命名与存放

### 1.1 存放目录
```
<项目知识库根>/requirements/context/
```
`<项目知识库根>` = 项目一览中该项目 `context.md` 所在目录（如 `knowledge/gateway/hmi1-7`）。

### 1.2 单页命名规则（菜单路径 PascalCase）
取被测页的 **Web 导航菜单路径**，每段 PascalCase、去空格/特殊字符、下划线连接，尾缀 `_context.md`：

| 菜单路径 | 文件名 |
|----------|--------|
| Devices / Data Log / AcuCloud | `Devices_DataLog_AcuCloud_context.md` |
| System Settings / Network | `SystemSettings_Network_context.md` |
| Protocols / MQTT / SSL·TLS | `Protocols_MQTT_SSL_TLS_context.md` |

- **一子页一份**（可独立路由/操作的页面）。
- **代表性合并**：结构一致仅编号不同的子页合并为一份，文件名取无编号形式，并在正文注明覆盖范围（如 `Devices_DataLog_DataLogger_context.md`「覆盖 Data Logger 1/2/3」）。
- 与 `projects/<项目>/tests/ui/` 的一级目录（小写，如 `datalog`）是**多对多**关系，禁止用测试目录名拼文件名。

### 1.3 索引 `_INDEX_context.md`
产出所有单页文档后，**必须**同目录维护 `_INDEX_context.md`（见 §五）。已存在则**增量合并**（更新既有条目、补新条目，不删他人条目）。

---

## 二、单页文档骨架

章节可按页面复杂度微调（简单页可省「页面结构」「校验规则要点」），但**顺序与编号沿用现有文件**：

```
# {菜单路径} — 页面上下文

## 1. 页面标识              ← 表：路由 / 路由名 / 面包屑 / 顶级模块或上下文侧
## 2. 页面用途              ← 一句话用途 + 条件分支/二级 tab 说明
## 3. 页面结构（复杂页可选）  ← 区域/表/tab 纵向布局概述
## 4. 交互元素清单 / 表单字段 ← 核心表：元素 / 类型 / getByRole 定位 / 默认值 / 说明
## 5. 页面状态与分支         ← 状态矩阵（条件显隐、Enable/Disable、DHCP Auto/Manual…）
## 6. 校验规则要点（有表单则填）
## 7. 自动化测试要点         ← 核心用例点 + ⚠️ 高危提示
## 8. 机器可解析摘要         ← JSON 块（下游读取）

## 实测测试情报（pytest / Element Plus，来源：{日期} 联机探查）  ← 见 §四
```

> 章节编号连续即可（简单页 6 章、复杂页 8 章都合规）；**「机器可解析摘要」JSON 块与「实测测试情报」节不可省**。

---

## 三、各章字段规范

### 3.1 页面标识（表）
| 项 | 值 |
|----|----|
| 路由 | `/#/systemSettings/network` |
| 路由名 | `network` |
| 面包屑 | AcuHMI-1-7 / System Settings / Network |
| 顶级模块 / 上下文侧 | System Settings（或「Devices 侧」） |

### 3.2 交互元素清单（核心表）
每个交互元素一行，字段：**元素 / 类型 / 定位策略1（`getByRole` 优先）/ 默认值·示例 / 说明**。复杂控件可加「定位策略2」列。

```
| 字段/元素 | 类型 | 定位策略 1 | 默认/示例 | 说明 |
|-----------|------|-----------|-----------|------|
| AcuCloud Enable | radiogroup | `getByRole('radiogroup',{name:'AcuCloud Enable'})` | Disable | 必选(*)，控制显隐 |
| AcuCloud Token | textbox | `getByRole('textbox',{name:'AcuCloud Token'})` | — | 必填(*)，≤40 字符 |
| Save | button | `getByRole('button',{name:'Save'})` | — | 保存 |
```

**定位优先级**：`getByRole(name)` > `getByPlaceholder` > 框架类名 `.el-*` + scope > nth。**禁止**只给一种脆弱定位；同名元素必须标注父容器 scope。

### 3.3 页面状态与分支（状态矩阵）
```
| 状态 | 触发 | 结果 |
|------|------|------|
| DHCP = Auto | 默认 | Interface Status/IP 只读展示 |
| DHCP = Manual | 切换 | 显示 IP*/Subnet Mask*/Gateway* 三必填框 |
```
条件显隐、Enable/Disable 分支、协议切换后新增字段**必须触发并记录**（对应 EXHAUSTIVE.md 五「动态内容」）。

### 3.4 自动化测试要点
核心正/负/边界用例点，末尾用 `⚠️` 标高危（reboot / factory reset / clear logs / 改 IP 断连 / firmware upload / delete）。

### 3.5 机器可解析摘要（JSON，必填）
```json
{
  "route": "/systemSettings/network",
  "name": "network",
  "title": "Network",
  "module": "System Settings",
  "fields": {
    "DNS 1": {"type":"text","required":true,"format":"ip_or_domain"},
    "Ethernet1.DHCP": {"type":"radio","options":["Auto","Manual"],"default":"Auto"}
  },
  "conditional": {"when":"DHCP=Manual","shows":["IP*","Subnet Mask*","Gateway*"]},
  "buttons": ["Save"],
  "sub_tabs": ["..."]
}
```
字段随页面类型伸缩（表格页加 `tables`，tab 页加 `sub_tabs`），但 `route`/`name`/`title` 必填。

---

## 四、「实测测试情报」节（强制，实测采集不得推测）

标题固定：`## 实测测试情报（pytest / Element Plus，来源：{日期} 联机探查）`，首行给出对应测试目录指针（如 `> 对应测试目录：projects/AcuHMI_1_7/tests/ui/systemsettings/`）。按需含以下小节：

- **进入路径 / SPA 导航**（见 [PITFALLS.md](PITFALLS.md) §1/§2）：左导航/header 上下文切换的真实点击链；**goto 可达性**（直达可用 / 存在守卫须菜单点击链，附侧边栏+子菜单定位）；列表/分页页标注**同路由 goto 是否重置状态**（否则须显式重置或二次 goto）。
- **加载态 / 写操作 API**：`browser_network_requests` 实测的 GET/POST（method + path）。
- **pytest 选择器与控件**：框架真实选择器（`el-radio`/`el-select`/`.c_common_table`/Ant `tr.ant-table-row` 等）+ 默认值；含 Action 列表格标注**行按钮多样性**（所有行相同=固定 locator / 行间不同=下游动态查找，见 PITFALLS §5.2）。
- **校验时机（实测）**：逐字段标「blur 即报错」或「仅 Save 后」，错误文案抄原文；**依赖字段多状态**（B 依赖 A：A 未触发时 B disabled/不参与校验、A 触发后 B 启用/回填值，见 PITFALLS §4）；多入口表单以"空表单入口"干跑为准。
- **框架通用坑（Element Plus / Ant Design，先 JS 计数定框架，见 [PITFALLS.md](PITFALLS.md) §3）**：El Plus—radio 点 label 兜底、`el-select` aria-controls 且选项 teleport 到 body、同名 group 父容器 scope、"Post Channel N" 带空格、`is-disabled` 保留、`*` 号 CSS 生成不入 innerText、已选值跳过空节点；Ant—Table 用 `tr.ant-table-row`、Select 用 `.ant-select`/多选 `.ant-select-selection-search-input`。
- **异步/结果反馈 / 跨页传播**（见 [PITFALLS.md](PITFALLS.md) §5.1）：列表异步 reload 轮询、toast 文案（`.el-message--success`）、自定义 Yes/No 确认框；跨页断言场景记**传播延迟**（单跳/链路/长任务/跨上下文，含实际轮询次数×间隔或"超时未生效"）。
- **高危**：不可逆/断连操作，标「执行前须确认、禁无人值守」。
- **参考用例**：已有则填 pytest 路径；无则标「待补（用例生成后回填）」。

> 这些情报由探索阶段（SKILL.md 阶段1「API/框架识别（JS 计数）」、阶段2「SPA 导航实测」、阶段3+「校验时机/依赖字段多状态/结果反馈/跨页传播/高危」）实测采集，天然填入本节；五类工程坑的检测方法与记录字段详见 [PITFALLS.md](PITFALLS.md)。

---

## 五、`_INDEX_context.md` 规范

结构对齐现有索引（`knowledge/gateway/hmi1-7/requirements/context/_INDEX_context.md`）：

```
# {产品} — 页面上下文索引
> 产品 / 站点 / 登录 / 用途（每个可路由子页一个文件）/ 颗粒度

## 通用说明（所有页面适用）
- 技术栈（如 Vue3 + Element Plus，hash 路由）
- 框架通用坑（radio label 兜底、两个导航上下文…）
- 破坏性操作清单（Factory Reset / Reboot / Clear Logs / Firmware Upload / Delete）

## {导航上下文A，如「设置侧」}
### {菜单组}
- `Prefix_SubPage_context.md`（备注）
...
## {导航上下文B，如「设备侧」}
- ...

## 局限与未尽项
- 动态详情页 / 代表性合并 / 超大表仅记结构 / 情报覆盖范围等
```

---

## 六、质量自检清单

产出后逐项自检，任一缺失即不合格：

```
[ ] 文件名符合菜单路径 PascalCase + _context.md
[ ] 落在 <项目知识库根>/requirements/context/，中间快照零残留
[ ] 每个可路由子页各一份文档（代表性合并已注明覆盖范围）
[ ] 交互元素清单：每元素含 getByRole 定位（同名元素标 scope）+ 类型 + 默认值 + 说明
[ ] 页面状态与分支：条件显隐/Enable/协议切换分支已触发并记录
[ ] 机器可解析 JSON 块存在且含 route/name/title
[ ] 「实测测试情报」节：API / 校验时机(blur|Save) / 框架真实选择器与坑 / 高危 已实测填写
[ ] 框架类型经 JS DOM 计数确认（含下拉/选择控件时，非仅凭快照类名猜），Ant/El Plus 未混淆（PITFALLS §3）
[ ] SPA 站点已记 goto 可达性/路由守卫；列表页已记同路由 goto 是否重置状态（PITFALLS §1/§2）
[ ] 依赖字段/多入口表单已探两种状态，未把从未触发的校验错误当必填规则（PITFALLS §4）
[ ] 可枚举下拉/单选选项已全量抄录真实文案（无"运行时确认/若干"占位）
[ ] _INDEX_context.md 已更新（分组+新条目，未删既有条目）
```
