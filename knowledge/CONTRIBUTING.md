# 知识库维护说明

## 知识库设计原则

1. **Claude 查阅效率优先**：文件结构、命名、内容组织以 Claude 快速定位为第一目标——精简索引优于完整罗列，关键信息前置，避免重复。
2. **测试工程师可读性次之**：在不损害第一原则的前提下，保持 Markdown 格式清晰、中文表述准确。
3. **不存可推导的内容**：测试用例、代码注释、git 历史能覆盖的内容不重复维护；知识库只存"代码读不出来的东西"：历史决策、已知坑、Bug 历史、需求背景。
4. **每次新需求到来时，分析是否需要更新 CLAUDE.md**：若新增了项目、设备、全局约定或改变了工作流，则同步更新；若只是项目内部细节，记入对应 context.md 即可。

---

## 目录结构与维护责任

| 路径 | 存放内容 | 维护人 |
|------|---------|--------|
| `CLAUDE.md` | 全局入口，Claude 自动加载。新增项目/设备时同步更新项目表和设备速查 | 各项目负责人 |
| `.claude/commands/` | 斜杠技能。新增工作流时添加 .md 文件，文件名即命令名 | 知识库维护人 |
| `knowledge/shared/devices/` | 每台电表一个 .md 文件，记录协议、寄存器、已知坑、项目适配 | 各设备负责人 |
| `knowledge/shared/modbus_tables/raw/` | Modbus 地址表原始 Excel，代码直接读取，更新时替换并同步 INDEX.md | 硬件/固件团队 |
| `knowledge/shared/templates/raw/` | 参数模板原始 Excel（blockParams），更新时替换并同步 INDEX.md | 硬件/固件团队 |
| `knowledge/shared/conventions.md` | 团队编码约定，规范变更时更新 | 知识库维护人 |
| `knowledge/shared/decisions.md` | 历史决策记录，每次做重要设计决策后追加 | 决策发起人 |
| `knowledge/gateway/<项目>/context.md` | 项目全貌，模块增减时更新 | 项目负责人 |
| `knowledge/gateway/<项目>/bugs/INDEX.md` | Jira bug 精简索引，每次导出 Jira 后人工更新 | 项目负责人 |
| `knowledge/gateway/<项目>/bugs/raw/` | Jira 导出 Excel 原件，只存档不编辑 | 项目负责人 |
| `knowledge/gateway/<项目>/requirements/raw/` | 需求文档原件（Word/PDF），只存档 | 项目负责人 |
| `knowledge/gateway/<项目>/requirements/summaries/` | 需求摘要 .md，新需求到来时人工撰写 | 项目负责人 |
| `knowledge/meters/` | 电表测试项目组，结构与 gateway 相同，内容待填入 | 电表测试负责人 |
| `knowledge/cloud/` | 云平台测试项目组，结构与 gateway 相同，内容待填入 | 云平台测试负责人 |
| `knowledge/_template/` | 新项目初始化模板，`/new-project` 技能使用，结构变更时同步更新 | 知识库维护人 |

## 日常维护流程

| 触发事件 | 需要更新的文件 |
|---------|--------------|
| 新增电表设备 | `knowledge/shared/devices/<name>.md`、`modbus_tables/INDEX.md`、`templates/INDEX.md`、`CLAUDE.md` 设备速查表 |
| Jira 导出新 bug 单 | `knowledge/<类型>/<项目>/bugs/INDEX.md`（追加精简行）、`bugs/raw/`（存原始 Excel） |
| 新需求文档到来 | `requirements/raw/`（存原件）、`requirements/summaries/`（写摘要）、对应项目 `context.md`（同步功能模块）、**分析是否需要更新 `CLAUDE.md`**（新项目/新设备/全局约定变更才更新，项目内细节不更新） |
| 新增项目 | 运行 `/new-project`，然后更新 `CLAUDE.md` 项目一览表 |
| 重要设计决策 | `knowledge/shared/decisions.md` 追加一条 |
| 团队约定变更 | `knowledge/shared/conventions.md` 更新对应条目 |
