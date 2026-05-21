# 历史决策记录

记录"为什么这样设计"，而不是"做了什么"。代码能说明做了什么，这里记录背景和理由。

---

## [web2] 元数据检查只验证单位，不验证 description
**时间：** 2026-05
**决策：** MetaCheckResult 只保留 unit_ok，不检查 desc_ok。
**原因：** param_key 已作为唯一索引；description 是人类可读备注，不同固件版本写法不一致，实测误报率高，实际价值低。

---

## [web2] BACnet 读取超时参数调高
**时间：** 2026-05
**决策：** READ_TIMEOUT=30s、MAX_RETRIES=4、CONNECT_WAIT=3.0s、WHOIS_TIMEOUT=10s。
**原因：** 当前网关设备（web2，192.168.2.209）性能较差，低参数下频繁超时导致测试失败。若换用高性能网关可适当调小。

---

## [web2] IOM-03/04 返回空 map，暂不支持
**时间：** 2026-05
**决策：** devices/acuiom03.py 和 acuiom04.py 的 build_param_map() 返回空 {}。
**原因：** DI Status 使用 FC 0x02（Discrete Inputs），当前 ModbusReader 只实现 FC 0x03（Holding Registers）。BACnet 侧发的是 Binary Input 对象，不是 Analog Input，整个框架结构不匹配。
**待办：** 需单独实现 FC02 读取路径 + Binary Input 解析才能支持。
**影响设备：** → [acuiom03.md](./devices/acuiom03.md)、[acuiom04.md](./devices/acuiom04.md)

---

## [web2] AcuCloud 比对容差设为 ±5% / ±1.0
**时间：** 2026-05
**决策：** CLOUD_TOLERANCE_PERCENT=5.0，CLOUD_TOLERANCE_ABSOLUTE=1.0。
**原因：** 快照数据与实时 Modbus 读取存在时序差异（非同一时刻采集），±1% 误报率极高。BACnet 比对为近实时并发读取，可保持 ±1%。

---

## [知识库] 不存测试用例，bug 单优先于用例
**时间：** 2026-05
**决策：** 知识库不维护测试用例文档；bug 单精简索引（bugs/INDEX.md）是重点维护内容。
**原因：** 测试用例是可从需求推导出的可执行规范，随每次需求变更都需同步修改，放进 Markdown 会产生双重维护负担且快速过时。Bug 单记录的是已发生的失败，代码无法告诉你"这里曾经真的坏过"——已关闭的 Bug 也是未来回归测试的重点区域标记，信息密度更高、维护成本更低（只追加，不改旧）。

---

## [知识库] 项目一览使用正式产品型号作为项目名
**时间：** 2026-05
**决策：** CLAUDE.md 项目一览的「项目」列改用产品型号（ACM-41-WEB2 / AcuHMI-1-7 / AcuCloud），废弃内部简称（web2 / hmi1-7 / acucloud）。
**原因：** 简称在跨组沟通和 Jira Bug ID 前缀对应时歧义大；产品型号是官方命名，与需求文档、硬件标签、Jira 项目 Key 保持一致，减少认知负担。目录路径（web2/、hmi1-7/）仍保持简短不变，仅显示名称更新。

---

## [知识库] 临时文件清理必须用 Glob 全量扫描，不能手动列举

**时间：** 2026-05
**决策：** 任务结束前，用 `Glob("_tmp_*")` 扫描根目录所有临时文件并一次性删除，禁止手动列举文件名清理。
**原因：** 跨多轮操作创建的临时文件（如多个任务的 JSON/txt）分散在不同步骤，手动列举极易遗漏。用 Glob 全量扫描才能保证彻底清理。

---

## [知识库] 触发关键词后先扫目录再行动

**时间：** 2026-05
**决策：** 收到"上传 X"类触发词时，不立即要求用户提供文件，而是先检查对应目录（testcase/ / requirements/raw/ / bugs/raw/）是否已有文件，有则直接处理。
**原因：** 用户说"上传 X"不代表文件附在消息里，文件可能已被手动放入对应目录。直接开口要文件会产生无效交互轮次，且违背"触发词→立即执行流程"的设计初衷。

---

## [web2] AcuIOM 模板格式不同，范围检查 graceful skip
**时间：** 2026-05
**决策：** get_bacnet_params() 在 IOM 模板上会抛 ValueError（无 BACnetIP 列），用 try/except 捕获并 log warning，跳过范围检查继续数值比对。
**原因：** IOM 模板由不同团队维护，格式规范不统一，短期内无法修改模板文件。
