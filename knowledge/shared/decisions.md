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

## [协议脚本] 新增协议脚本时必须同时检查两类盲区

**时间：** 2026-05
**决策：** 新协议脚本完成后，在开始写代码前以及 code review 前必须问两个问题：
1. **"哪些检查是我们的工具能做但当前没做的？"** — 从自身工具能力出发，盘点静态分析、元数据校验、格式合规等已具备条件却未实现的检查项。
2. **"哪些检查是第三方工具会做但我们脚本感知不到的？"** — 模拟其他消费方（PLC、EDS 安装工具、云平台、ODVA 验证器）的视角，找出我们脚本因绕过某层协议而天然缺失的覆盖面。

**原因：** EtherNet/IP 脚本用显式消息（Get_Attribute_Single）直接读 Assembly，绕过了 `[Connection Manager]`，导致 Connection1 字段缺失、LREAL 8 字节对齐违规、EDS Revision 不一致等问题在我们测试中完全无症状，全部漏检——直到外部 PLC 公司测试才发现。三项问题均可通过 EDS 静态解析在无 PLC 环境中自动检出，不存在技术障碍，只是当初视角局限。

---

## [调试] 页面读取/元素定位问题优先用 Playwright 脚本直接调试

**时间：** 2026-05
**决策：** 遇到页面元素定位、表格读取、CSS 类名等不确定的问题时，第一时间用 Bash 运行独立 Playwright headless 脚本直接探查（打印 table 数量、class、首行内容等），不让用户反复跑完整测试流程再贴日志。
**原因：** ACM-41-WEB2 接线检查调试中，因页面含弹窗 el-table（class 同为 el-table__body），导致索引偏移、结果全部 Not Checked。通过 Bash 直接调试 30 秒内定位根因；此前让用户多次跑 20s+ 的完整用例循环，效率极差。

---


## [web2] AcuIOM 模板格式不同，范围检查 graceful skip
**时间：** 2026-05
**决策：** get_bacnet_params() 在 IOM 模板上会抛 ValueError（无 BACnetIP 列），用 try/except 捕获并 log warning，跳过范围检查继续数值比对。
**原因：** IOM 模板由不同团队维护，格式规范不统一，短期内无法修改模板文件。

---

## [全局·最高优先级] 绝对禁止在项目任何文件夹保留 Claude 中间调试脚本

**时间：** 2026-06
**决策：** Claude 在调试/探查过程中创建的任何临时脚本（探查页面的 Playwright 脚本、临时 Python 验证脚本、`_tmp_*` 文件、一次性 JSON/txt 输出等），必须在当前任务结束前**全部删除**，不允许残留在仓库的任何目录中。任务收尾时除 `Glob("_tmp_*")` 外，还应核对本次会话中所有 Write 创建的非交付文件并逐一清理。
**原因：** 中间调试脚本不是交付物，残留会污染 git 状态、误导后续维护者把调试代码当成正式脚本，且可能被误提交进版本库。本条为最高优先级硬性约束，优先级高于其他清理约定。
**关联：** 补充并强化 [知识库] 临时文件清理必须用 Glob 全量扫描 一条——Glob 扫描只是手段，"零残留"才是验收标准。

---

## [hmi1-7] 配置保存类 UI 用例必须规避前端"脏检查"，并用真实 toast 文案/类名断言

**时间：** 2026-06
**决策：** AcuHMI-1-7 Web 页面保存类用例（如 About → Device Information 的 Name/Location/Description）一律遵循三条：
1. **规避脏检查**：进入页面先读输入框回显值（即设备已保存值），保存时选一个**与之不同**的目标值，确保触发真实保存。单字段用固定值与备用值交替（如 `"a"*40 ↔ "z"*40`），多字段用 A/B 两组值轮换。
2. **成功断言用真实信号**：成功 toast 真实 class 为 `.el-message--success`、文本为 `"Device info saved"`，用 `expect(page.locator(".el-message--success")).to_contain_text("Device info saved", timeout=8000)`。**禁止**用 `get_by_text("success")`——真实文案不含 "success" 一词，永远匹配不到，是失效断言（负向"无成功提示"判定同理改用 `.el-message--success` count==0）。
3. **auto-wait 替固定等待**：toast 为瞬时元素（点击后约 62ms 出现、存活约 3.5s 后 DOM 移除），必须用 `expect(...).to_be_visible/to_contain_text` 轮询捕获，禁止 `wait_for_timeout(N)` + 一次性 `is_visible()/count()` 的脆弱写法。

**原因：** 真机观测确认设备前端对配置保存做了脏检查：当提交值与已保存值**完全相同**时，不发 API 请求、直接弹 `el-message--warning`「No change to save」，不出现成功提示。`009_08` 全家用例原先填固定值（如 `"a"*40`），连跑第二遍时该值已是已保存值 → 弹 warning → 断言成功 toast 失败，表现为"第二遍必挂"的抖动；这也是 2026-06-16 整夜连跑中 about case01/02/03 失败的真实根因（最初误判为"断言文案写错"，实为值未变 + 断言写法脆弱）。

**真机 toast 速查（el-message）：**
- 成功：class `el-message--success`，文本 `Device info saved`，无关闭按钮
- 无变更：class `el-message--warning is-closable`，文本 `No change to save`
- 出现时机 ≈ 点击后 62ms；稳定可见 ≈ 2.4s；DOM 完全移除 ≈ 3.5s

**关联：** 与 [调试] 页面读取/元素定位问题优先用 Playwright 脚本直接调试 一条呼应——本结论即由真机 Playwright 探查得出。模式已落地于 `projects/acuhmi_1_7/tests/ui/about/general/test_TestCase_AcuHMI_009_08_case*.py`，其他保存类用例可直接复用。
