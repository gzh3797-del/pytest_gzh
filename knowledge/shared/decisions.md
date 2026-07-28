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

---

## [AcuRev-100] RACG 接线检查表 V-N 反接系数 K 不一致点裁定：以各判据表实际值为准
**时间：** 2026-07
**决策：** RACG 接线检查表（`knowledge/meters/AcuRev100/requirements/raw/RACG_接线检查表.xlsx`）中 V-N 反接比值阈值 K：2E3W1P=1.4、3E4WY=1.3，均以各接线方式判据表 sheet 内的实际数值为准；「修改说明」sheet 声明的"统一 K=1.4"仅对 2E3W1P 成立，对 3E4WY 为笔误，不采纳。
**原因：** 用户于 2026-07-14 裁定：3E4WY 判据表（含条件 8/9/10）实测使用的 K=1.3 才是正确值，「修改说明」sheet 的统一声明有误。此前曾有一次"统一改成 1.4"的更新任务中途被手动停止。
**待办：** 若后续从研发处拿到官方勘误确认，可移除本条"用户裁定"表述，直接固化为定论。
**影响设备：** → [RACG_接线检查表.md](../meters/AcuRev100/requirements/summaries/RACG_接线检查表.md)、[context.md](../meters/AcuRev100/context.md)、[wiring_reference.md](./wiring_reference.md)

---

## [AcuRev-100] RACG 接线检查表检测周期口径 + PF/Q 符号约定裁定
**时间：** 2026-07
**决策：**
1. **检测周期**：接线检查对外告警检测节奏按 **1 次/min** 设计；补充机制——接线检查开关**开启的瞬间立即执行一次检查**，此后若一直开着则每 1 分钟检查一次。原摘要中"检测周期 1s（APP_WiringTask，HLD §3.6.3）"与 context 必读区"1 次/min"的口径矛盾以此裁定为准：HLD §3.6.3 的 1s 是 APP_WiringTask 内部任务轮询周期，非对外告警检测周期，二者非同一概念。
2. **PF/Q 符号约定**：按目前口径执行（进线为正，沿用现有判据符号方向），原"★ 待与 IIV3 口径核对一致性（未定稿）"风险标注解除。
**原因：** 用户于 2026-07-14 对两处存疑点直接裁定，消除口径矛盾与未定稿风险，测试设计可按裁定口径固化用例。
**影响设备：** → [RACG_接线检查表.md](../meters/AcuRev100/requirements/summaries/RACG_接线检查表.md)、[context.md](../meters/AcuRev100/context.md)

---

## [AcuRev-100] acuview_auto GUI write_verify 引擎升级：4 类用例函数 + allow_write 护栏放行
**时间：** 2026-07
**决策：** `comm/ctl_acuview/testcase_engine.py` 新增 4 类共享引擎函数支撑 015/016/017 及 009 GUI 类用例：`run_button_action_case`（按钮动作+确认弹窗+Modbus 效果验证，`is_reset=True` 走"等掉线→等重连恢复"判据）、`run_comm_param_case`（通讯参数 GUI 下发后用新参数 Modbus 重连校验，还原走 Modbus 直写兜底；因改 RS485 口参数会断 Acuview 自身 COM11 链路，016 波特率/校验统一改 USB 口参数）、`run_reject_case`（非法输入拒绝类，判据=Modbus 回读≠非法值）、`run_multi_write_verify_case`（多值序列写回读）。同时补齐 config 中声明却从未实现的 `safety.forbid_write_addr/forbid_write_name_substr` 护栏，并新增逐用例 `allow_write` 受控放行机制（016 需改被禁写的 SlaveID/波特率/校验寄存器，经 allow_write 显式放行+自带重连/还原兜底）；`modbus_client.py::MeterClient` 增加 slave_id/baudrate/parity 覆盖参数；`gui_driver.launch_or_connect` 改用 `_find_main_window()` 按标题排除干扰+取最大可见窗口，修复 portable debug 构建多同名窗口导致 pywinauto `connect()` 抛 ElementAmbiguousError。
**原因：** GUI 类用例（按钮动作/通讯参数/非法输入拒绝/多值写回读）此前缺乏共享引擎，各用例各自实现易漏判据；护栏声明与实现脱节存在误写入风险；通讯参数用例天然需要突破护栏但必须受控且自带还原兜底，避免设备遗留在非默认通讯参数；多同名窗口歧义是真机调试中实测复现的连接失败根因。
**影响设备：** → [context.md](../meters/AcuRev100/context.md)

---

## [AcuRev-100] RACG 接线检查表 3E4WY 条件13（相序配置错误）比较口径裁定
**时间：** 2026-07
**决策：** 条件13「ABC/ACB 相序配置错误」判据"相序检查方法输出 ≠ 配置相序"中，**输出 2（未判定）不参与比较，不报条件13**；只有算法明确判出 0/1 且与配置相序不符才上报。推论：①单相相位错误（条件11/12 触发，此时算法输出 2）不伴报条件13；②幅值不对称>10%（相序检查方法 Step2 失败，输出 2）不报任何相序相关告警；③条件13 仅能经"源实际相序与配置相反（算法判出 0/1 且 ≠ 配置）"触发，且该场景下相位判定窗口随配置切换，源相序与配置相反必导致 ∠Vb/∠Vc 落入错误窗口，故**条件13 必伴条件11/12 同报，无法单独触发**。
**原因：** 用户于 2026-07-14 裁定该比较口径，消除"输出2 是否算不符"的歧义；用户提到开发有书面解释，文件暂未拿到，拿到后补入 `requirements/raw/`。同时修正 context.md 必读区中一处与 RACG 不符的旧表述（"相序配置错误判据幅值不对称度 ≥145%"——该值是 IIV3 多回路老模型判据，RACG 已改用相序检查方法输出比较）。
**待办：** ~~待拿到开发书面解释文件后补入 `knowledge/meters/AcuRev100/requirements/raw/`~~ 已归档（2026-07-14）：`requirements/raw/RACG_相序配置错误bit8判定_开发解释_20260714.png`，实现细节（bit8/WIRING_VERR_SEQ、0x3002/0x3010、前置门三条件）已补入 `RACG_接线检查表.md`。
**影响设备：** → [RACG_接线检查表.md](../meters/AcuRev100/requirements/summaries/RACG_接线检查表.md)、[context.md](../meters/AcuRev100/context.md)
