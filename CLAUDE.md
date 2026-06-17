# 测试组知识库

## 项目一览
| 项目      | 类型  | 描述                                | 详情                                  |
|---------|-----|-----------------------------------|-------------------------------------|
| ACM-41-WEB2 | 网关  | AcuRev-4100-WEB2 网络扩展模块测试，附加在表本体，支持下挂 3 台 4100 + 8 台 AcuIOM，提供 BACnet/IP、Modbus TCP、MQTT、SNMP、Ethernet/IP 北向协议及 AcuCloud/AWS IoT/Azure IoT 数据上传 | knowledge/gateway/web2/context.md   |
| AcuHMI-1-7  | 网关  | AcuHMI-1-7 独立工业物联网网关测试，支持下挂多型号电表，提供 BACnet/IP、Modbus TCP、SNMP、MQTT 北向协议及 AcuCloud/AWS IoT/Azure IoT 云端上传，含接线检查功能 | knowledge/gateway/hmi1-7/context.md |
| AcuRev-4100 | 电表  | AcuRev-4100 多回路交流电表测试，24路电流输入，12个用户通道，支持 Modbus TCP/RTU + BACnet MS/TP，可搭配 ACM-41-WEB2 扩展北向协议 | knowledge/meters/AcuRev4100/context.md |
| AcuDC-320   | 电表  | AcuDC-320 直流电能表测试，面向 EV 快充充电桩，支持 OCMF 交易日志与 ECDSA 签名、Echilog 法定计量事件日志，MID + UL 认证，配分立显示模块（AcuDC-RDU），通信仅 RS485 Modbus RTU | knowledge/meters/AcuDC320/context.md   |
| AcuRev-1320 | 电表  | AcuRev-1320 三相电表测试，支持 Modbus TCP/RTU + BACnet MS/TP/IP，含电压闪变、Independent Input Channel、Dual Source Energy、AcuCloud、Data Post、HTTPs Web 页面等新功能；型号分 1321（精简）和 1322（全功能） | knowledge/meters/AcuRev1320/context.md |
| RPP         | 电表  | RPP（Remote Power Panel）远程电力面板测量系统，MH 主控通过 AccuBus 连接最多 2 VMM + 4 CMM（96 路电流通道），最多 96（基线）/192（期望）个 Meter Point；北向支持 Modbus TCP/RTU、SNMP、BACnet/IP、MQTT、AcuCloud；兼具 Gateway 功能（≥32台 Modbus 下挂设备） | knowledge/meters/RPP/context.md     |
| AcuCloud    | 云平台 | AcuCloud 云端能源管理平台（EMS SaaS）测试，支持电力/水/天然气三种能源类型，覆盖 Installation/Dashboard/Analysis/Billing/Power Quality/Carbon Model/Report/Data 全模块，订阅 Plan 分 Free/Lite/AcuBilling/AcuPQ/Plus/AcuEMS 六级 | knowledge/cloud/context.md          |

## 支持设备速查
| 型号                 | FC   | 典型端口 | BACnet | EtherNet/IP | Modbus Mapping | SNMP | MQTT | DataLog | AcuCloud | AWS IoT | Azure IoT | Device Mirror | 备注 |
|--------------------|------|------|:------:|:-----------:|:--------------:|:----:|:----:|:-------:|:--------:|:-------:|:---------:|:-------------:|------|
| AcuRev4100         | FC03 | 502  | ✅     | ✅           | —              | —    | ✅   | ✅       | ✅        | —       | —         | —             | 端口非标准 |
| AcuRev2100         | FC03 | 502  | ✅     | —           | —              | —    | ✅   | ✅       | ✅        | —       | —         | —             | |
| AcuvimIIW          | FC03 | 502  | ✅     | —           | —              | —    | ✅   | ✅       | ✅        | —       | —         | —             | |
| AcuvimIIR          | FC03 | 502  | ✅     | —           | —              | —    | ✅   | ✅       | ✅        | —       | —         | —             | 地址表与IIW完全相同 |
| AcuVIM3            | FC03 | 502  | ✅     | —           | —              | —    | ✅   | ✅       | ✅        | —       | —         | —             | |
| AcuRev1300（PXM350） | FC03 | 502  | ✅     | —           | —              | —    | —    | —        | —         | —       | —         | —             | |
| AcuIOM-01          | FC03 | 502  | ✅     | ✅           | —              | —    | —    | —        | —         | —       | —         | —             | 8 AI通道，原始输入+物理量读数各8个 |
| AcuIOM-02          | FC03 | 502  | ✅     | ✅           | —              | —    | —    | —        | —         | —       | —         | —             | 16 AI通道，原始输入+物理量读数各16个 |
| AcuIOM-03          | FC02 | 502  | ✅     | ✅           | —              | —    | —    | —        | —         | —       | —         | —             | 14 DI通道，BACnet BI对象 |
| AcuIOM-04          | FC02 | 502  | ✅     | ✅           | —              | —    | —    | —        | —         | —       | —         | —             | 28 DI通道，BACnet BI对象 |

完整设备知识 → knowledge/shared/devices/INDEX.md

## 共享资源
- 设备知识库 → knowledge/shared/devices/INDEX.md
- Modbus 地址表索引 → knowledge/shared/modbus_tables/INDEX.md
- 参数模板索引 → knowledge/shared/templates/INDEX.md
- 团队编码约定 → knowledge/shared/conventions.md
- 历史决策记录 → knowledge/shared/decisions.md
- 测试设计方法学（strategy-design / testcase-design 共用） → knowledge/shared/test_design_methods.md

## 技能（斜杠命令）
| 命令                 | 用途                                          |
|--------------------|---------------------------------------------|
| /new-device        | 在当前项目适配新电表设备                                |
| /add-bug           | 将 bug 记录追加到项目索引                             |
| /new-module        | 在当前项目新增功能模块                                 |
| /new-project       | 初始化新项目目录结构                                  |
| /read-modbus-table | 解析 Modbus 地址表 Excel，提取地址、数据类型、功能码、缩放系数      |
| /strategy-design   | 将功能需求转化为测试策略脑图并导出 XMind 文件（兼容 XMind 26）     |
| /testcase-design   | 将测试策略/测试点转化为规范的 Excel 测试用例文件（.xlsx）         |
| /coverage-check    | 对已生成脑图/用例做查漏复核（独立重导+多基准回扫，三档分级，仅对话结论）；由 strategy-design/testcase-design 末尾自动调用，也可单独触发 |

## 高频约定（完整版见 knowledge/shared/conventions.md）
1. 新增设备必须同时实现 `build_param_map()` 和 `build_cloud_col_map()`
2. BACnet比对容差：±1% / ±0.05；AcuCloud比对容差：±5% / ±1.0
3. bug 精简索引记入各项目 bugs/INDEX.md，原始 Jira Excel 存 bugs/raw/
4. 需求原件存 requirements/raw/，摘要存 requirements/summaries/
5. 模板参数范围必须按协议列筛选：BACnet→主模板`BACnetIP`列，MQTT→`MQTT`列，SNMP→`SNMP`列，DataLog→`DataLog`列；**AcuCloud 范围禁止用主模板的`AcuCloud`列**，必须用 `get_cloud_acucloud_params()` 读取 `template/AcuCloud 模板适配/` 下各设备文件的 `paramType_AcuCloud` 列（4100 暂无该列，自动回退主模板）；禁止用 `load_template()` 全量加载作为某一协议的范围；**AcuIOM 设备模板无 `BACnetIP` 列**，必须用 `get_bacnet_params_by_range()` 按 `range` 列过滤（IOM-01/02 传 `"8"`，IOM-03/04 传 `"10"`），通过 `config.BACNET_RANGE_MARKER` 和 `comparator.py` 的 `_RANGE_MARKER_MAP` 自动切换
6. tools/Protocols/ 目录的脚本统一从仓库根目录执行：`python tools/Protocols/BACnetIP/comparator.py`；README 示例命令须包含 `tools/Protocols/` 前缀
7. 适配新设备前必须先列出 Modbus 地址表 Excel 的**全部 sheet 名称**，逐一确认哪些包含可读寄存器后再写设备文件，禁止仅凭部分 sheet 推断全量参数
8. 所有 Python 代码须符合 PyCharm 代码规范，不得出现任何告警（含 PEP 8 格式告警、未使用 import/变量、类型注解不一致），零告警是硬性要求
9. Playwright UI 测试交互优先用 `locator` + `expect` 风格（保留 auto-wait 与 actionability 检查）；仅当 Element Plus 拦截事件链（如 el-radio 合成事件）才降级坐标点击，且必须先 `scrollIntoView` 再取坐标；JS 循环点按钮须加 `offsetParent !== null` 可见性过滤；测试体内禁止 `asyncio.run()`（用 `_run_coro` 线程模式）。完整版见 conventions.md「Playwright UI 测试编码约定」

## 知识库维护快速参考

### 触发关键词 → 对应操作

| 说了什么 | 执行什么 |
|---------|---------|
| 上传 / 导入 Jira、bug 单、缺陷表、导出 CSV | 解析 ID 前缀定位项目 → 刷新 bugs/INDEX.md 统计 → 追加新 ID 行 → 原件存 bugs/raw/ |
| 上传需求、新需求来了、需求文档、SRS、变更说明 | 提取全文（docx 图片用 zipfile 解包后读取）→ 写 requirements/summaries/<版本>.md → 同步更新项目 context.md → **检查项目一览中该项目名是否为占位符或与文档中官方名不符，是则同步更新** → 判断是否改 CLAUDE.md |
| 记录决策、记一下为什么、写进 decisions | 追加到 knowledge/shared/decisions.md |
| 适配新设备、新加一台表、新增设备 | 创建 devices/<name>.md → 更新 devices/INDEX.md → 更新 templates/INDEX.md → 更新 CLAUDE.md 设备速查表 |
| 更新 context、同步项目状态、项目有新模块 | 仅更新对应项目 context.md |
| 约定变了、规范更新 | 仅更新 knowledge/shared/conventions.md |
| Bug 状态变了、关掉了几个 bug | 重新导入 Jira 导出文件走完整 Jira 流程；不接受口头修改单条状态 |
| 上传测试用例、用例文件、Excel 用例 | 存入对应项目 testcase/ → 在 testcase/ 内生成摘要 .md → 不上浮至 CLAUDE.md / context.md |
| 更新支持设备列表、同步设备表、扫描代码更新设备 | 扫描 `tools/Protocols/devices/` 下所有设备文件（FC/地址范围/build_param_map）+ 各协议 comparator 实现情况 → 重建 `tools/Protocols/README.md` 设备支持表 → 同步更新 `CLAUDE.md` 支持设备速查表 |

### CLAUDE.md 何时需要改

✅ 改：新增项目、新增设备型号、全局约定变更、**需求文档中的官方名与项目一览中的名称不符（含占位符如 `[xxx项目]`）**
❌ 不改：项目内部细节、Bug 状态、需求摘要内容、协议实现细节

### 知识库不写的内容

- 代码直接可读的（函数签名、配置项、脚本路径）
- 可从需求推导的测试用例
- 重复内容——选一个权威位置，其他地方用相对路径链接

### Bug 索引规则

只追加，不改旧条目。已有条目状态变化不回填；下次 Jira 导出时整体刷新统计概览。

### testcase 目录规则

各项目 `testcase/` 目录用于存放测试用例文件（如 .xlsx）及其摘要（.md）。
- 上传用例文件后，在 `testcase/` 内生成对应摘要 .md（模块/版本/用例数概览）
- **不上浮至 CLAUDE.md 或 context.md**：用例详情只在 testcase/ 内维护，不扩散到全局文件

### 禁止操作

- **【最高优先级】在项目任何文件夹保留 Claude 中间调试脚本**：调试/探查用的临时脚本（Playwright 探查脚本、临时验证 .py、`_tmp_*`、一次性 JSON/txt 等）必须在任务结束前全部删除，零残留（详见 knowledge/shared/decisions.md）
- 编辑 bugs/raw/ 和 requirements/raw/ 原始文件
- 将测试用例内容写入 testcase/ 以外的知识库文件（CLAUDE.md、context.md、summaries 等）

---

知识库维护说明 → knowledge/CONTRIBUTING.md

## 在线文档集成（可选，待管理员授权）

当前以本地知识库为主。Jira 直连（Atlassian MCP）和 SharePoint 直连（Microsoft 365 MCP）均需管理员授权，授权前使用本地替代：Jira 导出 Excel 存 `bugs/raw/`，SharePoint 文件下载到 `knowledge/` 目录。
