# 测试组知识库

## 项目一览
| 项目      | 类型  | 描述                                | 详情                                  |
|---------|-----|-----------------------------------|-------------------------------------|
| web2    | 网关  | AcuRev-4100-WEB2 网络扩展模块（ACM-41-WEB2）测试，附加在表本体，支持下挂 3 台 4100 + 8 台 AcuIOM，提供 BACnet/IP、Modbus TCP、MQTT、SNMP、Ethernet/IP 北向协议及 AcuCloud/AWS IoT/Azure IoT 数据上传 | knowledge/gateway/web2/context.md   |
| hmi1-7  | 网关  | AcuHMI-1-7 独立工业物联网网关测试，支持下挂多型号电表，提供 BACnet/IP、Modbus TCP、SNMP、MQTT 北向协议及 AcuCloud/AWS IoT/Azure IoT 云端上传，含接线检查功能 | knowledge/gateway/hmi1-7/context.md |
| [电表项目]  | 电表  | 各电表独立测试项目                         | knowledge/meters/                   |
| [云平台项目] | 云平台 | AcuCloud 平台测试                     | knowledge/cloud/                    |

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

## 技能（斜杠命令）
| 命令           | 用途              |
|--------------|-----------------|
| /new-device  | 在当前项目适配新电表设备    |
| /add-bug     | 将 bug 记录追加到项目索引 |
| /new-module  | 在当前项目新增功能模块     |
| /new-project | 初始化新项目目录结构      |

## 高频约定（完整版见 knowledge/shared/conventions.md）
1. 新增设备必须同时实现 `build_param_map()` 和 `build_cloud_col_map()`
2. BACnet比对容差：±1% / ±0.05；AcuCloud比对容差：±5% / ±1.0
3. bug 精简索引记入各项目 bugs/INDEX.md，原始 Jira Excel 存 bugs/raw/
4. 需求原件存 requirements/raw/，摘要存 requirements/summaries/
5. 模板参数范围必须按协议列筛选：BACnet→主模板`BACnetIP`列，MQTT→`MQTT`列，SNMP→`SNMP`列，DataLog→`DataLog`列；**AcuCloud 范围禁止用主模板的`AcuCloud`列**，必须用 `get_cloud_acucloud_params()` 读取 `template/AcuCloud 模板适配/` 下各设备文件的 `paramType_AcuCloud` 列（4100 暂无该列，自动回退主模板）；禁止用 `load_template()` 全量加载作为某一协议的范围；**AcuIOM 设备模板无 `BACnetIP` 列**，必须用 `get_bacnet_params_by_range()` 按 `range` 列过滤（IOM-01/02 传 `"8"`，IOM-03/04 传 `"10"`），通过 `config.BACNET_RANGE_MARKER` 和 `comparator.py` 的 `_RANGE_MARKER_MAP` 自动切换
6. Protocols/ 目录的脚本统一从仓库根目录执行：`python Protocols/BACnetIP/comparator.py`；README 示例命令须包含 `Protocols/` 前缀
7. 适配新设备前必须先列出 Modbus 地址表 Excel 的**全部 sheet 名称**，逐一确认哪些包含可读寄存器后再写设备文件，禁止仅凭部分 sheet 推断全量参数

知识库维护说明 → knowledge/CONTRIBUTING.md

## 在线文档集成（可选，待管理员授权）

当前以本地知识库为主。Jira 直连（Atlassian MCP）和 SharePoint 直连（Microsoft 365 MCP）均需管理员授权，授权前使用本地替代：Jira 导出 Excel 存 `bugs/raw/`，SharePoint 文件下载到 `knowledge/` 目录。
