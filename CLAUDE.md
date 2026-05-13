# 测试组知识库

## 项目一览
| 项目      | 类型  | 描述                                | 详情                                  |
|---------|-----|-----------------------------------|-------------------------------------|
| web2    | 网关  | 网关自动化测试（BACnet/Modbus/AcuCloud比对） | knowledge/gateway/web2/context.md   |
| hmi1-7  | 网关  | HMI 界面自动化测试                       | knowledge/gateway/hmi1-7/context.md |
| [电表项目]  | 电表  | 各电表独立测试项目                         | knowledge/meters/                   |
| [云平台项目] | 云平台 | AcuCloud 平台测试                     | knowledge/cloud/                    |

## 支持设备速查
| 型号                 | FC   | 典型端口 | BACnet | AcuCloud | 备注          |
|--------------------|------|------|--------|----------|-------------|
| AcuRev4100         | FC03 | 502  | ✅      | ✅        | 端口非标准       |
| AcuRev2100         | FC03 | 502  | ✅      | ✅        |             |
| AcuvimIIW          | FC03 | 502  | ✅      | ✅        |             |
| AcuvimIIR          | FC03 | 502  | ✅      | ✅        | 地址表与IIW完全相同 |
| AcuVIM3            | FC03 | 502  | ✅      | ✅        |             |
| AcuRev1300（PXM350） | FC03 | 502  | ✅      | —        |             |
| AcuIOM-01          | FC03 | 502  | ✅      | —        | 8 AI通道      |
| AcuIOM-02          | FC03 | 502  | ✅      | —        | 16 AI通道     |
| AcuIOM-03          | FC02 | —    | ⚠️     | —        | DI型号，暂不支持   |
| AcuIOM-04          | FC02 | —    | ⚠️     | —        | DI型号，暂不支持   |

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
5. 模板参数范围必须按协议列筛选：BACnet→`BACnetIP`列，AcuCloud→`AcuCloud`列，MQTT→`MQTT`列，SNMP→`SNMP`列；禁止用 `load_template()` 全量加载作为某一协议的范围
6. Protocols/ 目录的脚本统一从仓库根目录执行：`python Protocols/BACnetIP/comparator.py`；README 示例命令须包含 `Protocols/` 前缀

知识库维护说明 → knowledge/CONTRIBUTING.md
