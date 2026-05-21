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

## [web2] AcuIOM 模板格式不同，范围检查 graceful skip
**时间：** 2026-05
**决策：** get_bacnet_params() 在 IOM 模板上会抛 ValueError（无 BACnetIP 列），用 try/except 捕获并 log warning，跳过范围检查继续数值比对。
**原因：** IOM 模板由不同团队维护，格式规范不统一，短期内无法修改模板文件。
