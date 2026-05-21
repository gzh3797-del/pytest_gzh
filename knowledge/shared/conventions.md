# 团队编码约定

## 通用
- 所有自动化脚本使用 Python 3.10+
- IO密集型操作使用 asyncio，不用 threading
- 配置集中：所有可调参数放 config.py，不散落在业务代码中
- 报告格式：HTML，包含可折叠分节（details/summary），无需 JS

## Modbus 读取约定
- 默认使用 FC03（Holding Registers），Float32，Big Endian，2寄存器/参数
- FC02（Discrete Inputs）暂不支持，DI型设备（IOM-03/04）创建空 stub 并注释说明
- 地址表模块命名：devices/<设备名小写>.py
- 必须实现 build_param_map() → dict[param_key, ModbusRegister]
- 支持 AcuCloud 的设备还须实现 build_cloud_col_map() → dict[xlsx列标题, param_key]

## BACnet 约定
- 只处理 Analog Input 对象（BACnet AI），BACnet Binary Input 暂不支持
- BACnet 对象名（object_name）= param_key
- 单位检查：比对 BACnet units 属性 vs 模板 unit 列（不检查 description）
- 比对容差：|diff| ≤ max(TOLERANCE_ABSOLUTE, ref × TOLERANCE_PERCENT/100)

## AcuCloud 约定
- xlsx 文件名须与设备名精确匹配（如 AcuvimIIW.xlsx）
- build_cloud_col_map() 返回 {xlsx列标题: param_key}
- AcuvimIIR 的 cloud_col_map 与 AcuvimIIW 完全相同，直接 import 复用
- 比对容差：±5% / ±1.0（补偿时序差异）

## 新增设备检查清单
- [ ] devices/<name>.py — build_param_map()
- [ ] devices/<name>.py — build_cloud_col_map()（若支持Cloud）
- [ ] config.py — MODBUS_DEVICE_MAP 添加条目
- [ ] comparator.py — _DEVICE_MAP 添加条目
- [ ] cloud_comparator.py — _DEVICE_MAP 添加条目（若支持Cloud）
- [ ] README.md — 支持设备表更新
- [ ] shared/devices/<name>.md — 设备知识文件
- [ ] shared/modbus_tables/INDEX.md — 更新索引
- [ ] shared/templates/INDEX.md — 更新索引（若有模板）

## 知识库维护约定
- 每次 Jira 导出后更新对应项目 bugs/INDEX.md（5分钟内完成）
- 新需求文档下来后，2个工作日内完成 requirements/summaries/ 摘要
- 重要决策（为什么这样设计）记入 shared/decisions.md
