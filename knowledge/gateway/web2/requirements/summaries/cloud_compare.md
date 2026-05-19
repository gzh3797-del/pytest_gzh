# AcuCloud比对模块需求摘要

原始文档：requirements/raw/（待存入）

## 核心功能
1. **范围检查**：模板 AcuCloud 列参数 vs 设备 build_cloud_col_map() 覆盖情况
2. **数值比对**：读取 xlsx 快照某行数据，与实时 Modbus 值对比

## 数据来源
- 快照：AcuCloud 导出的 xlsx 文件，存放于 Acuclouddatas/
- 文件命名规则：<设备名>.xlsx（如 AcuRev4100.xlsx），大小写须精确匹配
- 默认取最新行，可用 --row 指定

## 容差规则
```
pass if |diff| <= max(CLOUD_TOLERANCE_ABSOLUTE, ref × CLOUD_TOLERANCE_PERCENT / 100)
默认：CLOUD_TOLERANCE_PERCENT=5.0%，CLOUD_TOLERANCE_ABSOLUTE=1.0
（比 BACnet 比对宽松，补偿时序差异）
```

## 报告要求
- 格式：HTML，两段可折叠（范围检查 / 数值比对）
- 输出路径：reports/cloud_<设备名>_<时间戳>.html

## 命令行接口
```bash
python cloud_comparator.py [--device <名>] [--file <xlsx路径>] [--row <行号>] [--keys KEY1 KEY2...]
```

## 支持设备
AcuRev4100、AcuRev2100、AcuvimIIW、AcuvimIIR、AcuVIM3
IOM系列不支持 AcuCloud 比对。
