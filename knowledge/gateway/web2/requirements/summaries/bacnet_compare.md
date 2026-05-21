# BACnet比对模块需求摘要

原始文档：requirements/raw/（待存入）

## 核心功能
1. **范围检查**：模板 BACnetIP 列参数 vs 网关实际发布的 AI 对象，检查缺失和多余
2. **单位检查**：读取每个 AI 对象的 BACnet units 属性，与模板 unit 列对比
3. **数值比对**：并发读取 BACnet Present Value 与 Modbus 实时值，按容差规则判通过/失败

## 容差规则
```
pass if |diff| <= max(TOLERANCE_ABSOLUTE, ref × TOLERANCE_PERCENT / 100)
其中 ref = max(|bacnet_value|, |modbus_value|)
默认：TOLERANCE_PERCENT=1.0%，TOLERANCE_ABSOLUTE=0.05
```

## 报告要求
- 格式：HTML，三段可折叠（details/summary），summary行显示彩色badge
- 输出路径：reports/compare_<设备名>_<时间戳>.html

## 命令行接口
```bash
python comparator.py [--device <名>] [--quick] [--no-meta] [--keys KEY1 KEY2...]
```

## 开放问题 / 待跟进
- IOM-03/04 DI型号支持：需实现 FC02 读取路径 + Binary Input 解析
- PXM350 模板格式确认：是否与标准格式一致
