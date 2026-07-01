# 参数模板索引

模板文件分两套，由 `template_reader.py` 解析，Claude 参考本文件了解结构，无需读原始 Excel。

---

## 一、主模板（BACnet / Modbus / MQTT / SNMP / DataLog 范围基准）

存放于 `raw/`，命名规则：`<设备名>_v<版本>_<日期>.xlsx`

关键列：
- `paramType`  — 参数标识（param_key）
- `BACnetIP`   — 非空 → BACnet/IP 比对范围
- `MQTT`       — 非空 → MQTT 比对范围
- `SNMP`       — 非空 → SNMP 比对范围
- `DataLog`    — 非空 → Datalog 比对范围
- `AcuCloud`   — ⚠️ **范围不准，禁止用于 AcuCloud 比对**，仅供参考

| 设备 | 文件（raw/） | BACnetIP | DataLog/MQTT/SNMP |
|------|------------|----------|-------------------|
| AcuRev4100 | AcuRev-4100_v1.01_20260427.xlsx | 1869 | 1059 |
| AcuRev2100 | AcuRev-2100_v1.01_20260416.xlsx | 1225 | 595 |
| AcuRev1300 | AcuRev-1300_v1.01_20260416.xlsx | 39 | 39 |
| AcuvimIIW | AcuvimIIW_v1.01_20260509.xlsx | 478 | 106 |
| AcuvimIIR | AcuvimIIR_v1.01_20260509.xlsx | 478 | 106 |
| AcuVIM3 | Acuvim3_v1.01_20260416.xlsx | — | — |
| AcuIOM-01 | AcuIOM-01_v1.01_20250205.xlsx | ⚠️ 无标准列 | — |
| AcuIOM-02 | AcuIOM-02_v1.01_20250205.xlsx | ⚠️ 无标准列 | — |
| AcuIOM-03 | AcuIOM-03_v1.01_20250205.xlsx | — | — |
| AcuIOM-04 | AcuIOM-04_v1.01_20250205.xlsx | — | — |

---

## 二、AcuCloud 专用模板（AcuCloud 比对范围唯一基准）

存放于 `raw/AcuCloud 模板适配/`，由 `get_cloud_acucloud_params()` 读取。

关键列：`paramType_AcuCloud` — 各设备 AcuCloud 实际比对参数标识。

| 设备 | 文件 | AcuCloud 参数数 | 备注 |
|------|------|----------------|------|
| AcuRev4100 | — | 1059（回退主模板） | 专用模板暂无 paramType_AcuCloud 列，自动回退 |
| AcuRev2100 | AcuRev-2100.xlsx | 341 | |
| AcuvimIIW | AcuvimIIW.xlsx | 106 | |
| AcuvimIIR | AcuvimIIR.xlsx | 106 | |
| AcuVIM3 | Acuvim3.xlsx | 125 | |
| AcuRev1300 | AcuRev-1300.xlsx | 33 | |

**代码调用说明：**
```python
# ✅ 正确：AcuCloud 比对范围
from template_reader import find_template_file, get_cloud_acucloud_params
path   = find_template_file(config.ACUCLOUD_TEMPLATE_DIR, device_name)
params = get_cloud_acucloud_params(path)   # 自动回退逻辑在 cloud_comparator 中

# ❌ 错误：不要用这个作为 AcuCloud 范围
from template_reader import get_cloud_params   # AcuCloud 列范围不准
```

---

## 已知格式问题
- AcuIOM 系列：主模板无 BACnetIP / AcuCloud 列，代码用 try/except 跳过范围检查
- AcuRev-4100 AcuCloud 模板：专用文件中使用第二个 paramType 列而非 paramType_AcuCloud，当前代码对此不处理，自动回退主模板（1059 个参数）

## 更新说明
新增设备模板后，同时在两套模板中更新，并在本表追加对应行。
