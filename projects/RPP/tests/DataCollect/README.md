# RPP · DataCollect（Metering 采集与比对）自动化测试

由 `projects/AcuHMI_1_7/tests/DataCollect` 适配而来。1.7 的三步流水线
（采集网页 Metering 显示值 → 匹配寄存器模板 → Modbus 直读比对，断言稳定量 FAIL=0）
在 RPP 上的采集对象改为 **RPP 自身面板**：

- 1.7：Physical Devices → 逐台下挂设备 → Metering 各视图
- RPP：Monitoring → **RPP Panel** → Metering（Realtime / Energy / Demand /
  Power Quality / Max Demand / Min/Max），按 **VMM × Meter Point(channel)** 下拉遍历

当前对接 **RPP demo**：`http://192.168.2.94:3030`。

**配置来源（与 1.7 相同的加载链，单文件自包含）**：优先 `tests/config.py`（本地适配层，
gitignored）→ 回退框架分层配置（`configs/env` + `projects/RPP/config.yaml`）→ 回退
demo 默认值；同名环境变量（`RPP_URL` 等）随时可覆盖。`HEADED=1` 有头模式。

## 运行

```powershell
cd C:\JrJ\ai_auto
python -m pytest "projects/RPP/tests/DataCollect" -v
```

## 用例清单（4 条）

| 用例 | 函数 | demo 现状 |
|---|---|---|
| case00 页面布局（新增） | test_dc_000_page_layout | ✅ 通过（六视图菜单 + VMM/Update Rate/Meter Point 下拉齐备） |
| Step1 采集 | test_dc_001_collect_json | ⏭ demo VMM 下拉无选项（No Data）；采集代码已就绪，真机/完整 demo 数据即自动启用 |
| Step2 寄存器匹配 | test_dc_002_register_match | ⏭ 待 RPP 寄存器模板（blockParams xlsx），二期按 1.7 metering.py 移植 |
| Step3 Modbus 比对 | test_dc_003_compare_modbus | ⏭ 待真机 Modbus + 模板，二期移植（含稳定量/波动量分类判定） |

## 真机到位后需要补的两样东西

1. **RPP 寄存器模板**：放到 `knowledge/shared/templates/raw/`（blockParams sheet 格式，
   与其他型号一致），Step2/Step3 才能移植。
2. **确认 Modbus 直读路径**：读 RPP 自身用镜像 SlaveID=1（Device Mirror 本机行），
   还是有独立的 native Modbus 地址表——影响 compare 的 unit 取值。

## demo 探测记录（2026-07-03）

- RPP Panel 左侧菜单：Metering（六视图）/ Alarms / Logs / PQ Event / Trend Log / Waveform。
- 注意 Logs 下也有 Realtime/Energy 同名菜单项——`open_view()` 取"首个可见"项，
  已通过先展开 Metering 子菜单来保证命中 Metering 组。
- Monitoring 二级导航：Overview / RPP Panel / Gateway Devices(=/#/physicalDevices，
  即 1.7 Physical Devices 的对应物，demo 空表) / Virtual Devices / Web Devices / Alarm / Data Log。
  → 下挂设备接入后，Device Mirror/Pass Through 二期的 B 路（真实电表 IP/ModbusID）
  大概率从 Gateway Devices 页抓取，页面还有 Download List 按钮可用。
