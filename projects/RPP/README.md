# RPP 测试项目

本文件只做**模块索引 + 运行命令**；各模块的详细说明（用例清单、前置条件、报告）见对应模块目录下的独立 README。

- **配置**：`configs/global.yaml`（全局）+ 本项目 `config.yaml`（项目差异、显示名）+ `configs/.env`（敏感值，模板见 `configs/.env.example`）
- **报告**：`reports/RPP/<时间戳>/`
- `pages/` Page Object ｜ `helpers/` 项目内客户端/匹配器 ｜ `data/` 测试数据/模板

## 模块

| 模块 | 路径 | 说明 | 详细 README |
|---|---|---|---|
| 需量精度 | `tests/demand_test/` | 需量精度测试（控源 + Modbus 回读比对），由 AcuRev1320/demand_test 迁移 | `tests/demand_test/README.md` |
| BACnet/IP | `tests/BacnetIP/` | 北向协议测试（UI 配置 + EPICS 可读性 + 六段值比对），由 AcuHMI_1_7/BacnetIP 迁移 | `tests/BacnetIP/README.md` |
| 接线检查 | `tests/Wiring_check/` | 电压/电流信号驱动 + Wiring Status 推导核验，由 AcuHMI_1_7/Wiring_check 迁移 | `tests/Wiring_check/README.md` |
| AWS IoT | `tests/aws_iot/` | AWS IoT Core 云端上传测试（Playwright UI 配置 + MQTT 订阅校验），由 AcuHMI_1_7/aws_iot 迁移 | `tests/aws_iot/README.md` |
| Data Log | `tests/data_log/` | 本地数据记录测试（FTP/SFTP/HTTP/HTTPS Post + Modbus 回读比对），由 AcuHMI_1_7/data_log 迁移 | `tests/data_log/README.md` |
| MQTT | `tests/mqtt/` | 从 AcuHMI-1-7 迁移过来的骨架，外层配置（settings.py/config.yaml/conftest.py）已指向 RPP 真机，**用例主体（页面选择器/断言）尚未对照 RPP 真机适配，暂不可直接跑通** | [`tests/mqtt/README.md`](tests/mqtt/README.md) |
| Acuview2 自动化 | `tests/acuview_auto/` | AcuRev1320 相关用例 | 待补 |
| 快速精度 | `Quick_Accuracy_Test/` | 快速精度测试工具 | — |
| 脉冲灯精度 | `Pulse Light Accuracy Test/` | 脉冲灯精度测试（SDG 自动测试） | — |

> ⚠️ **迁移说明**：`BacnetIP / Wiring_check / aws_iot / data_log` 四模块自 AcuHMI_1_7 迁移，
> 已完成包路径改写（`projects.AcuHMI_1_7.*` → `projects.RPP.*`）+ 结构落位 + 收集冒烟（`--collect-only` 零错误）。
> **尚未做 RPP 设备级适配**：模块内的设备逻辑、`config.yaml`（网关 IP/表 IP/AWS 端点/证书）、
> Web 页面选择器、Modbus 地址表、接线类型、用例编号（仍为 `AcuHMI-1-7_*`）均沿用 AcuHMI，
> 对真实 RPP 台架执行前需按 RPP 硬件重新适配。

`tests/` 下其余子目录（SNMP / azure_iot / parameter_settings / DataCollect / Device Mirror / Pass Through / ui 等）为迁移骨架或空占位，尚未完成 RPP 适配。

## 真机信息

- Web UI：`http://192.168.2.94:3030`（HTTP，非 HTTPS，端口 3030）
- MQTT 配置页路由：`/#/protocols/mqtt/deviceToPublish`

## 运行命令（仓库根目录）

```bash
# 用例收集校验（不需要硬件）
pytest projects/RPP/tests/mqtt/test_mqtt.py --collect-only -q   # MQTT（暂不建议直接执行）
pytest projects/RPP/tests/BacnetIP projects/RPP/tests/Wiring_check \
       projects/RPP/tests/aws_iot projects/RPP/tests/data_log --collect-only -q   # 迁移四模块共 260 条

# 单模块示例（需 RPP 台架在环 + 完成设备级适配后再执行）
pytest projects/RPP/tests/BacnetIP/ -v --html=reports/bacnet_ui.html --self-contained-html
pytest projects/RPP/tests/Wiring_check/ -v -s
pytest projects/RPP/tests/aws_iot/ -m aws_iot -v
pytest projects/RPP/tests/data_log/ -v
```

需量测试（demand_test）完整说明（三维度参数、用例数据列约定、结果产物）见 `tests/demand_test/README.md`。
