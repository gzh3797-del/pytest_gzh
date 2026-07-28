# RPP · Pass Through 自动化测试

由 `projects/AcuHMI_1_7/tests/Pass Through` 适配而来。验证 RPP 网关「Pass Through
透传」：主站以透传 SlaveID（101–247）经网关直达下游电表寄存器。

当前对接 **RPP demo**（假数据前端）：`http://192.168.2.94:3030`。
页面与 Device Mirror 同构（Enable 单选 + 六列设备表格 + Save，无 Download All）。

**配置来源（与 1.7 相同的加载链，单文件自包含）**：优先 `tests/config.py`（本地适配层，
gitignored）→ 回退框架分层配置（`configs/env` + `projects/RPP/config.yaml`）→ 回退
demo 默认值；同名环境变量（`RPP_URL` 等）随时可覆盖。`HEADED=1` 有头模式。

## 运行

```powershell
cd C:\JrJ\ai_auto
python -m pytest "projects/RPP/tests/Pass Through" -v
```

报告自动生成在本目录「用例执行结果.html」。

## 用例清单（10 条）

| 用例 | 函数 | demo 现状 |
|---|---|---|
| case00 页面布局（新增） | test_pt_000_page_layout | ✅ 通过 |
| 前置检查 配置 | test_pt_001_config | ⏭ demo 透传表格无设备行（"No Data"） |
| case05(子) 透传取数 | test_pt_002_data_collected | ⏭ 同上 + 无 Modbus 服务 |
| case05 透传↔直读一致 | test_pt_003_passthrough_matches_direct | ⏭ B 路来源页面待真机确认（二期） |
| case01 Enable/Disable | test_pt_case01_enable_disable_toggle | ✅ UI 持久化通过；透传阻断/恢复待真机 |
| case02 有效 SlaveID | test_pt_case02_valid_slaveid_save | ⏭ demo 无可编辑设备行 |
| case03 SlaveID 边界(100/248 拒, 101/247 收) | test_pt_case03_slaveid_boundary | ⏭ 同上 |
| case04 重复 SlaveID | test_pt_case04_duplicate_slaveid | ⏭ 需 ≥2 台可编辑设备 |
| case06 关闭后阻断访问 | test_pt_case06_disabled_blocks_access | ⏭ 需可透传读通的设备 |
| case07 多主站并发 | test_pt_case07_concurrent_masters | ⏭ 同上 |

跳过条件均为运行时自动探测，真机到位后自动启用（同 Device Mirror 套件，见其 README）。

## 与 1.7 的差异（已在测试文件内适配）

- 登录跳 overview、菜单导航替代 hash 直跳、Save 无改动时 disabled、Disable 隐藏表格、
  锁定行排除、demo 错误 toast 过滤（与 Device Mirror 套件相同的适配点）。
- **Modbus 总开关**（Modbus Config → Modbus Enable）关闭时本页整体不可用——
  baseline 前置 `_ensure_modbus_config_enabled()` 自动检查并启用。
- A 路寄存器表沿用 1.7 机制：`knowledge/shared/templates/raw/<Model>*.xlsx`（AcuCloud 模板，
  blockParams sheet，按表头定位列）。RPP 下挂设备型号确定后如缺表在该目录补充。
- demo 实测：透传表格与 Device Mirror 同六列表头；RPP 本机不出现在透传列表（本机仅参与镜像）。
