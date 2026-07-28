# RPP · Device Mirror 自动化测试

由 `projects/AcuHMI_1_7/tests/Device Mirror` 适配而来。验证 RPP 网关「Device Mirror
设备镜像」：对外用分配的 SlaveID 提供 Modbus TCP 访问，A(镜像)↔B(直读电表) 数据一致。

当前对接 **RPP demo**（假数据前端）：`http://192.168.2.94:3030`。

**配置来源（与 1.7 相同的加载链，单文件自包含）**：优先 `tests/config.py`（本地适配层，
gitignored，可定义 `RPP_URL / RPP_USERNAME / RPP_PASSWORD / RPP_DEMO`）→ 回退框架分层配置
（`configs/env` + `projects/RPP/config.yaml` 的 `rpp_url` 等键）→ 回退 demo 默认值；
同名环境变量（`RPP_URL` 等）随时可覆盖。

## 运行

```powershell
cd C:\JrJ\ai_auto
python -m pytest "projects/RPP/tests/Device Mirror" -v

# 有头模式（看浏览器操作过程）
$env:HEADED="1"; python -m pytest "projects/RPP/tests/Device Mirror" -v; Remove-Item Env:HEADED
```

报告自动生成在本目录「用例执行结果.html」。

## 用例清单（10 条）

| 用例 | 函数 | demo 现状 |
|---|---|---|
| case00 页面布局（新增） | test_dm_000_page_layout | ✅ 通过 |
| case06 配置+导出 | test_dm_001_config_and_export | ⏭ 下载动作已验证；CSV 是假文件，解析跳过 |
| case05(子) 网关读取 | test_dm_002_modbus_read | ⏭ demo 无 Modbus 服务 |
| case05 镜像↔直读一致 | test_dm_003_mirror_matches_direct | ⏭ B 路来源页面待真机确认（二期） |
| case01 Enable/Disable | test_dm_case01_enable_disable_toggle | ✅ UI 持久化通过；供数停止/恢复待真机 |
| case02 有效 SlaveID | test_dm_case02_valid_slaveid_save | ⏭ demo 无可编辑设备行 |
| case03 SlaveID 边界 | test_dm_case03_slaveid_boundary | ⏭ 同上 |
| case04 重复 SlaveID | test_dm_case04_duplicate_slaveid | ⏭ 需 ≥2 台可编辑设备 |
| case07 离线稳定性 | test_dm_case07_offline_stability | ⏭ 人工执行 |
| case08 多主站并发 | test_dm_case08_concurrent_masters | ⏭ demo 无 Modbus 服务 |

跳过条件都是运行时自动探测（可编辑行 / CSV 有效性 / Modbus 可达性），
**真机到位后无需改用例**，条件满足即自动启用；只需：

1. 指向真机：环境变量 `RPP_URL=真机地址`、`RPP_DEMO=0`（或写进 `tests/config.py` /
   `projects/RPP/config.yaml` 的 `rpp_url` / `rpp_demo` 键）；
2. dm_003 的 B 路（直读真实电表的 IP/ModbusID 来源页面）确认后，从 1.7 套件移植
   `PhysicalDevices` + `comparison` fixture + `write_xlsx` 报告——RPP 的对应物大概率是
   Monitoring → Gateway Devices（路由 `/#/physicalDevices`，有 Download List 按钮）。

## RPP 与 AcuHMI 1.7 的页面差异（2026-07-03 demo 实测，已在代码中适配）

| 差异 | 适配 |
|---|---|
| 登录后落 `/#/overview`（1.7 是 dashboard） | `_login()` 等待 URL 改为 overview |
| 路由守卫拦 hash 直跳 | `MirrorPage.goto()` 走菜单：Settings → Protocols → Modbus 子菜单 → Device Mirror（路由 `logicalParameterMapping`） |
| Save 无改动时 disabled | `save()` 先判断按钮状态，无改动视为无需保存 |
| Disable 保存后表格+Download All 消失 | 读表格前 `ensure_enabled()` |
| 本机行(SlaveID=1)锁定且 Device Name 可能为空 | 按勾选框 `is-disabled` 排除本机，不按设备名 |
| demo 后台随机弹 "There's been an error." toast | `save_and_result()` 扫描全部消息、优先识别 success 类 |
| **Modbus 总开关**（Modbus Config → Modbus Enable）关闭时本页整体不可用（提示 Configuration unavailable） | baseline 前置 `_ensure_modbus_config_enabled()` 自动检查并启用 |

列位置（勾选/SlaveID/Device Name/Interface/Model）与 1.7 相同，COL_* 常量未改。
本套件为**单文件自包含**（与 1.7 相同）：Modbus / 登录 / 页面驱动全部内联在测试文件中，
browser 用 pytest-playwright 内置共享实例，`HEADED=1` 有头模式。
