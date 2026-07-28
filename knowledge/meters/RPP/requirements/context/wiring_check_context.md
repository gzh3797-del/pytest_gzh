# RPP · Wiring Check 模块 UI 选择器沉淀

> 对应测试目录：`projects/RPP/tests/Wiring_check/`（迁移自 AcuHMI-1-7 同名目录）
> 通用信息（Web 地址、登录、顶级导航组、协议路由总表）见 [../../context.md](../../context.md#ui-选择器沉淀) 「UI 选择器沉淀」节，本文件只记差异点。
> 数据来源：RPP 纯前端联调环境 http://192.168.2.94:3030 实测。

## 导航路径

顶级组 **Maintenance** → 左侧 **Diagnostics**（可展开）→ **Wiring Check** 子项。
路由 `/maintenance/diagnostics/wiringCheck`（HMI 路由为 Settings 侧栏下的 `/diagnostics/wiringCheck`，RPP 挂在 Maintenance 组下且多一段 `diagnostics` 前缀）。

## 与 HMI 迁移用例的适配点

- `navigate()` 助手函数需改为：先 `page.goto()` 到 Maintenance 落地页 `/maintenance/systemStatus` 以激活该顶级导航组，再 `page.goto()` 目标 URL `/maintenance/diagnostics/wiringCheck`；
  兜底路径为点击 **Maintenance → Diagnostics → Wiring Check**（URL 直达失败时的降级方案）。
- 页面地址协议为 **http**（非 https）且**带端口 3030**，硬编码 base_url 的用例需同步改。

## 落地页 landmark（已验证）

- 文本 `"Wiring Check"`
- `"Device"` 下拉（选择待测设备）

## 未验证 / 差异提醒

- `"Nominal Voltage"` 是点击 **Wiring Check** 按钮后弹出的**确认弹窗**内容，**不在落地页出现**；若迁移用例把它当作落地页 landmark 会误判页面未加载，需改为触发弹窗后再断言。
