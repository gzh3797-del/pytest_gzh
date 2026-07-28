# RPP · BACnet/IP 模块 UI 选择器沉淀

> 对应测试目录：`projects/RPP/tests/BacnetIP/`（迁移自 AcuHMI-1-7 同名目录）
> 通用信息（Web 地址、登录、顶级导航组、协议路由总表）见 [../../context.md](../../context.md#ui-选择器沉淀) 「UI 选择器沉淀」节，本文件只记差异点。
> 数据来源：RPP 纯前端联调环境 http://192.168.2.94:3030 实测（无后端，设备数据类接口报错属正常）。

## 导航路径

顶级组 **Settings** → 左侧 **Protocols** → **BACnet/IP**，路由 `/protocols/bacnet`（与 HMI 相同）。

## 与 HMI 迁移用例的适配点

- HMI 原用例第 1 级导航靠点击含 `'AcuHMI'` / `'AcuHMI-1-7'` 文案的 nav-item；RPP 无此文案，改为**精确匹配 `'Settings'`**：
  ```python
  page.locator(".nav-item, .nav-item-menu, header span").filter(
      has_text=re.compile(r"^Settings$")
  ).first.click()
  ```
  用精确匹配（`textContent.trim() == 'Settings'`）而非包含匹配，避免命中 `'System Settings'`。
- Protocols / BACnet/IP 子级选择器与 HMI 完全相同，无需改动。

## 落地页 landmark（已验证）

- `label "BACnet Enable"` + Enable/Disable 单选 + `button "Save"`。
- BACnet Port / COV 等具体配置字段依赖 Enable 后的动态渲染或后端回读数据，当前无后端环境下不出现，**不作为该模块可用性判据**，仅用上述 landmark 判定页面加载成功。
