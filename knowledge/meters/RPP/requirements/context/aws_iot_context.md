# RPP · AWS IoT 模块 UI 选择器沉淀

> 对应测试目录：`projects/RPP/tests/aws_iot/`（迁移自 AcuHMI-1-7 同名目录）
> 通用信息（Web 地址、登录、顶级导航组、协议路由总表）见 [../../context.md](../../context.md#ui-选择器沉淀) 「UI 选择器沉淀」节，本文件只记差异点。
> 数据来源：RPP 纯前端联调环境 http://192.168.2.94:3030 实测。

## 导航路径

顶级组 **Settings** → 左侧 **Protocols** → **AWS IoT**，路由 `/protocols/awsIot`。

## 与 HMI 迁移用例的适配点

- Page Object 用 `device_name` 字段定位顶级 nav-item 文案，HMI 填 `'AcuHMI-1-7'`，RPP 需改填 **`'Settings'`**。
  配置位置：`projects/RPP/tests/aws_iot/config.yaml` 的 `gateway.device_name`。
- **Virtual Devices 页面在顶级组 Monitoring 下**（HMI 中在 `'Devices'` 组下），涉及 Virtual Device 联动步骤的用例需同步改导航组文案。

## 字段名（与 HMI 一致，已验证，无需改动）

- 占位符：`Enter Client Id` / `Enter URL` / `Enter Topic`
- label：`Client Id` / `URL` / `Topic` / `Interval` / `Cert File` / `Key File` / `Devices Selection`
- 按钮：`Browse` / `Test Connection` / `Save`
- Enable/Disable 单选（el-radio）
