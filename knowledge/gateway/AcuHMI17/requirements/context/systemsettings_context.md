# AcuHMI-1-7 System Settings 模块 UI 选择器沉淀（实测事实）

> 探查日期：2026-07-03，设备 192.168.3.71，Playwright MCP 只读探查。
> 用途：写/调试 `projects/AcuHMI_1_7/tests/ui/systemsettings/` 用例前先查本文档，命中即复用，不重复现场探查。

## 子页导航与路由

| 子页名（导航文案） | hash 路由 | 加载时 GET API |
|---|---|---|
| Date & Time | `#/systemSettings/dateTime` | `/api/settings/ntpConfig`、`/api/command/getTime` |
| Network | `#/systemSettings/network` | `/api/settings/networkConfig`、`/api/command/getIp`、`/api/settings/proxyServerConfig` |
| Access Control（页面标题同名；即 Whitelist） | `#/systemSettings/whitelist` | `/api/settings/whitelistConfig`、`/api/whitelist/list` |
| Email | `#/systemSettings/email` | `/api/settings/smtpConfig` |
| Alarm Notification | `#/systemSettings/alarm` | `/api/settings/alarmEmailConfig` |
| Certificate Management | `#/systemSettings/certificateManagement` | `/api/command/getCertInfo` |
| Configuration Management（在 "...More" 下） | `#/systemSettings/configurationManagement` | 无独立新增 GET |
| Remote Access（在 "...More" 下） | `#/systemSettings/remoteAccess` | `/api/settings/deviceInfo`、`/api/settings/remoteDeviceAccess` |

- 菜单项：`page.get_by_role("menuitem", name="<Tab名>")`。
- **"...More" 是 tooltip 弹出层**（非常驻 DOM）：先 `page.get_by_role("menuitem", name="...More").click()`，再点 tooltip 内的 `menuitem "Configuration Management"` / `menuitem "Remote Access"`。用例内直接 hash 跳转（`page.goto(base + "#/systemSettings/xxx")`）可绕过。

## 关键规律：两类校验时机（写断言前必须区分）

1. **blur 即报错**（可直接 `expect(.el-form-item__error).to_be_visible()`，无需点 Save）：
   - IP/域名格式类：Network 的 DNS 1/2（实测非法值报 `DNS 1 must be a valid domain or a valid IP address`）、Whitelist 弹窗的 From/To/IP Address。
2. **仅点 Save 后异步渲染错误**（blur 不报错）：
   - 长度上限类：NTP Server（40）、Sender Name（40）、Recipient（100）、From Email Address（100）、Username（40）、Password（40）
   - 数值范围类：Email Port（1-65535，实测填 99999 blur 无错）、Alarm Email Interval（1-10）
   - 邮箱正则格式类：Recipient / From Email Address

## Element Plus 坑（实测）

- **el-radio 原生 click 会被 `.el-radio__inner` 拦截 timeout**：优先点文案 label，如 `page.get_by_text("Enable", exact=True).click()`（scope 到对应 `.el-form-item`）。
- **Time Zone 为 El Plus v2 select**（`.el-select__input`），下拉 405 个 IANA 时区项：用 `input.getAttribute('aria-controls')` 取专属 listId 后再在该 popper 下查 `.el-select-dropdown__item`，禁止全局 querySelectorAll。
- 多组同名元素需父容器 scope：Certificate 页 Issuer/Subject 各有一组同名 group（"Common Name" 等）；Network 页 Ethernet 1/2 各有 "Interface Status"/"IP"。

## 各子页控件与选择器

### Date & Time（`dateTime`）
- NTP Enable*：el-radio Enable/Disable，默认 Enable；Disable 时 NTP Server 1/2/3 禁用。
- Device Clock*：el-date-picker（datetime），placeholder `--Select Device Clock--`，`page.get_by_role("combobox", name="Device Clock*")`。
- Sync 按钮：`get_by_role("button", name="Sync")`（点击即触发同步）。
- NTP Server 1/2/3：`page.get_by_placeholder("NTP Server 1")` 等；提示 "Maximum 40 characters"，**长度校验仅 Save 后触发**（无原生 maxlength）。
- Time Zone*：默认 `Asia/Shanghai(CST)`，`page.get_by_role("combobox", name="Time Zone*")`，支持搜索。
- Save：`get_by_role("button", name="Save")`（各子页同）。

### Network（`network`）
- RSTP Enable：el-switch，默认 OFF。
- Default Interface (Outbound Traffic)：el-select，默认 `Ethernet 1`，`get_by_role("combobox", name="Default Interface (Outbound Traffic)")`。
- Ethernet 1/2 DHCP Enable：el-radio Auto/Manual，默认 Auto。**Auto 时 Interface Status/IP 为只读 textbox（disabled）；切 Manual 时替换为 IP*/Mask*/Gateway* 三个必填输入框**（Eth1 回显 192.168.8.101/255.255.255.0/192.168.8.1），切回 Auto 恢复。
- Interface Status：Eth1=`Connected`、Eth2=`Disconnected`（同名两个，需 scope 到 Ethernet 区块）。
- DNS 1*/DNS 2：`get_by_placeholder("Enter DNS 1")`/`("Enter DNS 2")`，默认 8.8.8.8 / 8.8.4.4，**blur 即校验**。

### Access Control / Whitelist（`whitelist`）
- 页面/导航文案为 **Access Control**，开关文案为 **IP Allow List Enable***（el-radio Enable/Disable，设备常态 Disable）。
- Enable 后出现：`button "Add Allow List"` + el-table（列：No / Description / From IP / To IP / Action；空态 `No Data`）。
- 弹窗 **New Allow List**：`page.get_by_role("dialog", name="New Allow List")`
  - IP Range*：el-radio Yes/No，默认 Yes；Yes→From Address*/To Address*，No→单一 IP Address*（互斥展示）
  - Description：`Enter Description`，非必填
  - Cancel / Confirm 按钮；字段选择器需 scope 到 dialog。
- 注意：当前页面**无 Port Range / Protocol 字段**（旧用例中步骤含 Protocol/Port Range 的均已过时）。

### Email（`email`）
- Email Server*（"Must be valid ip or domain"）/ Email Port*（"Range: 1 - 65535"）/ TLS/SSL*（el-radio Auto/On/Off）/ Sender Name（40）/ From Email Address*（100）/ Username*（40）/ Password（40，带显隐眼睛图标）。
- 选择器统一 `get_by_placeholder("Enter Email Server")` 等。
- **本页没有 Test Email 按钮**（Test Emails 在 Alarm Notification 页）。

### Alarm Notification（`alarm`）
- Alarm Beep*：el-switch，默认 ON，`get_by_role("switch", name="Alarm Beep*")`。
- Alarm Acknowledgement Enable*：el-radio，默认 Enable。
- Alarm Email Enable*：el-radio，默认 Enable。
- Recipient 1*/2/3：`get_by_placeholder("Enter Recipient 1")` 等，Maximum 100 characters。
- Email Interval*：`get_by_placeholder("Enter Email Interval")`，Range 1-10，单位 mins。
- 按钮：Save、**Test Emails**（会真实发邮件，慎点）。

### Certificate Management（`certificateManagement`）
- 只读信息 group（`get_by_role("group", name="Common Name")` 等，Issuer/Subject 同名需 `.first`/`.nth()`）：
  - Issuer/Subject：Common Name / Company Name / Division Name / City / State / Country Code（自签名证书 Subject=Issuer）
  - Validity：Valid From / Expiration
  - Details：Public Key Size(2048) / Serial Number / Public Key Type(RSA) / Certificate Version(3) / Signature Algorithm(sha256WithRSAEncryption) / Extensions
- 按钮：Import / Generate New Self-Signed Certificate / Generate CSR / Export（Generate 类会导致 web 服务重载）。

### Configuration Management（`configurationManagement`）
- 三个 group：`Import Configuration`（Browse + Import）、`Export Configuration`（Export）、`Factory Reset`（Reset，高危）。
- 页面提示文案：`Caution: Importing configuration files between different software versions is not supported.`
- **配置导入和导出都会触发设备重启**（导入重启在用例表有记录；导出重启为用户 2026-07-03 口头确认，用例表未写）。自动化/探查中点击 Export 与 Import 一律视为重启类操作：执行前须经用户确认，脚本需带重启等待+重登逻辑，禁止进无人值守连跑。

### Remote Access（`remoteAccess`）
- Remote Access Enable*：el-radio Enable/Disable，设备常态 Disable。
- **未保存态切 Enable 不展开任何新字段/按钮**（仅本地状态，无 POST）；Manual Register / Refresh Status / Deregister 推测在保存并注册成功后才出现——**未探明**，需允许保存的窗口期补充探查。
