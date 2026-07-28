# RPP · Data Log 模块 UI 选择器沉淀

> 对应测试目录：`projects/RPP/tests/data_log/`（迁移自 AcuHMI-1-7 同名目录）
> 通用信息（Web 地址、登录、顶级导航组、协议路由总表）见 [../../context.md](../../context.md#ui-选择器沉淀) 「UI 选择器沉淀」节，本文件只记差异点。
> 数据来源：RPP 纯前端联调环境 http://192.168.2.94:3030 实测。

## 导航路径

顶级组 **Monitoring** → **Data Log**（URL 驱动导航，不依赖 hover 展开）。

## 路由差异（已验证，与 HMI 逐条对比）

| 子页 | RPP 路由 | 备注 |
|---|---|---|
| Data Logger N | `/dataLog/dataLogger/dataLoggerN` | 与 HMI 相同 |
| Post Channel N | `/dataLog/dataForwarding/postChannels/postChannelN` | **RPP 多一段 `dataForwarding`，HMI 无此段** |
| Rapid Logger | `/dataLog/dataLogger/rapidLogger` | 与 HMI 相同 |
| Data Log Parameter Config | `/dataLog/dataLogger/dataLogParameterConfig` | 与 HMI 相同 |
| Post Historical Data | `/dataLog/dataForwarding/postHistoricalData` | 与 HMI 相比同样多 `dataForwarding` 段 |
| AcuCloud | `/dataLog/dataForwarding/acucloud` | 同上 |

**迁移用例改造要点**：所有硬编码 `/dataLog/postChannels/...`、`/dataLog/postHistoricalData`、`/dataLog/acucloud` 的直达 URL 均需在路径中插入 `dataForwarding` 段，否则跳转 404。

## 字段名（与 Post Channel 页已验证一致，与 HMI 相同，无需改动）

- 占位符：`Enter FTP URL` / `Enter FTP Port` / `Enter FTP User Name` / `Enter FTP Password`
- label：`Post Method`
- 按钮：`Test Post Channel` / `Clear Post Channel Logs` / `Save`
- 复选框：`Enable anonymous mode`
- Enable/Disable 单选（el-radio）

其余控件（Data Logger 配置、Post Channel 下拉不可选判定等）沿用 HMI datalog 模块既有事实，见
[../../../../gateway/AcuHMI17/requirements/context/_INDEX_context.md](../../../../gateway/AcuHMI17/requirements/context/_INDEX_context.md)
（该模块沉淀已按子页拆分为 `Devices_DataLog_*_context.md`，如 `Devices_DataLog_DataLogger_context.md` / `Devices_DataLog_PostChannel_context.md`），
仅路由段差异按上表调整。
