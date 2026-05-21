# AcuHMI-1-7 Sprint3 软件需求规格变更说明书 — 摘要

**原文件：** AcuHMI-1-7 Sprint3 软件需求规格变更说明书_v1.00_20260430.docx  
**文件编号：** JZZY-A4-RD-50101  
**版本：** v1.00（正式发布）  
**生效日期：** 2026-04-30  
**图片数：** 11 张（见原始 docx）

---

## 概述

在 Sprint3 主需求规格（V1.03）发布后，新增以下 5 项需求变更：

---

## 变更内容

### 3.1 Alarm 功能修改

**Alarm Acknowledgement Enable 处于 Disable 状态时的行为：**

- 触发告警 → 在 Alarm Log 中记录日志
- 告警恢复后 → 自动在 Alarm Log 记录一条日志，Reason 字段显示：
  - `UNDERFLOW_CLEARED`
  - `ONLINE`
  - `OVERFLOW_CLEARED`
- 告警自动恢复后，蜂鸣器停止发声

### 3.2 Post Channel 测试通道显示更详细连接信息

- 参考 AXM-WEB2 的实现方式，展示更详细的连接状态信息

### 3.3 Data Log 参数配置中增加默认参数配置

| 设备类型 | 默认参数配置 |
|---------|------------|
| 物理设备 | Realtime（电压、电流、功率、频率）+ Energy（进线/出线有功能量） |
| 虚拟设备 | 全量参数默认加入 Data Log |

### 3.4 BACnet/IP 协议支持 Harmonic 参数

- 新增对 Harmonic 参数的 BACnet/IP 协议支持
- 支持设备：AcuRev-2100、AcuRev-4100、Acuvim2、Acuvim3、AcuRev-1300

### 3.5 WEB Device 超时功能

- 超时时间：20 秒
- 超时提示语：`"We couldn't reach the webpage. Please double-check the URL or network connection and try again."`
- 加载中显示加载中页面
