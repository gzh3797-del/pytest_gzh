# 接线检查需求变更摘要 v1.00（2026-05-19）

原件：`raw/AcuRev-4100-WEB2 Sprint2 软件需求接线检查需求变更_v1.00_20260519.docx`
算法规格：`raw/接线检查AcuRev-4100 Sprint 3v1.1 20260520.xlsx`（前5个 Sheet）
文档编号：JZZY-A4-RD-50101  变更类型：MODIFIED

---

## 功能概述

为 AcuRev-4100 提供接线检查，辅助用户在安装调试阶段排查接线故障。

**入口**：Settings → Diagnostic → Wiring Check

**运行模式**：
- **WEB Module 模式**：设备不可选，自动检查所有在线设备
- **Gateway 模式**：Device 下拉可选单台或全部在线设备

---

## 触发流程

```
点击绿色 Wiring Check 按钮
    ↓
弹出 Confirm Nominal Voltage 对话框
    列出所有设备：在线设备可修改额定电压并 Save；离线设备置灰不可操作
    ↓
点击 Start Wiring Check → 开始检查（按钮变 Stop）
    Cancel → 取消
    ↓
检查完成，结果展示于两张独立表格
```

> **注意**：测试环境中实测（v1.00 固件）点击 Wiring Check 后未出现 Nominal Voltage 弹窗，直接开始检查。可能为固件版本差异，测试脚本已按无弹窗方式处理。

---

## 结果表格结构

**电压表**
| 列 | 内容 |
|----|------|
| Device | 设备名称 |
| Wiring Configuration | 接线方式 |
| Phase | A / B / C（Delta 为线电压） |
| Measurement | 幅值∠相角 |
| Wiring Status | Pass / Missing / Reversed / Phase Shift / Phase Order Error |

**电流表**
| 列 | 内容 |
|----|------|
| Meter Point | User Channel 名称（1E2W 为 Input Channel） |
| Device | 设备名称 |
| Wiring Configuration | 接线方式 |
| Phase | A / B / C（视接线方式） |
| Input Channel | 物理输入通道编号 |
| Measurement | 幅值∠相角 |
| Wiring Status | Pass / Missing / Polarity Reversed / Phase Shift |

---

## 各接线方式检测范围

| 接线方式 | 电压侧 | 电流侧 | 活跃 User Channel |
|---------|-------|-------|-----------------|
| 3E4WY | A/B/C 相电压 | A/B/C 三相 | User 1–8 |
| 2E3W Delta | Vab/Vbc/Vca 线电压 | A/C 两相 | User 1–12 |
| 2E3W Network | A/B/C 相电压（同3E4WY） | A/B/C 三相 | User 1–12 |
| 2E3W 1Phase | A/C 相电压 | A/C 两相 | User 1–12 |
| 1E2W | A 相电压 | A 相（24 Input Channel 各自独立） | 24 |

---

## 算法规格要点

算法完整定义见 `raw/接线检查AcuRev-4100 Sprint 3v1.1 20260520.xlsx` 前5个 Sheet：
- 优先级：缺失 > 反接 > 相位错误（2E3W Delta 无反接检测）
- 电压侧与电流侧**相互独立**，不互相影响
- 条件 11–13（3E4WY 相位错误）**相互独立**，可同时触发
- 电流反接（PF ≤ −0.9）与相位错误（|PF| ≤ 0.9）互斥
- 4100 多回路表电流相位错误仅显示 "Phase Shift"，不显示具体原因

---

## Modbus 寄存器（接线检查专用）

| 寄存器 | 地址 | 说明 |
|-------|------|------|
| Wire Check Start | 0x1300 | 写1触发，写0停止 |
| Wire Check Status | 0x1301 | 0=未开始，1=进行中，2=已完成 |
| Phase Voltage Error Code | 0x1302 | 9 bits，见 core/config.py |
| User1–12 Current Error Code | 0x1303–0x130E | 各9 bits |
| 测量值（电压幅值/相角/电流幅值/相角） | 0x130F–0x13AC | float × 若干寄存器 |
