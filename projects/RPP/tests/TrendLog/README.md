# TrendLog（接入设备 Trend Log 用例组）——占位骨架

RPP 项目「接入设备日志管理 / Trend log」子模块 12 条手工用例
（Function_RPP_025_001_case1 ~ case12）的**占位脚本**（全部 `@pytest.mark.skip`）。

## 为什么是占位（2026-07-17 实测结论）

Trend Log 为 **RPP 需求**（编号 Function_RPP_025），当前替身机 AcuHMI-1-7
（192.168.3.71）固件**未实装该页面**：

- 设备详情 Logs → Trend Log 子菜单存在（Realtime Log / Energy Log / Management
  三个子项），但点击不发生路由跳转（页面停留在 Metering）
- 直连路由 `/#/physicalDevices/deviceDetails/<id>/3:2` 可进入但渲染空白
  （无控件、无图表、无报错），`/3:1` 的 Data Log 则正常
- AcuvimIIW / AcuRev4100_392 两台设备表现一致

## RPP 真机就绪后的补实现要点

1. 现场探查三个子页结构（Time Frame / Time Interval / meterpoint / Checkpoint /
   Data review / Image / Data 下载按钮），沉淀到 knowledge 选择器文档
2. 打点周期校验建议走图表数据 API 响应（echarts canvas 不可直接读数）
3. Interval 档位随区间跨度联动的规格：总间隔时长 / TimeInterval < 1000（点数）
4. case11 / case12 **需重启被测设备**，执行前必须与负责人确认时间窗口
5. 参数覆盖：Realtime 折线（VLN/VLL/I/P/S/PF/FREQ）、Energy 柱状
   （EP_IMP/EP_EXP/EP_Net/EP_total/EQ_*/ES）
