# 接线方式参考：实测相/CT 速查

## 文档来源

原件：`knowledge/gateway/AcuRev4100WEB2/requirements/raw/接线检测总表_ver1.05.xlsx`（10 个 Sheet）
算法全文摘要：`knowledge/gateway/AcuRev4100WEB2/requirements/summaries/wiring_check_v1.05.md`

> 原件和算法全文由 AcuRev-4100-WEB2 项目维护，本文件仅做跨项目速查入口，不重复算法内容。

---

## 各接线方式实测相/CT 一览

| 接线方式代码 | 标准名称 | 电压接入 | 电流接入 | 适用设备说明 |
|------------|--------|---------|---------|------------|
| 3e4wy | 3E4WY（3 Element 4 Wire Y） | Va、Vb、Vc（3相，额定=相电压） | Ia、Ib、Ic（3CT） | 全相接入；AcuHMI、ACM-41-WEB2 均支持 |
| 2.5e4wy | 2.5E4WY（2.5 Element 4 Wire Y） | Va、Vc（2相；无 Vb） | Ia、Ib、Ic（3CT） | 电流检查逻辑同 3E4WY；无多回路表对应；仅 HMI1-7 支持 |
| 2e3wn | 2E3W Network | Va、Vc（2相） | Ia、Ic（2CT；无 B 相） | 多回路 4100/2100 电压检查同 3E4WY；不判相序 |
| 2e3wd | 2E3W Delta（2LL Delta） | Va、Vb、Vc（3相，额定=线电压；N接B） | Ia、Ic（2CT；1320/4100 仅测 A/C） | 电压全相；电流仅 A/C |
| 3e3wd | 3E3W Delta（3LL Delta） | Va、Vb、Vc（3相，N浮空） | Ia、Ib、Ic（3CT） | 电压/电流检查逻辑同 2E3W Delta |
| 3e4wd | 3E4W Delta（HighLeg Delta） | Va、Vb、Vc（3相，额定=线电压；B为高腿） | Ia、Ib、Ic（3CT） | 电流检查逻辑同 2E3W Delta；Van_rated=0.5×VRATE，Vbn_rated≈0.866×VRATE；仅 HMI1-7 支持 |
| 2e3w1p | 2E3W 1Phase | Va、Vc（2相） | Ia、Ic（2CT） | 1310/2100/IIV3 用 A+B；**AcuVim3/4100/1320/AcuRev-100(RACG) 用 A+C** |
| 1e2w1p | 1E2W（1LN） | Va（1相） | Ia（1CT） | 单相单回路；无相序检测 |

> **1320 需量测试特别注意**：2E3W 1Phase（2e3w1p）下 1320 使用 A+C 通道（非 A+B），与 1310/2100 不同；2E3W Delta（2e3wd）下 1320 同 4100 仅测 Ia/Ic。

---

## 各接线方式检测故障类型

| 接线方式 | 支持检测的故障类型 |
|---------|----------------|
| 3e4wy | 电压缺失（单/双/三相）、电压反接、Vb/Vc 相位偏移、相序错误（不对称度≥145%）、电流缺失、电流反接、电流相位偏移 |
| 2.5e4wy | Va/Vc 缺失（单/双）、Va/Vc 反接；电流同 3E4WY |
| 2e3wn | Va/Vc 缺失（单/双）、Va/Vc 反接；电流仅 Ia/Ic 缺失与反接 |
| 2e3wd | 电压缺失（全/单/双相线电压）、相序错误；电流仅 Ia/Ic 缺失与反接/相移 |
| 3e3wd | 同 2E3W Delta |
| 3e4wd | 电压缺失、反接、A-B/B-C/A-C 互换；电流同 2E3W Delta |
| 2e3w1p | Va/Vc 缺失（单/双）、Va/Vc 反接；电流 Ia/Ic 缺失与反接/相移 |
| 1e2w1p | Va 缺失、Ia 缺失、Ia 反接 |

---

## 产品支持矩阵

| 产品 | 支持接线方式 |
|------|-----------|
| ACM-41-WEB2（4100 多回路） | 3E4WY、2E3W Delta、2E3W Network、2E3W 1Phase、1E2W（5 种） |
| AcuHMI-1-7 | 以上 5 种 + 2.5E4WY、3E3W Delta、3E4W Delta（HighLeg）（共 8 种） |
| AcuRev-1320 | 1E2W、2E3W 1Phase、2E3W Network、2E3W Delta、3E4W Y、3E4W Delta（6 种） |
| AcuRev-100（RACG，单回路子表） | 1E2W、2E3W 1Phase（A+C）、3E4WY（3 种，仅单回路）；**判据模型与本表不同**——改用绝对阈值 VLN<10V（非 ×VRATE 比例），V-N 反接比值阈值按接线方式不同（2E3W1P=1.4、3E4WY=1.3，已裁定 2026-07-14），完整判据见 `knowledge/meters/AcuRev100/requirements/summaries/RACG_接线检查表.md` |

---

## 查找算法细节

完整判决规则（条件编号、阈值、优先级、互斥逻辑）见：
`knowledge/gateway/AcuRev4100WEB2/requirements/summaries/wiring_check_v1.05.md`
