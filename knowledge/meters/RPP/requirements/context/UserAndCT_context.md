# User and CT 页面选择器沉淀

**Web 导航路径**：Physical Devices → `<设备>`（如 Acurev1234100）→ 顶部 Settings 标签 → User and CT
**页面 URL（hash 路由）**：`.../#/physicalDevices/deviceDetails/MODTCP<hash>/4:4?deviceModel=AcuRev-4110-mA`
**对应测试目录**：`projects/RPP/tests/parameter_settings/acurev4100/`（`_src_user_and_ct.py` + `test_Testcase_AcuHMI_001_07_case38~42.py`）
**探查环境**：真机网关 `https://192.168.3.71`（AcuHMI-1-7，型号 AcuRev-4110-mA），账号 q，前端 Element Plus v2

> 注意：此页与 `context.md` 里旧记的"纯前端联调 http://192.168.2.94:3030"**不是同一台机**。192.168.3.71 有真实后端（登录/保存均后端校验）。

---

## 页面结构

三个可折叠区块，底部一个**固定**（`position:fixed;bottom:0`）Save 操作栏。两张表都是自定义
`.custom-table`（原生 `<table>`，无虚拟滚动），非 `.el-table`。

| 区块 | 定位 |
|------|------|
| Wiring Configuration | `.c_card_layout`(filter has_text 'Wiring Configuration') → `.el-select__wrapper` 第 1 个 |
| Current Input Channel（24 行） | `.custom-table` nth(0) |
| User and Channel Mapping（12 行） | `.custom-table` nth(1)；**1E2W 下整块不渲染**（count 从 2 降为 1） |

### Current Input Channel 列（`td` nth）
0=Logic ID(文本) / 1=Input Wiring(el-select) / 2=CT Type(el-select) / 3=Primary(input) / 4=Direction(el-select) / **5=Voltage Assignment(el-select)**

- 某行 VA 下拉：`.custom-table` nth(0) → `tbody tr` nth(row) → `td` nth(5) → `.el-select__wrapper`
- 是否只读：wrapper class 含 `is-disabled`
- 读可选项：click wrapper → `.el-select-dropdown__item:visible` 取文字 → Esc

### User and Channel Mapping 列（`td` nth）
0=ID(文本 "User Channel N") / **1=Description(el-input)** / 2=Phase A / 3=Phase B / 4=Phase C(el-select)

- Description 输入框：`tbody tr` nth(uc) → `td` nth(1) → `input`；**无 maxlength、无 placeholder**
- Phase A/B/C 下拉：`td` nth(2/3/4) → `.el-select__wrapper`；只读判据同上（class `is-disabled`）

---

## 接线方式（Wiring）行为 —— 实测

**切换即自动保存**（无需点 Save），触发约 5s "Running" 重算忙态。
**等待判据**（唯一可靠）：轮询底部 Save 按钮 `.buttonFixed.c_common_button_fix button` 文案，
由 `Running` 变回 `Save` 即结束。绿色 ✅ 图标转瞬即逝不可靠；VA 下拉可交互性在 1E2W/Delta 下永久
disabled，也不能当判据。

| 接线方式 (Service Config 编码 reg 4162) | VA 列规律 | VA 下拉 | VA 可选项 | User Channel 区块 | UC1 Phase A/B/C |
|---|---|---|---|---|---|
| 1 Element 2 Wire (0) | 全 Va | 只读 | — | **不显示** | — |
| 2 Element 3 Wire 1 Phase (1) | 奇 Va / 偶 Vc | 可编辑 | {Va, Vc} | 显示 12 行 | A/C 可编辑，B 固定 disabled |
| 2 Element 3 Wire Delta (2) | 奇 Vab / 偶 Vbc | 只读 | — | 显示 12 行 | 全部固定 disabled |
| 2 Element 3 Wire Network (3) | Va/Vb/Vc 循环 | 可编辑 | {Va, Vb, Vc} | 显示 12 行 | 全部可编辑 |
| 3 Element 4 Wire Y (4) | Va/Vb/Vc 循环 | 可编辑 | {Va, Vb, Vc} | 显示 12 行 | 全部可编辑 |

> Delta 下 Input Wiring 列也是只读；其余模式 Input Wiring/Direction 可编辑。
> Phase A/B/C 下拉候选受"Input Channel 池"限制：当前设备 8 个 User×3 相占满 24 路，
> 候选只剩 `none`。要看到完整 `Input Channel N` 候选须先腾出通道。

## 保存反馈 —— 实测（Description 字段校验）

- 前端**无即时校验**（中文、超长字符输入框照收），**点 Save 后 ~800ms 内**才校验拦截。
- 反馈统一走 Element Plus 顶部 toast：`.el-message.el-message--warning.is-closable`（失败也是
  `--warning`，非 error 变体），约 3s 自动消失。**不走 `.el-form-item__error` 内联**。
- 文案样例：非 ASCII → `"User Channel N name must contain only ASCII characters"`；
  超 20 字符 → `"User Channel N name must be less than 20 characters"`；无变化 → `"No change to save"`。
- 登录页报错是另一套非标准 toast（`user and password don't match`），与保存反馈机制不同。

## 相关 Modbus 寄存器（Basic Setting，供跨传输回读）

| 项 | 地址 | 说明 |
|----|------|------|
| Service Configuration (Wiring) | 4162 (0x1042) | 0:1E2W 1:2E3W1P 2:Delta 3:Network 4:3E4WY |
| 通道 N Voltage Assignment | 4170+(N-1)*5 (0x104A起) | 0:Va 1:Vb 2:Vc 3:Vab 4:Vbc 5:Vca |
| User N 输入通道分配(位图) | 4286+(N-1)*2 (0x10BE起) | uint32，bit0~23=channel1~24 |
| User N Description | 4608+(N-1)*10 (0x1200起) | ASCII 20 字节（v1.03 新增，组 B 用） |
