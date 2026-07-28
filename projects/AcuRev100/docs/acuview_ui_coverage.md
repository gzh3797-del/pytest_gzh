# AcuRev-100 上位机界面坐标覆盖对照（用例转自动化可行性）

> 对照来源：
> - 界面规格：`spec/pages.json`（源文件 Default_1.01，8 个页面）、`spec/registers.json`
> - 手工用例：`tests/manual_testcase/AcuRev-100测试用例_已回填自动化覆盖情况.xlsx`（200 条，35 个子模块；2026-07-17 改名）
>
> 生成日期：2026-07-08。pages.json 更新后本文档需同步复核。

## 一、结论速览

| 覆盖档位 | 用例模块（编号） | 约用例数 | 自动化可行性 |
|---|---|---|---|
| ✅ 页面结构完整 | 001~009、015、016(页面部分)、017(页面部分) | ~110 | 坐标定位可直接做；写操作仍被全局弹窗层阻塞（见三-1） |
| 🟡 页面在、仅整块坐标 | 010 接线检查、011 计时功能、013 铅封 | ~52 | 页面能进；断言/操作点在复合组件内部，无坐标 |
| ❌ 界面完全缺失 | 014 Firmware升级、校表方案、全局弹窗层 | ~35 | 无法定位，需补规格 |
| ➖ 不涉及上位机 UI | 012 LCD/LED、J-Flash 类(014 case13~15) | ~8 | 硬件/外部工具，不需要坐标 |

## 二、页面结构完整的模块（可定位）

pages.json 已收录 8 个页面：General / Communication / Current & Wiring / Device Information / Real-Time Metering / Energy / Wiring Check / System Status。

| 用例模块 | 用例数 | 依赖页面 | 已有关键控件坐标 |
|---|---|---|---|
| 001 交流频率 | 4 | Real-Time Metering | System_Table (30,60,1000,450) |
| 002 相电压 / 003 线电压 | 7+4 | Real-Time Metering | 同上 |
| 004 电流测量 | 19 | Real-Time Metering + Current & Wiring | 表格；Channel A/B/C 的 Input Wiring / CT Type / CT Primary / Direction 各控件坐标齐全（y=220/263/306 三行） |
| 005 PF、功率 | 14 | Real-Time Metering | 表格 |
| 006 相角验证 | 4 | Real-Time Metering | Phase_Angle_Polar_Chart (30,520,900,620) |
| 007 Ep/Eq/Es 能量累计 | 13 | Energy | Real_Time_Table (30,60,860,360)；Energy_Clear_Button (750,23,140,25) |
| 008 Ep+脉冲（光/电） | 11 | General | Energy_Pulse_Constant_Value_Edit (255,253)、Pulse_LED_Width_Value_Spin (215,306)、Energy_LED_Constant_Value_Edit (615,310) |
| 009 接线方式验证 | 12 | Current & Wiring | Service_Configuration 下拉 (210,60,200,30) + 通道配置 |
| 015 数据重置/恢复出厂 | 3 | System Status | Network_Reset (150,4)、Factory_Reset (320,4)、Meter_Reboot (490,4)，均 160×28 |
| 016 通讯模块（页面部分） | 15 | Communication | RS485_Baud_Rate_Combo (205,60)、RS485_Parity_Combo (775,60)、USB_Baud_Rate_Combo (205,150)、USB_Parity_Combo (775,150)、Modbus_Slave_ID_Spin (205,240) |
| 017 密码（页面部分） | 5 | General | Password_Value_Edit (155,60,130,30) |

说明：Real-Time Metering / Energy / Device Information 内各 readValueLabel/readNameLabel 坐标为 null（表格行内元素）。若断言走 Modbus 回读（acuview_auto 现行方案）不受影响；若需 UI 层取值/截图断言，需补行级定位方式。

## 三、缺失界面（pages.json 完全没有）

### 1. 全局弹窗层——阻塞所有写操作，最高优先级

| 缺失界面 | 需要的坐标/元素 | 影响范围 |
|---|---|---|
| 连接/断开连接界面 | 连接类型(RTU/USB/TCP)、COM 口、波特率、校验、Slave ID、Connect/Disconnect 按钮、连接状态指示、设备列表（双击连接、列表头） | 几乎全部用例的步骤 1（"链接状态为 connected"×123）；016 通讯模块 15 条直接测该界面 |
| 密码输入弹窗 | 密码输入框、确认/取消、错误提示 | 017 密码 4 条；所有 setting 写操作、factory reset、设置时间均触发 |
| Update（下发）按钮 | 各 Setting 页全局 Update 按钮 | General / Communication / Current & Wiring 所有写用例 |
| 二次确认与结果提示弹窗 | factory reset / reboot / clear energy / clear run(load) time 的确认框、成功/失败提示 | 015 全部、013 铅封 case4、011 计时 case8~11 |

### 2. 缺失的功能界面

| 缺失界面 | 需要的坐标/元素 | 影响用例 |
|---|---|---|
| Firmware Update 界面 | 升级文件选择、升级波特率下拉、开始升级按钮、进度显示（百分比、clearing 状态）、设备信息栏(Model/Hardware/Firmware) | 014 Firmware升级 17 条（case13~15 为 J-Flash 外部工具，除外）+ 013 铅封 case3 |
| 校表（Calibration）界面 | 整个校准界面 | 校表方案 15 行（用例正文尚为空，可暂缓） |
| 导航树节点坐标 | Setting/Reading 树各节点点击坐标（nav 仅有 Sequence 顺序，无坐标） | 所有页面切换；017 case1/5 要求切换 reading 再回 general |

⚠️ 疑点：017 case5 路径写的是 "Setting—Maintenance—Password"，但 Setting 树只有 General / Communication / Current & Wiring / Device Information。**Maintenance 节点疑似缺失，或用例路径过时**，需与上位机实测界面核对后二选一修正。

## 四、页面在、但内部控件无坐标（半覆盖）

| 模块 | 用例数 | 已有 | 缺失 |
|---|---|---|---|
| 010 接线检查 | 35 | Wire_Check 页；Wire_Check_Switch_Button (30,60,60,26) | customwirecheckcomponent 整块 (30,110,1112,620) 内部的三相电压/电流值、相角、Error Code、相序显示均无独立坐标（断言点全在里面；xref 中 Wire_Check 13 个读项全部 unmatched） |
| 011 计时功能 | 12 | System Status 页；Run_Time_Clear (526,432)、Load_Time_Clear (526,480) | Time_Status_Component 整块 (30,60,640,350) 内部的 Use PC Time / Use Customer Time 单选、年月日时分秒输入框、时间下发按钮 |
| 013 铅封 | 5 | 依赖的操作页（General、Current & Wiring、clear energy / factory reset 按钮）齐全 | Seal_Status 读值标签坐标为 null（在 System_Status_Table 行内）；另依赖缺失的升级界面（case3）和结果提示弹窗（case4） |

## 五、补坐标优先级建议

1. **连接界面**——不补则任何用例第一步都走不通。
2. **密码弹窗 + Update 按钮 + 确认/提示弹窗**——不补则所有 setting 写用例、重置类用例只能读不能写。
3. **Firmware Update 界面**——独立解锁 17 条用例。
4. **Time_Status_Component 内部控件**——解锁计时功能 12 条的设置时间操作。
5. **Wire Check 组件内部读数区**——若接线检查断言可走 Modbus 回读（Wire_Check 读序列 12288/12292 已在 read_sequence 中），此项可降级为低优先。
6. 校表界面——待用例正文补全后再定。

## 六、xref 匹配情况（registers.json 侧参考）

pages.json 自带 xref 统计：read_list 共 145 项，与寄存器匹配 97 项（66.9%）。未匹配集中在 Wire_Check（13 项全缺）、Current_Wire（通道配置 12 项）、System_Status（时间/重置 6 项）等，与上文缺口一致。做 Modbus 回读断言前需先在 registers.json 中补齐这些项的地址映射。
