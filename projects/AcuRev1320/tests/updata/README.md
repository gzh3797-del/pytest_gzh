# AcuRev-1320 固件升级自动化用例（tests/updata）

依据 `knowledge/meters/AcuRev1320/testcase/AcuRev1320_Firmware升级_用例.xlsx` 生成的 pytest 脚本。
源用例共 **40 条**，分 `023_01` / `023_02` 两个子模块；**每条手工用例都一一对应一个 pytest 方法**，
方法名内嵌完整用例编号，便于与手工用例对照。

- 能脚本化的 **16 条**走自动化（`auto` / `semi`）；
- 其余 **23 条**因需物理按键 / STM32 烧录 / 接源 / 铅封 / 第二上位机等，天生只能手动，写成
  `pytest.skip` 并在报告里记录手动步骤（即"已覆盖、待人工执行"，不是没实现）。
- 注：原 `023_01 case6`（TCP 正常升级）与 `case7` 自动化动作完全相同，已合并入 `case7`，故方法实际为 **39 个**。

底层升级动作复用 `projects/AcuRev1320/QT_Auto` 的 GUI 驱动（pywinauto 取窗口矩形 + pyautogui
图像/坐标点击），升级结果以屏幕上的 `Write_Success` / `Connect_Failed` / `Write_Failed` 判定。

> **⚠️ 实测状态（2026-06-30）：仅 TCP 系列已跑通；RTU 系列因 Acuview 波特率下拉自动点不开，当前未跑通、暂不可用。**
> 详见下方「用例清单」的实测状态说明与各表「实测」列。

---

## 为什么分成 023_01 / 023_02 两个文件

两个文件对应手工用例 xlsx 的两个子模块，按测试主题拆分，便于与手工用例编号一一追溯：

- **`test_firmware_023_01.py`**（类 `TestFirmware02301`）—— 子模块 023_01：**铅封未封闭**下的正常升级路径
  （RTU/TCP 各波特率与网络拓扑、固件签名校验、压力/稳定性、升级期间持续 ping）。
- **`test_firmware_023_02.py`**（类 `TestFirmware02302`）—— 子模块 023_02：**铅封封闭升级失败**、STM32
  烧 boot、通道并发、升级前后寄存器/setting 一致性。

两个文件共用 `firmware_base.py`（公共基类）与 `firmware_actions.py`（升级动作库），升级流程只维护一处。

---

## 目录结构

```
tests/updata/
├── test_firmware_023_01.py   # 子模块 023_01（28 个方法：13 可自动 + 15 手动）
├── test_firmware_023_02.py   # 子模块 023_02（11 个方法：3 可自动 + 8 手动）
├── firmware_base.py          # 用例公共基类（helper/layout 构造 + 安全 teardown）
├── firmware_actions.py       # TCP/RTU 升级动作库（复用 QT_Auto 流程）+ 非法文件校验
├── modbus_helpers.py         # Modbus 寄存器快照/比对（升级前后一致性用例）
├── conftest.py               # 截图 + Excel/JSON 报告 + 防睡眠
├── pytest.ini                # 独立 rootdir + marker 注册
└── README.md
```

---

## 自动化分级（marker）

| 分级 | marker | 含义 | 跑法 |
|------|--------|------|------|
| 全自动 | `auto` | GUI 驱动 + Modbus 回读可直接跑，自动判定升级成功 | 直接执行 |
| 半自动 | `semi` | 升级动作自动化，网络拓扑/并发为**物理前置**（脚本不校验拓扑） | 接好线后执行 |
| 手动 | `manual` | 需物理按键/STM32/接源/铅封/HMI 等，**自动 skip 并在报告里附手动步骤** | 仅作记录，人工执行 |

通道 marker：`tcp` / `rtu`；稳定性：`stress`。

---

## 用例清单（编号 / 分级 / 执行内容 / 节点 ID）

> 单条运行（推荐）：在**仓库根目录**用 `-k 编号` 过滤，不用管它在哪个文件/类——
> `pytest projects/AcuRev1320/tests/updata/ -k case10 -v -s`（只换编号即可）。
> Windows 控制台先设 `$env:PYTHONUTF8=1`（本会话一次）。
> 注：`-k` 是子串匹配，`-k case7` 会同时命中 `case7/case7_1/case7_2`，要精确就用独有后缀（如 `case7_2`、`case9_1`）。

> ### ⚠️ 当前实测状态（2026-06-30）
> - **TCP 系列：✅ 已实测可跑通**（连接/升级/持续 ping 等，节点 ID 含 `tcp` 或走 `_do_tcp`）。
> - **寄存器一致性（023_02 case7_1/7_2）：⏭️ 未配 `CONFIG_REGISTER_BLOCKS`，当前自动 skip，尚未实测。**
> - **RTU 系列：⚠️ 尚未跑通**——升级动作已实现，但卡在 **Acuview 波特率下拉自动点不开**（`_select_baud`
>   检测不到下拉展开态），导致波特率选不上、后续连接失败。模板编码问题（cv2 读不了中文路径）已修复，
>   但「点开下拉」这一步在真机上仍未稳定，台架定位中。**RTU 用例暂不可用，勿计入回归通过率。**
> - 下表「实测」列：`✅` = 已实测可跑；`⚠️` = 未跑通（RTU 待修）；`⏭️` = 依赖项未配置、当前自动 skip，**尚未实测**。

### 一、可自动执行（auto / semi，共 16 条）

**023_01** —— 文件 `test_firmware_023_01.py`，类 `TestFirmware02301`

| 编号 | 分级 | 实测 | 执行内容 | 节点 ID |
|------|------|------|----------|---------|
| case1 | auto·rtu | ⚠️ | RTU 9600 升级，刷到目标版本，判 `Write_Success` | `test_Function_AcuRev1320_023_01_case1` |
| case2 | auto·rtu | ⚠️ | RTU 19200 升级 | `test_Function_AcuRev1320_023_01_case2` |
| case3 | auto·rtu | ⚠️ | RTU 38400 升级 | `test_Function_AcuRev1320_023_01_case3` |
| case4 | auto·rtu | ⚠️ | RTU 57600 升级 | `test_Function_AcuRev1320_023_01_case4` |
| case5 | auto·rtu | ⚠️ | RTU 115200 升级 | `test_Function_AcuRev1320_023_01_case5` |
| case7 | semi·tcp | ✅ | **TCP 正常升级基准**（电表直连电脑、同网段） | `test_Function_AcuRev1320_023_01_case7` |
| case8 | semi·tcp | ✅ | TCP 升级（电表与电脑经路由器） | `test_Function_AcuRev1320_023_01_case8` |
| case9 | semi·tcp | ✅ | TCP 升级（电脑经 WIFI 接路由器） | `test_Function_AcuRev1320_023_01_case9` |
| case9_1 | semi·tcp | ✅ | TCP 升级（电表与电脑接入公网） | `test_Function_AcuRev1320_023_01_case9_1` |
| case10 | semi | ✅ | 加载非法 .MFEA，期望被上位机拒绝（不出现可升级态） | `test_Function_AcuRev1320_023_01_case10` |
| case12 | auto·rtu·stress | ⚠️ | RTU 连续升级 `STRESS_ROUNDS` 次，每次成功 | `test_Stable_AcuRev1320_023_01_case12` |
| case13 | auto·tcp·stress | ✅ | TCP 连续升级 `STRESS_ROUNDS` 次，每次成功 | `test_Stable_AcuRev1320_023_01_case13` |
| case14 | auto·tcp | ✅ | TCP 升级期间后台持续 ping 电表 IP，升级正常完成 | `test_Function_AcuRev1320_023_01_case14` |

**023_02** —— 文件 `test_firmware_023_02.py`，类 `TestFirmware02302`

| 编号 | 分级 | 实测 | 执行内容 | 节点 ID |
|------|------|------|----------|---------|
| case5 | semi·rtu | ⚠️ | 后台起 TCP Modbus 轮询线程模拟并发通信，同时做 RTU 升级（无 pymodbus/连不上则降级为纯 RTU 升级） | `test_Function_AcuRev1320_023_02_case5` |
| case7_1 | auto·tcp | ⏭️ | 升级前后各做一次 Modbus 寄存器快照并比对，验证寄存器不变（**未配 `CONFIG_REGISTER_BLOCKS`，当前自动 skip，未实测**） | `test_Function_AcuRev1320_023_02_case7_1` |
| case7_2 | auto·tcp | ⏭️ | 同 case7_1，比对配置/setting 类寄存器块不变（**同前置，未配置则自动 skip，未实测**） | `test_Function_AcuRev1320_023_02_case7_2` |

### 二、仅手动（manual，自动 skip 并在报告记录手动步骤，共 23 条）

**023_01** —— 类 `TestFirmware02301`

| 编号 | 分级 | 手动步骤摘要 | 节点 ID |
|------|------|--------------|---------|
| case4_01 | manual·rtu | 升级到 ~40% 时关上位机，重开点升级（需物理按键恢复）；预期关闭瞬间提示失败、重连后成功 | `test_Function_AcuRev1320_023_01_case4_01` |
| case9_2 | manual·tcp | 公网 + DHCP 取到 IP 后关闭电表 DHCP，再 TCP 升级 | `test_Function_AcuRev1320_023_01_case9_2` |
| case11 | manual·rtu | RTU 按键强制升级（强制 9600）：上电瞬间物理按 OK 键进 boot | `test_Function_AcuRev1320_023_01_case11` |
| case11_1 | manual·rtu | RTU 按键强制升级（强制 19200） | `test_Function_AcuRev1320_023_01_case11_1` |
| case11_2 | manual·rtu | RTU 按键强制升级（强制 38400） | `test_Function_AcuRev1320_023_01_case11_2` |
| case11_3 | manual·rtu | RTU 按键强制升级（强制 57600） | `test_Function_AcuRev1320_023_01_case11_3` |
| case11_4 | manual·rtu | RTU 按键强制升级（强制 115200） | `test_Function_AcuRev1320_023_01_case11_4` |
| case11_5 | manual | 全新单板 STM32 烧 boot + 升 app 后，basic setting 无非法值（如 slaveID 默认 1） | `test_Function_AcuRev1320_023_01_case11_5` |
| case11_01 | manual·tcp | TCP 按键强制升级、直连：物理按键进 boot + Scan Mode 手动 Add IP | `test_Function_AcuRev1320_023_01_case11_01` |
| case11_02 | manual·tcp | TCP 按键强制升级、公网 DHCP：物理按键进 boot | `test_Function_AcuRev1320_023_01_case11_02` |
| case15 | manual·tcp | 升级界面 IP/Model/Hardware/Firmware 显示正确（肉眼/OCR 核对） | `test_Function_AcuRev1320_023_01_case15` |
| case16 | manual | 升级后接交流源，查电压/电流/角度精度 | `test_Function_AcuRev1320_023_01_case16` |
| case17 | manual | 先升 boot 再升 app，接源查电压/电流精度 | `test_Function_AcuRev1320_023_01_case17` |
| case18 | manual·tcp | 改 slave id=2、TCP port=502 后 TCP 升级（需重配设备 + 建匹配会话） | `test_Function_AcuRev1320_023_01_case18` |
| case19 | manual·tcp | 改 TCP port=508（slave id=1）后 TCP 升级 | `test_Function_AcuRev1320_023_01_case19` |

**023_02** —— 类 `TestFirmware02302`

| 编号 | 分级 | 手动步骤摘要 | 节点 ID |
|------|------|--------------|---------|
| case1 | manual·rtu | 铅封封闭、RTU 升级失败，提示设备处于铅封状态（需物理封铅封） | `test_Function_AcuRev1320_023_02_case1` |
| case2 | manual·tcp | 铅封封闭、TCP 升级失败，提示设备处于铅封状态 | `test_Function_AcuRev1320_023_02_case2` |
| case2_01 | manual | 铅封 sealed 无法升级，手动改 unsealed 后升级界面可继续升级 | `test_Function_AcuRev1320_023_02_case2_01` |
| case3 | manual | STM32 工具反复 Connect/DisConnect，单板稳定不异常重启 | `test_Function_AcuRev1320_023_02_case3` |
| case4 | manual | STM32 烧写 boot（CM7.hex + CM4.hex），重启后 boot Version 为最新 | `test_Function_AcuRev1320_023_02_case4` |
| case4_01 | manual | boot 模式下 HMI 界面信息显示正确（Model=AcuRev1320 及版本号） | `test_Function_AcuRev1320_023_02_case4_01` |
| case6 | manual·tcp | RTU 通信进行时（sscon 循环发帧），TCP 升级成功 | `test_Function_AcuRev1320_023_02_case6` |
| case7 | manual·tcp | TCP 升级中第二个上位机抢连同一电表（连接失败，原升级不受影响） | `test_Function_AcuRev1320_023_02_case7` |

---

## 运行前准备（必做）

1. **装依赖**：`pyautogui`、`pywinauto`、`psutil`、`pyperclip`、`pandas`、`openpyxl`、
   `pytest`、`pytest-html`、`pymodbus`（仅一致性用例需要）。

2. **补连接配置**（`configs/global.yaml`，QT_Auto 同款）：
   ```yaml
   QT_path: 'C:\Users\<你>\Acuview2\Acuview 2.exe'
   device_image_path: 'page_elements\AucRev1320\1320 TCP'
   QT_tcp: { host: '192.168.x.x', port: 502, timeout: 10, slave_id: 1 }
   QT_rtu: { com_port: 11 }   # RTU 升级时设备所在 COM 口（Meter Update 列表里的 "Com 11"）
   ```

3. **Acuview 里建好 TCP 会话**（命名 `1320 TCP`，手动能连上电表），截图存
   `comm/QT_comm/page_elements/AucRev1320/1320 TCP.png`。

4. **升级包路径**：默认指向 `QT_Auto/data/` 下的 `.MFEA`（`firmware_actions.PACKAGE_TARGET` /
   `PACKAGE_BASE`）。备齐目标版本 + 基线版本后分别指向，可做往返刷写。

5. **显示环境固定为基线**：分辨率 **1920×1080**、缩放 **125%**、Acuview 主窗口**最大化**
   （流程已在升级前自动最大化主窗口）；换分辨率/缩放需按 `QT_Auto/firmware_layout.py` 重标坐标。

6. **RTU 用例前置（坐标已标定，换台架需重标）**：`QT_Auto/data/firmware_layout.json` 中
   `AddConn_Close`（关 Add Connection）与 `com_checkboxes`（按 COM 号的 Select 复选框，当前已标
   Com 1 / Com 11）须有标定坐标；`QT_rtu.com_port` 指定的 COM 口必须在 `com_checkboxes` 里有对应项，
   否则 RTU 升级会 fail-fast。

   **RTU「态」模板（条件等待，需真机抓图）**：RTU 流程改为「等控件可点再点」，依赖两张状态模板存于
   `comm/QT_comm/page_elements/Acuview_public/main_page/`：
   - `BaudRate_DropDown.png` —— 波特率下拉**展开后**的可识别区域（展开列表边框/选项行），
     用于判定下拉已弹出再点选项（解决「波特率没选上」）；
   - `Connect_Disabled.png` —— **未勾选 COM 口时** Connect 的灰色不可点态，以其「消失」判定 COM 已选中、
     Connect 已可点（解决「点在灰态 COM/Connect 上无效 → OK 对话框不出现而超时」）。

   ⚠️ **模板文件名必须用 ASCII**：`cv2.imread`（pyautogui 的 confidence 图像匹配底层）在 Windows
   读不了含中文的路径，中文名模板会**静默匹配失败**、等待退化为盲点。`Select_All_Disabled.png`（TCP 用）
   亦同此规则。抓图方式：在真机对应界面截取该控件局部，存为上述 ASCII 文件名。
   **未标定时不阻断**：对应等待会打日志后按「已就绪」放行，退化为原盲点行为，故强烈建议在台架补齐。

7. **一致性用例（023_02 case7_1/7_2）前置**：在 `modbus_helpers.CONFIG_REGISTER_BLOCKS`
   按 AcuRev-1320 Modbus 地址表填入 basic setting / 配置类寄存器块 `(起始地址, 寄存器数)`，
   未配置时这两条自动 skip。

8. **运行期间别动鼠标键盘、别锁屏**（坐标点击被打断会失败；已内置防睡眠）。

---

## 怎么跑

在**仓库根目录**（`autotest/`）执行：

```bash
# 全部用例（auto 跑、semi 接好线跑、manual 自动 skip）
pytest projects/AcuRev1320/tests/updata/ -v -s

# 只跑全自动用例（首次冒烟建议）
pytest projects/AcuRev1320/tests/updata/ -m auto -v -s

# 只跑 TCP / 排除压测
pytest projects/AcuRev1320/tests/updata/ -m "tcp and not stress" -v -s

# 只跑某一条（推荐：-k 编号过滤，不用写文件/类全路径，只换编号即可）
pytest projects/AcuRev1320/tests/updata/ -k case10 -v -s

# （等价的全路径写法，编号跨文件/类时容易写错，一般不用）
# pytest "projects/AcuRev1320/tests/updata/test_firmware_023_02.py::TestFirmware02302::test_Function_AcuRev1320_023_02_case7_2" -v -s
```

> Windows 控制台若报 GBK 编码错，加 `PYTHONUTF8=1`（conftest 已尝试把 stdout 切到 UTF-8）。

---

## 结果与报告

每次运行在仓库根 `reports/updata_<时间戳>/` 下生成：HTML / Excel（测试详情含跳过原因 + 统计）/
JSON / 分层截图（文件名带 PASSED/FAILED）。

---

## 与 QT_Auto 的关系

`firmware_actions.py` 直接复用 `QT_Auto/firmware_layout.py` 与 `QT_Auto/data/firmware_layout.json`
（坐标只标一处）。升级主流程逻辑与 `QT_Auto/test_firmware_new.py` 一致，差别仅在：本目录按**用例编号**
组织、覆盖完整的源用例并做了自动化分级，QT_Auto 侧重 TCP/RTU 往返回归压测。
