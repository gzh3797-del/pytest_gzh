# AcuRev-1320 上位机固件升级自动化（QT_Auto）

通过模拟操作 **Acuview 2 上位机软件**，自动完成 AcuRev-1320 的固件升级回归：在两个固件版本之间反复刷写（target ↔ base），支持 **TCP / RTU** 两条通道，自动出报告。

底层是「pywinauto 取窗口矩形 + pyautogui 图像/坐标点击」。所有控件坐标已外置到 `data/firmware_layout.json`（相对锚点的相对坐标），换机器只需重标锚点，不改代码。

---

## 1. 目录结构

```
QT_Auto/
├── test_firmware_new.py     # 升级流程主逻辑 + 动态生成用例
├── firmware_layout.py       # 坐标解析驱动（读 JSON / 取窗口矩形 / 语义点击）+ 标定工具
├── conftest.py              # 截图、防睡眠、生成 Excel/JSON/HTML 报告
├── pytest.ini               # pytest 配置
├── data/
│   └── firmware_layout.json # 控件坐标表（相对锚点偏移）
└── README.md                # 本文件
```

---

## 2. 运行前准备（必做）

> ⚠️ 这几项不配好，跑起来会直接报错。

1. **装依赖**：`pyautogui`、`pywinauto`、`psutil`、`pyperclip`、`pandas`、`openpyxl`、`pytest`、`pytest-html`。

2. **补连接配置**（当前 `configs/global.yaml` 里缺这几个键，脚本启动会 `KeyError`）：
   ```yaml
   QT_path: 'C:\Users\<你>\Acuview2\Acuview 2.exe'   # Acuview 2 可执行文件绝对路径
   device_image_path: 'page_elements\AucRev1320\1320 TCP'  # 设备连接的截图元素（不带 .png）
   QT_tcp:
     host: '192.168.x.x'
     port: 502
     timeout: 10
     slave_id: 1
   ```

3. **Acuview 里建好 TCP 连接会话**：命名为 `1320 TCP`，确保**手动能连上电表**；把该连接的截图存到
   `page_elements\AucRev1320\1320 TCP.png`。

4. **准备固件包**，并修改 `test_firmware_new.py` 顶部两行为 **1320 的实际固件包路径**：
   ```python
   package_path_target = r'C:\autotest_local\update_version\<1320 新版本>.MFEA'
   package_path_base   = r'C:\autotest_local\update_version\<1320 基线版本>.MFEA'
   ```
   > 当前仓库里这两行还是 AcuRev-4100 的包，**必须换成 1320 的**。

5. **显示环境固定为基线**：分辨率 **1920×1080**、缩放 **125%**、运行时让 Acuview 主窗口**最大化**。
   换分辨率/缩放需重新标定坐标（见第 5 节）。

6. **运行期间别动鼠标键盘、别锁屏**（脚本靠坐标点击，被打断会失败；已内置防睡眠 + 定时点击防锁屏）。

---

## 3. 怎么跑

在**仓库根目录**（`autotest/`）执行。

```bash
# 跑全部已启用用例（当前默认：TCP 2 轮）
pytest projects/AcuRev1320/QT_Auto/test_firmware_new.py -v -s

# 只跑某一条（先冒烟验证建议这样）
pytest projects/AcuRev1320/QT_Auto/test_firmware_new.py::TestAcuviewAutomation::test_TCP_update_round0 -v -s

# 遇到第一个失败就停（适合首次调试）
pytest projects/AcuRev1320/QT_Auto/test_firmware_new.py -v -s -x
```

也可以直接 `python projects/AcuRev1320/QT_Auto/test_firmware_new.py`（文件末尾自带 `pytest.main`，带 `-x` + HTML 报告）。

### 调整轮数 / 启用 RTU

文件末尾：
```python
TestAcuviewAutomation.generate_tests_tcp(rounds=2)      # TCP 跑几轮
# TestAcuviewAutomation.generate_tests_rtu(rounds=1)    # 取消注释启用 RTU（波特率全档 × rounds 轮）
```
- TCP 生成用例：`test_TCP_update_round0`、`test_TCP_update_round1` …
- RTU 生成用例：`test_RTU_update_115200_round1`、`test_RTU_update_57600_round1` …（115200/57600/38400/19200/9600 五档）
- **启用 RTU 前**必须先标定 `AddConn_Close`（见第 5 节），否则该用例会 fail-fast 报「坐标未标定」。

---

## 4. 结果与报告

每次运行在 `reports/<时间戳>/` 下生成：

| 产物 | 说明 |
|------|------|
| `test_report_*.html` | HTML 报告（self-contained） |
| `test_report_*.xlsx` | Excel 报告（测试详情 / 统计 / 文件统计，含通过率） |
| `test_report_*.json` | JSON 结构化结果 |
| `screenshots/` | 按 用例 分层的截图（文件名带 PASSED/FAILED） |
| `reports.txt` | 报告清单 + 截图目录树 |

**升级判定逻辑**（单个升级 30 分钟超时，每 20s 轮询屏幕）：
- 出现 `Write_Success` 图 → 通过
- 出现 `Connect_Failed` / `Write_Failed` 图 → 失败
- 超时无结果 → 失败

---

## 5. 坐标标定 / 排错（换机器或点歪了看这里）

坐标都在 `data/firmware_layout.json`，`x/y` = **相对锚点左上角的偏移（物理像素）**。锚点：
- `main` = Acuview 2 主窗口
- `add_conn` = `Add Connection` 对话框

### 5.1 跑前先肉眼核对坐标（强烈建议）

把 Acuview 摆到升级界面，运行：
```bash
python projects/AcuRev1320/QT_Auto/firmware_layout.py --overlay check.png
```
打开 `check.png`，红圈就是脚本会点的位置。圈歪了就去改 JSON 对应控件的 `x/y`。

### 5.2 标定一个新坐标（如 `AddConn_Close`）

打开对应窗口/对话框后运行（无参数 = 实时坐标模式）：
```bash
python projects/AcuRev1320/QT_Auto/firmware_layout.py
```
它会先打印各锚点矩形，然后**实时显示鼠标相对各锚点的偏移**：
```
[锚点] main       rect=(0,0)-(1920,1080)  尺寸 1920x1080
[锚点] add_conn   rect=(510,180)-(1410,720)  尺寸 900x540
鼠标绝对(1398,205)  相对偏移 main:(1398,205)  add_conn:(888,25)
```
把鼠标移到目标控件（如 Add Connection 的关闭按钮）上，读它**相对所属锚点**的偏移，填回 JSON：
```jsonc
"AddConn_Close": { "anchor": "add_conn", "x": 888, "y": 25, "note": "..." }
```
`Ctrl+C` 退出。

### 5.3 常见问题

| 现象 | 排查 |
|------|------|
| `KeyError: 'QT_path'` | 第 2 节连接配置没补 |
| 启动报「坐标未标定(x/y 为空)」 | 该控件 JSON 里 `x/y` 是 `null`，按 5.2 标定（目前只有 `AddConn_Close` 待标） |
| 点击位置整体偏移 | 分辨率/缩放不是基线，或窗口没最大化；用 5.1 核对、必要时重标 |
| 找不到 `Add Connection` 窗口 | 该对话框没打开就调用了 `add_conn` 锚点 |
| 升级中途失败 | 看 `reports/<时间戳>/screenshots/` 的 FAILED 截图定位卡在哪一步 |

---

## 6. 设计说明（维护者）

- **图像模板点击**（`Operation`/`firmware`/`Select_Firmware_File`/`Connect`/`Yes`/`OK` 及结果判定图）走 `AutoHelper.click_image`，自定位、可移植，**不依赖坐标**。
- **坐标点击**仅用于无法模板化的控件（勾选框、下拉、列表行、防锁屏点），全部经 `FirmwareLayout` 解析相对坐标 → pywinauto 取锚点矩形 → 绝对坐标。
- DPI 感知由 `comm/ctl_acuview/dpi` 提供（`firmware_layout.py` / `test_firmware_new.py` 顶部最先导入），保证 125% 缩放下截图/点击/窗口矩形坐标一致。
- 通用点击原语 `AutoHelper.click_rel(rect, rx, ry)` 在 `comm/QT_comm/QT_utils/QT_auto_utils.py`。
