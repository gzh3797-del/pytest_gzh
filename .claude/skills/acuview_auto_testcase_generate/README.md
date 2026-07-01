# acuview_auto_testcase_generate — 使用说明书

把电表**手工用例 xlsx** 自动转成 **pytest 自动化用例**：驱动 **Acuview 2 上位机**做界面配置下发 / 读取，再用 **Modbus 跨传输回读**做闭环断言。一套引擎多项目共享。

> 本 README 给工程师看（原理 / 用法 / 排错）。同目录 `SKILL.md` 是给 Claude 的触发指令——在 Claude Code 里说"把这份手工用例转成自动化"即触发本 skill。

---

## 1. 用途

- **输入**：某项目的手工用例 xlsx（含「测试步骤/预期结果/是否可自动化」等列）。
- **输出**：`projects/<项目>/tests/acuview_auto/test_acuview_<用例编号>.py`（一条用例一个文件），以及把结果回填进手工 xlsx 的「测试结果/调试结果」列。
- **干的事**：模拟测试人员在 Acuview 2 里改设置→下发→回读确认，可重复、可断言。
- **两类用例**：
  - `write_verify`（配置下发）：界面设值 → Update 下发 → **跨传输**回读断言 → 自动还原。
  - `read_compare`（界面数据读取比对）：导航读数页 → OCR 取显示值 → 与 Modbus 真值/期望比对（需 Tesseract）。

---

## 2. 原理与架构

### 2.1 共享引擎 + 按设备数据

```
autotest/                                  ← 仓库根
├─ conftest.py             已把仓库根加入 sys.path → import comm.ctl_acuview 各项目可用
├─ comm/                   ← 共享层（全项目复用）
│   ├─ ctl_acuview/        引擎包（import comm.ctl_acuview）
│   │   ├─ config.py           读 config（点号访问；相对路径锚到 config 文件自身目录）
│   │   ├─ spec_loader.py      data 的 xlsx+json → 统一模型(registers/pages)
│   │   ├─ modbus_client.py    TCP/RTU 直读直写 + 类型编解码（校验真值源）
│   │   ├─ gui_driver.py       GUI 驱动：widget_abs(JSON坐标→屏幕)/模板匹配/设值/锁屏护栏
│   │   ├─ app_locator.py      Acuview2 exe 自动发现（跨 PC 免配置）
│   │   ├─ calibrate.py        标定 content_origin（JSON 逻辑坐标→屏幕物理像素）
│   │   ├─ verify.py           跨传输闭环断言 + 报告(json/csv)
│   │   ├─ testcase_engine.py  通用引擎：run_write_verify_case / run_read_compare_case / navigate_to / PAGE_NAV
│   │   ├─ find_register.py    CLI：按关键词搜寄存器
│   │   ├─ find_widget.py      CLI：按关键词搜控件（反查寄存器地址）
│   │   ├─ manual_xlsx.py      手工用例 xlsx 读取 + 回填
│   │   └─ dpi.py / uia_probe.py
│   └─ templates_acuview/  Acuview2 导航树 + 对话框按钮模板（OpenCV 匹配，引擎按 ctl_acuview/../templates_acuview 解析 = comm/templates_acuview）
└─ projects/<项目>/         ← 按设备/项目层（各自独立）
    ├─ data_acuview/        Modbus 地址表 xlsx + Acuview 控件 JSON（用户提供）
    ├─ spec_acuview/        spec_loader 生成（registers.json / pages.json；可重建，gitignore）
    ├─ config_acuview.yaml  本项目 IP/COM/connection_row/content_origin/safety/spec 相对路径
    ├─ reports_acuview/     运行报告 json/csv/截图（可重生成）
    └─ tests/acuview_auto/  生成的 test_acuview_<编号>.py（+ manual_testcase/ 放手工用例输入）
```

**为什么多项目能共享、又互不干扰**
- `import comm.ctl_acuview` 各项目零配置可用：仓库根 `conftest.py` 已把仓库根加入 `sys.path`（`comm` 是命名空间包，`comm.ctl_acuview` 为其子包）。
- 数据隔离：每条用例传 `config_path=<项目>/config_acuview.yaml`；`comm/ctl_acuview/config.py` 的 `resolve()` 把相对路径锚到该 config 自身目录，于是各项目 `data_acuview/`、`spec_acuview/`、报告各走各的。
- 模板共享：`comm/templates_acuview/` 是 Acuview 2 软件本身的界面元素（导航树、对话框按钮），与具体设备无关，全项目共用；引擎按"引擎包同级目录"解析它（`comm/ctl_acuview/../templates_acuview`），故两者必须同在 `comm/` 下。按设备不同的是 `data_acuview/` 里的控件 JSON 与寄存器表。

### 2.2 GUI 怎么定位控件（重点）

控件 JSON（如 `Default_xxx.json`）含全部控件的**逻辑坐标**(`basicAttribute.x/y/w/h`)，被 `spec_loader` 抽进 `spec_acuview/pages.json`。实测 **JSON 逻辑坐标 → 屏幕物理像素是 1:1**（scale=1.000）：

```python
x, y = drv.widget_abs("General", "Backlight_Value_Combo")   # JSON 坐标 -> 屏幕坐标
# widget_abs = content_origin + (x + w/2, y + h/2) - scroll_y
```

输入框/下拉/开关等**页面控件一律用此法定位**。仅**左侧导航树**和 **Update/确认对话框按钮**用 `comm/templates_acuview/*.png` 模板匹配（它们不在页面 JSON 控件表里）。

`content_origin` = 页面内容区左上角相对窗口左上的偏移，写在项目 `config_acuview.yaml > gui_backend.content_origin`。基线（1920×1080 / 125% / 最大化）：`fixed_x=294, fixed_y=144`。

### 2.3 闭环为什么走两路传输

Acuview 2 占用 **TCP**（单客户端）时，校验脚本若再开 TCP 会让 Acuview 写入 `Update Failed`。故 **GUI 走 TCP、校验走 RTU**，跨传输互不争用。这就是 `config` 里 `transport.gui` 与 `transport.verify` 必须错开的原因。

---

## 3. 调用流程

```
工程师: "把这份手工用例转成自动化"  （Claude Code 中）
        │  触发 SKILL.md
        ▼
1) 确定目标项目 projects/<项目> → 用其 config_acuview.yaml
2) 建模型  python -m comm.ctl_acuview.spec_loader --build --config projects/<项目>/config_acuview.yaml
3) 读手工 xlsx（仅「是否可自动化==是」的行）
4) 逐条转换:
     判类型(write_verify / read_compare)
     → 抽设置项/值 或 被读项/期望
     → find_register / find_widget 定位寄存器/控件
     → 候选交工程师确认（人在环）
     → 安全校验(不在 forbid_write_*、值在 range 内、可逆)
5) 生成 test_acuview_<编号>.py（docstring 5 项 + 调引擎 + assert report.passed）
6) pytest 跑 → 回填 xlsx「测试结果」/失败回填「调试结果」
7) 汇总: 转了哪些 / 跳过哪些 / 跑挂哪些及原因
```

**人在环**：大白话→控件名常不唯一，第 4 步定位的候选必须给工程师确认，不臆测。

---

## 4. 使用方法

### 4.1 准备一个项目（以 RPP 为例，已建好骨架）

1. 把该项目的两份素材放进 `projects/RPP/data_acuview/`：
   - **Modbus 地址表 xlsx**
   - **Acuview 控件 JSON**（导航树/页面/控件坐标）
2. 编辑 `projects/RPP/config_acuview.yaml`：
   - `spec.excel` / `spec.json` 改成上面两个文件的实际文件名；
   - `transport.tcp.host` / `transport.rtu.port` 等改成 RPP 测试环境的 IP / COM；
   - `app.connection_row` 改成 Add Connection 列表里 RPP 那条连接的行号（从 0）；
   - `safety.forbid_write_addr` 按 RPP 寄存器表复核填写（name 子串护栏已通用）；
   - `gui_backend.content_origin` 用基线值；换分辨率/缩放再重标（见 4.3）。
3. 准备手工用例 xlsx，放 `projects/RPP/tests/acuview_auto/manual_testcase/`。

### 4.2 常用命令（一律从**仓库根**运行）

```powershell
# 读 spec 的 CLI 都要用 --config 指定项目(或设环境变量 ACUVIEW_CONFIG)；下例以 RPP 为例
# 建模型(离线，不连设备)
python -m comm.ctl_acuview.spec_loader --build --config projects/RPP/config_acuview.yaml

# 定位寄存器 / 控件
python -m comm.ctl_acuview.find_register --config projects/RPP/config_acuview.yaml "backlight"
python -m comm.ctl_acuview.find_widget  --config projects/RPP/config_acuview.yaml "password"

# 列手工用例 / 回填(manual_xlsx 直接吃 xlsx 路径，不需 --config)
python -m comm.ctl_acuview.manual_xlsx <手工xlsx>
python -m comm.ctl_acuview.manual_xlsx <手工xlsx> --case <编号> --debug "问题" --result 未通过

# 跑生成的用例(用例内已用 config_path 指向本项目 config，不需 --config)
pytest projects/RPP/tests/acuview_auto/ -v -s
pytest projects/RPP/tests/acuview_auto/test_acuview_<编号>.py -v -s
```

> 也可一次性设环境变量替代每条 `--config`：`$env:ACUVIEW_CONFIG='projects/RPP/config_acuview.yaml'`（PowerShell）。

### 4.3 重标 content_origin（换分辨率/缩放时）

```powershell
python -m comm.ctl_acuview.calibrate --config projects/RPP/config_acuview.yaml --page General --ox <值> --oy <值>
# 看 reports 下叠加图核对计划点位是否压在控件上，调好后把 ox/oy 写进项目 config 的 gui_backend.content_origin
```

### 4.4 新增一个项目

照搬 RPP 的目录约定即可：
```
projects/<新项目>/
  data_acuview/        放该设备 xlsx + 控件 JSON
  spec_acuview/        留空（spec_loader 生成）
  config_acuview.yaml  复制 RPP 的改 IP/COM/文件名/safety
  tests/acuview_auto/  生成用例落这里
```
引擎与模板无需复制（在 `comm/` 下共享）。

---

## 5. 环境与前置（GUI 写闭环必须满足）

- 上位机 = **Acuview 2**（Qt5，按用户安装在 `%USERPROFILE%\Acuview2\Acuview 2.exe`）。**UIA 仅暴露窗口外框**，内部控件全不可见 → 用坐标 + OpenCV 模板匹配驱动（写闭环不需 OCR）。
- exe 路径**跨 PC 自动发现**：`app.exe_path` 留空即可。
- **必须能解锁活动桌面**：锁屏/远程断开时截图点击不可用，引擎用 `is_session_locked()` 自动拒绝（用例转 skip）。
- 基线 **1920×1080 / 125% / 窗口最大化**；换环境需重标 `content_origin`、必要时重截模板。
- **写授权**：写设置需把默认密码 **`0000`** 填入 General 页 Password 框并回车（权限 View→Admin）。引擎每次写入前会重新提权（Update 后权限会回落）。
- 写确认流程：`Update` →（"Please Enter Password" 若 View）→ "Do you want to update?" → **Yes** → "Update Successful!"，引擎用模板匹配 `btn_confirm/btn_yes` 逐个点掉。
- **⏱ TCP 连接 5 分钟空闲自动断开**：断开后 GUI 写入会弹 `Update Failed`。GUI 导航不续命，只有真正读/写才刷新连接。**用例要连续跑、中途别留长空闲**；调试隔几分钟后先做一次读/写唤醒再写。
- Modbus 细节：实时浮点块按对齐的 2 寄存器读；**写用 FC16**（FC6 响应帧会让 pymodbus 解析报错）；字序大端(ABCD)。
- `read_compare` 才需 **Tesseract-OCR** 引擎，装好后设 `config.gui_backend.tesseract_cmd`。

依赖（已并入工程依赖清单）：`pymodbus / pyyaml / pywinauto / uiautomation / comtypes / pyautogui / opencv-python / pillow / pytesseract`。

---

## 6. 已知限制 / 排错

| 现象 / 限制 | 处理 |
|---|---|
| 报 `Write to Device failed` / 回读不变 | 多为 **TCP 5min 空闲断连**；连续重跑即可，别留长空闲 |
| 导航到某页失败、提示缺模板 | 该页未在 `comm/ctl_acuview/testcase_engine.py` 的 `PAGE_NAV` 注册或缺 `comm/templates_acuview/tree_<page>.png`；补注册 + 补树节点截图 |
| 换了分辨率/缩放后点击落空 | 重标 `content_origin`（见 4.3），必要时重截 `comm/templates_acuview/*.png` |
| `read_compare` 记 FAIL/"需 OCR" | 未装 Tesseract；装好设 `tesseract_cmd`，否则该类用例转 skip 并回填「调试结果」 |
| 写目标定位不到名字 | 名称匹配率约 29%（Excel 人类描述 vs JSON UPPER_SNAKE）；**读取走地址不受影响**，写靶定位用 `find_register/find_widget` 给候选人工确认，可维护别名表提升 |
| trivial 写（目标值==当前值）回读"必然通过" | 验证务必用会真正改变的值，否则断言无意义 |
| spinBox 标定偏差 | 标定用满宽 lineEdit 的白框中心更可靠（spinBox 右侧箭头会截断白区） |

---

## 7. FAQ / 维护

- **加一个新导航页**：在 `comm/ctl_acuview/testcase_engine.py` 的 `PAGE_NAV` 注册该页的 tab/tree/expand/landmark，并在 `comm/templates_acuview/` 补 `tree_<page>.png`（必要时 `tab_*`/`landmark`）。
- **扩一种用例类型**：在 `comm/ctl_acuview/testcase_engine.py` 加 `run_<type>_case(...)`，SKILL.md 第 4 步「判类型」加识别词。
- **模板 / 标定归谁维护**：`comm/templates_acuview/`（Acuview2 界面元素）与 `content_origin` 标定是**按操作机/分辨率**的共享资产，换工位由该工位负责重标/重截；放 `comm/` 下共享，不要每项目复制。
- **spec 要不要入库**：`spec_acuview/` 由 `data_acuview/` 重建，默认 gitignore；想开箱即用可自行入库。
- **引擎是共享的**：改 `comm/ctl_acuview/` 影响所有项目，改动要回归各项目；包内用相对 import，迁移/改名时无需动内部 import。
```（注：引擎与模板必须保持在 comm/ 下同级，否则 gui_driver/testcase_engine 解析模板目录 `ctl_acuview/../templates_acuview` 会失败。）```
