---
name: acuview_auto_testcase_generate
description: 把电表手工用例 xlsx 转成 pytest 自动化用例(驱动 Acuview2 上位机配置下发/界面数据读取，Modbus 跨传输回读比对)。多项目共享同一套 comm/ctl_acuview 引擎，各项目自带 data_acuview/config_acuview.yaml。当用户说"把这份手工用例转成自动化""生成自动化用例""manual_testcase 转 pytest"等时使用。
---

# 手工用例 → 自动化用例 生成器（Acuview2 + Modbus）

把某项目 `manual_testcase/*.xlsx` 里的手工用例，转成 `projects/<项目>/tests/acuview_auto/test_acuview_<用例编号>.py` 的 pytest 用例（一条用例一个文件，文件名用完整用例编号，如 `test_acuview_Function_RPP_027_02_case1.py`），复用 `comm/ctl_acuview` 通用引擎。**人在环**：定位寄存器/控件时把候选给测试工程师确认。

> 详尽说明（原理/架构/排错/维护）见同目录 `README.md`。本文件是触发指令。

## 背景与约定（务必遵守）
- 引擎在 `comm/ctl_acuview/`（与 `comm/templates_acuview/` 同级），全项目共用；各项目数据/配置自带，互不干扰：
  - 设备数据：`projects/<项目>/data_acuview/`（Modbus 地址表 xlsx + Acuview 控件 JSON）
  - 配置：`projects/<项目>/config_acuview.yaml`（IP/COM/connection_row/content_origin/safety/spec 相对路径）
  - 生成模型：`projects/<项目>/spec_acuview/`（由 spec_loader 重建）
  - 生成用例：`projects/<项目>/tests/acuview_auto/`
- 手工用例是**大白话**，且常描述**电表 HMI 菜单路径**(如 `Setting—Maintenance—LCD Backlight Time`)；
  自动化走的是 **Acuview2 软件**(页面/控件/寄存器)，两者不一一对应——映射靠 spec 搜索 + 工程师确认。
- 用例类型(可扩展)：
  - `write_verify`(配置下发)：界面设值→下发→跨传输回读断言→**还原**。引擎 `run_write_verify_case`。
  - `read_compare`(界面数据获取比对)：导航读数页→OCR 取显示值→与 Modbus 真值/期望比对。引擎 `run_read_compare_case`(**需 Tesseract OCR**)。
- 安全：遵守 `config` 的 `safety.forbid_write_*`(禁写通信链路寄存器)；目标值须落在寄存器 `range` 内；选可逆值；用例自带还原。
- 基线：1920×1080 / 125% / 窗口最大化；`content_origin` 已标定写在项目 config；**TCP 5min 空闲断连 → 用例要连续跑**。
- 物理观察项(如背光亮灭)无寄存器可读 → 记 MANUAL 目视项，不计入自动判据。

## 步骤

1. **确定目标项目**：从会话/用例路径确定 `projects/<项目>`（默认 RPP）；用其 `config_acuview.yaml`。

2. **建模型**：`python -m comm.ctl_acuview.spec_loader --build --config projects/<项目>/config_acuview.yaml`（仓库根运行；读该 config 指向 `data_acuview/` 的 xlsx+json，产物落 `spec_acuview/`）。
   首次需把该项目的 Modbus 表 xlsx 与控件 JSON 放进 `data_acuview/`，并在 `config_acuview.yaml` 的 `spec.excel/json` 填实际文件名。
   > 所有读 spec 的 CLI 都用 `--config <项目>/config_acuview.yaml` 指定项目（或设环境变量 `ACUVIEW_CONFIG`）；不传则找仓库根 config.yaml（多项目下不存在）。

3. **读手工用例**：`python -m comm.ctl_acuview.manual_xlsx <手工xlsx>` 或代码 `from comm.ctl_acuview.manual_xlsx import read_cases, case_meta`。
   只处理 `是否可自动化==是` 的行(read_cases 默认已过滤)。

4. **逐条转换**（对每条用例）：
   1. **判类型**：步骤含"设置/下发/写"→`write_verify`；含"读取/查看/显示/核对/比对"→`read_compare`；含糊就问工程师。
   2. **定目标**：从「测试步骤/预期结果」抽出设置项与值(write)或被读项与期望(read)。
   3. **定位**：`python -m comm.ctl_acuview.find_register --config projects/<项目>/config_acuview.yaml "<关键词>"` 找寄存器(addr/dtype/rw/range)；
      `python -m comm.ctl_acuview.find_widget --config projects/<项目>/config_acuview.yaml "<关键词>"` 找 page/widget(并反查对应寄存器地址)。
   4. **🔴 CHECKPOINT · 确认**：把候选 register/page/widget 列给工程师确认(尤其大白话→控件名不唯一时)；未确认不进入第 5 步生成/下发。
   5. **安全校验**：寄存器不在 `forbid_write_*`、目标值在 `range` 内、可逆。

5. **生成用例**：每条用例写一个文件 `projects/<项目>/tests/acuview_auto/test_acuview_<用例编号>.py`：
   - 文件内一个 `def test_case():`，其 **docstring 头备注 5 项**：用例编号/标题/预置条件/测试步骤/预期结果。
   - 用例模板（**注意：导入纯用 `comm.ctl_acuview`，不再手插 sys.path**；仓库根 `conftest.py` 已把根加入 sys.path）：
     ```python
     import pytest
     from pathlib import Path
     from comm.ctl_acuview.gui_driver import is_session_locked
     from comm.ctl_acuview.testcase_engine import run_write_verify_case  # 或 run_read_compare_case

     # 用例在 projects/<项目>/tests/acuview_auto/，config 在项目根 → 上溯两级
     PROJECT_ROOT = Path(__file__).resolve().parents[2]      # acuview_auto -> tests -> <项目>
     TEST_CONFIG = str(PROJECT_ROOT / "config_acuview.yaml")

     pytestmark = pytest.mark.skipif(
         is_session_locked(),
         reason="会话锁屏/远程断开, 无法驱动 Acuview2 界面",
     )

     def test_case():
         """用例编号/标题/预置条件/测试步骤/预期结果 五项备注"""
         report = run_write_verify_case(
             case_meta={...},
             register=<addr>, page="<Page>", widget="<Widget>",
             target_value=<value>, physical_note="<物理项或空>",
             config_path=TEST_CONFIG,
         )
         assert report.passed
     ```
   - read 类用 `run_read_compare_case(case_meta, page, widget, register=或 expect=, config_path=TEST_CONFIG)`。

6. **运行 + 回填**：`pytest projects/<项目>/tests/acuview_auto/test_acuview_<用例编号>.py -v -s`（多条连续跑用 `pytest projects/<项目>/tests/acuview_auto/ -v -s`）。
   - 成功/失败 → `manual_xlsx.write_back(xlsx, 用例编号, result="通过"/"未通过")`(写「测试结果」列)。
   - **失败 → `manual_xlsx.write_back(xlsx, 用例编号, debug="<遇到的问题>")`**(写「调试结果」列)。

7. **汇总**：告诉工程师——转了哪些、跳过哪些(`是否可自动化==否` 或定位不了)、跑挂哪些及原因。

## 限制 / 需补的
- **导航**只覆盖 `comm/ctl_acuview/testcase_engine.py` 的 `PAGE_NAV` 里注册的页 + 有 `comm/templates_acuview/tree_<page>.png` 的页；
  新页需在 `PAGE_NAV` 注册或补树节点截图(读数页树多数暂无模板)。
- `read_compare` **依赖 Tesseract OCR**；未装时引擎记 FAIL/"需 OCR"，应转 skip 并回填「调试结果」。
- `comboBox` 选项设值依赖 OCR；`spinBox/lineEdit/ipEdit` 走键盘输入(不需 OCR)。
- 换分辨率/缩放后需重标 `content_origin`：`python -m comm.ctl_acuview.calibrate --config projects/<项目>/config_acuview.yaml --page <页> --ox <> --oy <>`，把值写进项目 `config_acuview.yaml`。

## 异常与兜底（if-then）

| 触发条件 | 一线处理 | 仍失败兜底 |
|---|---|---|
| spec 未建 / `data_acuview/` 缺地址表 xlsx 或控件 JSON | 先跑 `spec_loader --build`；缺原料 🛑 STOP 向工程师要文件 | 不得跳过建模直接猜地址 |
| `find_register`/`find_widget` 无结果 | 换关键词重搜（中英文、缩写、HMI 菜单词→Acuview 术语） | 仍无 → 记入"定位不了"清单报工程师，该条跳过，禁止硬编地址 |
| 候选多个不唯一 | 🔴 CHECKPOINT 列候选给工程师选定 | 未确认不生成、不下发 |
| 会话锁屏 / 远程断开 | 模板 `skipif(is_session_locked())` 已兜底，用例整体 skip | 需实跑时请工程师解锁后重跑 |
| Tesseract 未装（read_compare / comboBox 设值） | 用例转 skip 并 `write_back(debug="需 OCR")` | 装好 Tesseract 后重跑 |
| Modbus TCP 空闲 5min 断连 | 用例连续跑；断连报错先重连再续跑 | 仍断 → 查设备在线/端口占用后报工程师 |
| 页面不在 `PAGE_NAV` / 无 `tree_<page>.png` 模板 | 注册 PAGE_NAV 或补树节点截图 | 补不了 → 该条记"暂不可自动化"并回填调试结果 |
| OCR 取值乱码 / 点击落点偏移 | 核对基线 1920×1080/125%/最大化；跑 `calibrate` 重标 `content_origin` 写回项目 config | 仍偏 → 回填调试结果转人工 |
| 手工 xlsx 缺「是否可自动化」列 / 列结构不符 | 🛑 STOP 报工程师确认列结构 | 禁止自行猜列硬解析 |
| write_verify 回读值 ≠ 目标值 | 先排环境（重连后重读一次、确认读的是同一寄存器/scale） | 复现 → 按疑似产品 bug 回填 `debug` 并报工程师，不得改断言凑通过 |
| 还原步骤写失败（设备留在非默认状态） | 立即重试还原一次 | 仍失败 → 🛑 STOP 在汇总里置顶告警"设备 X 寄存器 Y 未还原(当前值 Z)"，待工程师手动恢复后再继续跑后续用例 |
| Acuview2 未启动 / 崩溃 / connection_row 连不上设备 | 启动/重启 Acuview2 并重连（锁屏另见上行） | 仍连不上 → 该批用例整体 skip 并回填"上位机不可用"，报工程师 |

## ❌ 反例黑名单（不要做）

以下任一命中即停止返工（安全类直接中止，不下发）：

| # | 反例 | 纠正 |
|---|------|------|
| 1 | 候选 register/page/widget 未经工程师确认就生成用例或下发 | 🔴 CHECKPOINT 确认后才进第 5 步（人在环铁律） |
| 2 | 写 `safety.forbid_write_*` 列出的通信链路寄存器 | 写了会断链失联 → 安全校验不过直接中止 |
| 3 | 目标值超出寄存器 `range`，或选了不可逆值 | 选 range 内可逆值，用例自带还原 |
| 4 | 用例文件里手插 `sys.path` | 仓库根 conftest.py 已加根路径 → 导入纯用 `comm.ctl_acuview` |
| 5 | 一个文件塞多条用例，或文件名不用完整用例编号 | 一条用例一个文件：`test_acuview_<完整用例编号>.py` |
| 6 | 物理观察项（如背光亮灭）计入自动断言 | 无寄存器可读 → 记 MANUAL 目视项，不计入判据 |
| 7 | 用例跑失败只报 pytest 输出，不回填手工 xlsx | `write_back(xlsx, 编号, debug=...)` 写「调试结果」列 |
| 8 | 多条用例间隔久跑（TCP 5min 空闲断连） | 连续跑，断连后需重连再跑 |

## 相关文件
- 引擎：`comm/ctl_acuview/testcase_engine.py`（`run_write_verify_case` / `run_read_compare_case` / `navigate_to` / `PAGE_NAV`）
- 定位：`comm/ctl_acuview/find_register.py`、`comm/ctl_acuview/find_widget.py`
- 用例 I/O：`comm/ctl_acuview/manual_xlsx.py`（`read_cases` / `case_meta` / `write_back`）
- 模板：`comm/templates_acuview/*.png`
- 项目侧：`projects/<项目>/{data_acuview, spec_acuview, config_acuview.yaml, tests/acuview_auto}`
