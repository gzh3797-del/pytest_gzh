# 团队编码约定

## 通用
- 所有自动化脚本使用 Python 3.10+
- IO密集型操作使用 asyncio，不用 threading
- 配置集中：所有可调参数放 config.py，不散落在业务代码中
- 报告格式：HTML，包含可折叠分节（details/summary），无需 JS

## Modbus 读取约定
- 默认使用 FC03（Holding Registers），Float32，Big Endian，2寄存器/参数
- FC02（Discrete Inputs）暂不支持，DI型设备（IOM-03/04）创建空 stub 并注释说明
- 地址表模块命名：devices/<设备名小写>.py
- 必须实现 build_param_map() → dict[param_key, ModbusRegister]
- 支持 AcuCloud 的设备还须实现 build_cloud_col_map() → dict[xlsx列标题, param_key]

## BACnet 约定
- 只处理 Analog Input 对象（BACnet AI），BACnet Binary Input 暂不支持
- BACnet 对象名（object_name）= param_key
- 单位检查：比对 BACnet units 属性 vs 模板 unit 列（不检查 description）
- 比对容差：|diff| ≤ max(TOLERANCE_ABSOLUTE, ref × TOLERANCE_PERCENT/100)

## AcuCloud 约定
- xlsx 文件名须与设备名精确匹配（如 AcuvimIIW.xlsx）
- build_cloud_col_map() 返回 {xlsx列标题: param_key}
- AcuvimIIR 的 cloud_col_map 与 AcuvimIIW 完全相同，直接 import 复用
- 比对容差：±5% / ±1.0（补偿时序差异）

## 新增设备检查清单
- [ ] devices/<name>.py — build_param_map()
- [ ] devices/<name>.py — build_cloud_col_map()（若支持Cloud）
- [ ] config.py — MODBUS_DEVICE_MAP 添加条目
- [ ] comparator.py — _DEVICE_MAP 添加条目
- [ ] cloud_comparator.py — _DEVICE_MAP 添加条目（若支持Cloud）
- [ ] README.md — 支持设备表更新
- [ ] shared/devices/<name>.md — 设备知识文件
- [ ] shared/modbus_tables/INDEX.md — 更新索引
- [ ] shared/templates/INDEX.md — 更新索引（若有模板）

## Playwright UI 测试编码约定
- **交互优先用 `locator` + `expect` 风格**，保留 Playwright 内置的 auto-wait 与 actionability 检查（可见性/遮挡/启用状态）；禁止默认使用 `page.evaluate` + JS 原生 `click()` 或 `force=True` 绕过检查
- **降级到坐标点击（`getBoundingClientRect` + `mouse.click`）仅限一种情况**：Element Plus 组件确实拦截事件链（如 el-radio 的合成事件，JS `element.click()` 不触发完整事件链）；降级时必须先 `scrollIntoView({block:'center'})` 再取坐标，防止前序用例改变滚动位置后坐标落在视口外、点击静默失败
- **JS 循环点按钮必须加可见性过滤** `btn.offsetParent !== null`：Element Plus 弹窗关闭后仍留在 DOM，其隐藏按钮在 DOM 顺序上可能先于主区按钮，误点会导致操作静默失效
- `page.evaluate` 内的原生 `document.querySelector` 不支持 Playwright 专有伪类（如 `:has-text`），混用会 SyntaxError 炸掉整个 evaluate
- 测试体内禁止直接 `asyncio.run()`（Playwright sync API 主线程持有运行中事件循环），同步包装协程用 `_run_coro` 独立线程模式（参考 `test_case/acuhmi_1_7/bacnet_ui/helpers/hmi_bacnet_client.py`）
- 背景：2026-06 HMI1-7 bacnet_ui 排查结论，4 个失败根因中 3 个源于绕过 actionability 检查的写法
- **选择器沉淀复用（防止重复现场探查）**：某页面的选择器/交互一旦探明，固化到「选择器沉淀文档」；后续写/调试同页 UI 用例前必须先查、命中即复用。
  - 沉淀文档目录（通用，跨项目）：`<项目知识库根>/requirements/context/`
    - `<项目知识库根>` = CLAUDE.md「项目一览」中该项目 `context.md` 所在目录（如 AcuHMI-1-7→`knowledge/gateway/AcuHMI17/`、AcuRev-4100-WEB2→`knowledge/gateway/AcuRev4100WEB2/`、AcuRev1320→`knowledge/meters/AcuRev1320/`）
  - **定位规则（先查索引，不再按测试目录名直接拼文件名）**：沉淀文档以**被测页面的 Web 导航菜单路径**命名并按可路由子页拆分（PascalCase，如 `Devices_DataLog_DataLogger_context.md`、`SystemSettings_Network_context.md`、`Protocols_MQTT_SSL_TLS_context.md`），一子页一份，与 `projects/<项目>/tests/ui/` 的一级目录名（小写，如 `datalog`）是多对多关系。因此**必须先读该目录下的 `_INDEX_context.md`**（Web 菜单路径 → 子页文档的全量清单），据此找到被测页面对应的 `<Prefix_SubPage>_context.md`，禁止用测试目录名直接拼文件名。若项目暂无 `_INDEX_context.md`，回退到「按被测页面菜单名在该目录内检索匹配文件」。
  - **编排约定（谁调 UI 探查 agent 谁负责）**：派 `ui-test-engineer` 做现场探查前，主 AI 必须先查上述文档；命中则把路径/内容随 prompt 传给 agent 并要求优先复用，仅缺失部分才现场探查。agent 探明的新选择器由**主 AI（或知识库维护）沉淀回** `requirements/context/`——`ui-test-engineer` 自身不写 `knowledge/`（职责分离）。
  - 新增沉淀文档后，在该项目 `context.md` 就近加一行指针，便于发现。
- **Playwright 浏览器版本同步**：本机存在多个 venv（如 `PycharmProjects\pythonProject\.venv`、`Desktop\testing-team\.venv`、`Desktop\方案设计\autotest\.venv`），各自的 playwright 包版本可能不同，要求的 Chromium 版本号也不同；浏览器是全局安装到 `%LOCALAPPDATA%\ms-playwright`。因此每次对某个 venv 执行 `pip install -U playwright`（或新建/切换 venv）后，必须在**同一个 venv** 下紧跟一次 `python -m playwright install chromium`，让包与浏览器版本保持一致，否则 UI 用例启动时会报 `Executable doesn't exist at ...chromium-XXXX`。仓库根 `conftest.py` 已加「Playwright 浏览器自检」钩子：收集到 UI 用例（nodeid 含 `/ui/`）时会校验当前解释器期望的 Chromium 是否存在，缺失则 fail fast 并打印含本 venv python 路径的精确安装命令；该自检不替代上述同步动作，只是把晦涩堆栈变成可照抄的命令

## 图像匹配模板命名约定（pyautogui）
- **文件名及路径必须全为纯 ASCII 字符**，禁止出现中文或任何非 ASCII 字符
- **根因**：pyautogui `confidence` 匹配底层走 pyscreeze → OpenCV `cv2.imread`；`cv2.imread` 在 Windows 上按系统 ANSI 码页打开文件，无法处理含中文的路径——会将中文字符错码（如 `下拉展开态` → `涓嬫媺灞…`）并返回 `None` 或抛 "file is missing"。该模板随即**永远匹配失败**。`check_image_exists` 等封装通常吞掉异常返回 `False`，导致**静默退化**（等待逻辑形同虚设、点击落空），且 `os.path.exists` 与 PIL 均可正常读取，只有 cv2 读不了，排查时极难发现
- **命名规范**：用英文单词 + 下划线（`Connect_Disabled.png`、`BaudRate_DropDown.png`），中文语义写在代码注释里，不要进文件名
- **排查指引**：遇到"图像匹配莫名失败但文件确实存在"，优先检查模板路径是否含非 ASCII 字符
- **历史案例**：AcuRev1320 固件升级 RTU 用例中，`BaudRate 下拉展开态.png`、`Connect 禁用态.png`、`Select All 禁用态.png` 三个模板因名含中文导致波特率下拉检测持续失败；改为 `BaudRate_DropDown.png`、`Connect_Disabled.png`、`Select_All_Disabled.png` 后修复

## 文档数字描述维护约定
- README / 知识文档中凡出现可枚举数字描述（如"N段式"、"N台设备"、"N个检查项"），修改后必须全文搜索旧值，确认出现次数为 0 再提交
- 同一数字描述通常散落在：**章节标题**、**正文首句**、**目录结构注释**、**代码块注释** 四处，只改其中一处是不够的
- **同时搜索派生数字**：若总数为 N，文档中可能还存在 N-1 的派生表述（如"其余N-1段仍完整执行"）；更新 N 时必须一并搜索 N-1（旧值）并更新为新的 N-1
- 适用场景：脚本新增检查段、新增设备、新增协议，凡改了"段数/台数/项数"均触发此规则


## pytest 用例命名约定
- **一函数一用例，函数名必须嵌入对应的用例编号**，确保测试报告中每个函数名能直接回溯到知识库手工用例的具体编号
- **禁止**用泛化函数名 + `@pytest.mark.parametrize` 把多条不同编号的用例合并进同一个函数——参数化合并后函数名无法对应到具体用例编号，覆盖率回溯和缺陷定位均失效
- **命名格式**（沿用 AcuRev1320/TOU 既有风格）：`test_<模块号>_<子模块号>_case<编号>`
  - 示例：`test_013_01_case01`、`test_013_01_case02`
- 作用域：工程下所有项目，不限某一项目

## 用例编号一一对应约定
- 测试代码/报告中出现的用例编号（pytest 参数化 ID、测试文件名、类名中的 `TestCase_xxx` / `Testcase_xxx`）必须与测试用例 Excel「用例编号」列**逐字符严格一致**，一编号一用例
- 不允许的差异类型（2026-07 AcuHMI-1-7 回填时实际发现）：
  - 项目前缀不一致（代码 `TestCase_WEB2_AWS_*` vs 用例表 `TestCase_AcuHMI-1-7_AWS_*`；`ACUREV4100WEB2_VD_*` vs `AcuHMI_VD_*`）
  - 编号与名称顺序颠倒（代码 `AcuvimIIW_CT41_043`（名称_编号）vs 用例表 `AcuvimIIW_005_CT41`（编号_名称））
  - FTS 段编号不一致（代码 `008_05_caseNN` vs 用例表 `008_03_caseNN`）
  - 分隔符不一致（`ARM_XXL` vs `ARM-XXL`）、冗余后缀（`_row598`、`_v2`）
- 回填测试报告结果到用例表时，只写结果相关列，**禁止改动用例表「用例编号」列**
- 发现编号不匹配：整理清单（报告编号 ↔ 用例表编号 ↔ 差异类型 ↔ 处理建议）交相关负责人，统一修改**代码侧** ID 向用例表对齐；用例表中无对应行的测试由负责人决定补建用例行或清理测试
- 作用域：工程下所有项目，不限某一项目

## 知识库维护约定
- 每次 Jira 导出后更新对应项目 bugs/INDEX.md（5分钟内完成）
- 新需求文档下来后，2个工作日内完成 requirements/summaries/ 摘要
- 重要决策（为什么这样设计）记入 shared/decisions.md
