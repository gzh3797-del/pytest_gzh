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
- **Playwright 浏览器版本同步**：本机存在多个 venv（如 `PycharmProjects\pythonProject\.venv`、`Desktop\testing-team\.venv`、`Desktop\方案设计\autotest\.venv`），各自的 playwright 包版本可能不同，要求的 Chromium 版本号也不同；浏览器是全局安装到 `%LOCALAPPDATA%\ms-playwright`。因此每次对某个 venv 执行 `pip install -U playwright`（或新建/切换 venv）后，必须在**同一个 venv** 下紧跟一次 `python -m playwright install chromium`，让包与浏览器版本保持一致，否则 UI 用例启动时会报 `Executable doesn't exist at ...chromium-XXXX`。仓库根 `conftest.py` 已加「Playwright 浏览器自检」钩子：收集到 UI 用例（nodeid 含 `/ui/`）时会校验当前解释器期望的 Chromium 是否存在，缺失则 fail fast 并打印含本 venv python 路径的精确安装命令；该自检不替代上述同步动作，只是把晦涩堆栈变成可照抄的命令

## 文档数字描述维护约定
- README / 知识文档中凡出现可枚举数字描述（如"N段式"、"N台设备"、"N个检查项"），修改后必须全文搜索旧值，确认出现次数为 0 再提交
- 同一数字描述通常散落在：**章节标题**、**正文首句**、**目录结构注释**、**代码块注释** 四处，只改其中一处是不够的
- **同时搜索派生数字**：若总数为 N，文档中可能还存在 N-1 的派生表述（如"其余N-1段仍完整执行"）；更新 N 时必须一并搜索 N-1（旧值）并更新为新的 N-1
- 适用场景：脚本新增检查段、新增设备、新增协议，凡改了"段数/台数/项数"均触发此规则


## 知识库维护约定
- 每次 Jira 导出后更新对应项目 bugs/INDEX.md（5分钟内完成）
- 新需求文档下来后，2个工作日内完成 requirements/summaries/ 摘要
- 重要决策（为什么这样设计）记入 shared/decisions.md
