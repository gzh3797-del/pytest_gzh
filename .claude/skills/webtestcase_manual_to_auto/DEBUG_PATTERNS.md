# 自动调试错误模式 → 修复策略（DEBUG_PATTERNS）

阶段 2 的自动调试循环对每条失败用例最多 **3 轮**。每轮：读 pytest `--tb=short` 错误 →
读测试文件 → 按下表匹配 → Edit 修改 → 单条重跑 `python -m pytest <file> --tb=short -q`。
通过则标 `是`；3 轮仍失败标 `调试失败|原因`。

> 修复必须基于**证据**：不确定实际文案/结构时，先 `page.content()` 或截图确认，再改，禁止盲猜。
>
> **框架级坑（Element Plus / Ant Design、SPA 路由守卫、同路由 goto、依赖字段、跨页传播）的完整成因与探查方法见
> `.claude/skills/webpage_analyze/PITFALLS.md`**（跨组沉淀，权威版）。下表标注 `PITFALLS §N` 的行是其调试期速查映射；
> 命中后应把确认的选择器/坑增量沉淀回对应 `*_context.md`（SKILL.md Step 6）。

## 错误模式表

| 错误特征 | 判断 | 修复策略 |
|---------|------|---------|
| `TimeoutError: ...locator("a").first` | 列表行无 `<a>` 链接 | 改 `locator("td").first.click()`；进详情后若找不到表单，先点 `.el-tabs__item` 对应 tab |
| `StrictModeViolation: resolved to N elements` | 定位到多个元素 | 末尾加 `.first`，或 `.filter(has_text="...")` 收敛 |
| `TimeoutError: get_by_role("button", name="Yes, continue")` | 确认按钮文案不同 | 依次试 `"Yes"` → `"Yes,continue"`（无空格）→ `"Confirm"` |
| `TimeoutError` 且在 popconfirm 中 | popconfirm 用 primary button | 改 `page.locator(".el-popconfirm__action .el-button--primary").click()` |
| `TimeoutError: get_by_placeholder("X")` | placeholder 与实际不符 | `page.content()`/截图找实际 placeholder 后修正 |
| `AssertionError: .el-form-item__error count == 0` | 产品未做该输入校验 | 加 `@pytest.mark.xfail(strict=False, reason="产品未校验该字段")` |
| 删除后元素仍存在 | 删除后未等刷新 | 删除确认后加 `page.wait_for_timeout(1000)` 再查询 |
| 结果应为空却显示 N 行 | 产品不过滤该边界值 | 加 `@pytest.mark.xfail(strict=False, reason="产品不过滤该参数边界值")` |
| `TimeoutError: filter(has_text="Enable")` | 开关名称不同 | 截图确认实际文案；或 `.el-form-item` filter → `.el-radio` filter 定位 |
| `expect(...).to_have_value()` 元素找不到 | 字段在未激活 tab 内 | 操作前先点对应 tab |
| `TimeoutError` 在功能入口点击 | 导航路径不对 | 对照 context「进入路径」重确认；补 `wait_for_load_state("networkidle")` |
| 「已达上限」重复添加 | 前置数据残留 | try 前加清理，或先统计已有条目再计算增量 |
| `ImportError` / `ModuleNotFoundError` | 依赖/路径错误 | 对齐同模块其他测试文件的 import；确认 `projects.<PROJECT>...` 路径 |
| `fixture 'login_page' not found` | conftest 未覆盖 | 确认从仓库根跑 pytest；`login_page` 由 `projects/<PROJECT>/conftest.py` 提供 |
| `SSL` / 连接错误 | BASE_URL/设备连通 | 检查 `settings.BASE_URL` 与设备可达性（不可达属环境问题，标调试失败） |
| `xfail strict=True but passed` | XPASS | 改 `strict=False` |
| `combobox`/下拉选不中，`.el-select__wrapper` 找不到 | 框架认错（Ant vs El Plus） | `browser_evaluate` 跑 `elSelectCount/antSelectCount` 计数；El→`.el-select__wrapper`+`.el-select-dropdown__item`，Ant→`.ant-select`+`.ant-select-item-option-content`（见 PITFALLS §3） |
| 弹窗内 `.el-select-dropdown__item` 找不到 | El 下拉 teleport 到 body | select 在 `.el-dialog` 内点，选项在 body 层 `.el-select-dropdown__item:visible` 找（PITFALLS §3） |
| 读下拉当前值得到空串 | `.el-select__selected-item` 双节点 | 遍历节点跳过空 `inner_text`，兜底读 `.el-select__placeholder`，禁用 `.first`（PITFALLS §3） |
| `el-dialog` 关闭后断言仍见弹窗 | 用了 `wait_for_timeout` | 改 `expect(page.locator(".el-dialog")).to_be_hidden(timeout=8000)` |
| Ant 表格加载态误判行存在 | 用了通用 `tr` | 行选择器改 `tr.ant-table-row`（PITFALLS §3） |
| `nav_to_*` 后落错页/`goto` 被重定向 | SPA 路由守卫 | 改菜单点击链（sidebar listitem→menuitem），末尾加 `assert "关键字" in page.url`；禁多次 goto 重试（PITFALLS §1） |
| 读到"最后一页/筛选后"数据，断言错 | 同路由 goto 空操作未重置状态 | nav 函数显式重置（点上一页/清筛选）或"goto 别路由→goto 回来"强制重挂载（PITFALLS §2） |
| 依赖字段留空却不报错 | 级联字段未触发时跳过校验 | 先选依赖字段(A)再操作目标字段(B)；断言前确认 B 已启用/回填（PITFALLS §4） |
| 跨页断言对着旧数据过/挂 | 写入传播延迟 | A 拿到成功信号后轮询 B，`expect(...).to_contain_text(timeout=实测值)`，勿 `wait_for_timeout` 后一次性查（PITFALLS §5） |

## 停止自动修复（直接标结论）

满足任一：
- 涉及 Factory Reset / 固件升级 / 设备重启等破坏性操作 → `跳过|<reason>`
- 需真实硬件/外部设备复现 → `跳过|<reason>`
- 3 轮后错误类型未变（根因未识别）→ `调试失败|<最终错误首行>`
- 错误信息过于通用、修复方向不明 → `调试失败|需人工排查:<摘要>`

判定为**产品行为与规格不符**（非脚本问题）时用 xfail 并在结果列写 `调试失败|<产品行为说明>`，
供人工判断是否为真实缺陷。

## 每轮进度输出示例

```
[自动调试] test_TestCase_AcuHMI_005_02_case02_02 — 第 1 轮
  错误：TimeoutError on get_by_placeholder("Enter IP").nth(1)
  识别：Ethernet 2 在 Auto 模式下 IP 字段被禁用
  修复：填 IP 前先切 Manual
  重跑：PASSED ✓ → 标记「是」
```
