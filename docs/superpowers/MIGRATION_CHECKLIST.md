# 项目迁移 Checklist（每个旧项目重复执行）

参考已完成样板：`projects/acuhmi_1_7/`。

1. 建 `projects/<snake_name>/{tests,pages,helpers,data}` + `__init__.py`，写项目 `README.md`
2. 写 `projects/<snake_name>/config.yaml`（`project_name` 显示名 + 项目特有项；敏感值进 `configs/.env`，模板见 `configs/.env.example`）
3. 建 `projects/<snake_name>/settings.py` 适配层，调用 `framework.config.load_config("<snake_name>")` 暴露本项目所需常量
4. 用 `git mv` 迁移测试/helpers/data，按 `tests/<模块>` 分子目录；保留必要的 `__init__.py` 包结构
5. 改 import（逐一改到 grep 无残留）：
   - `from <旧项目>.config import X` → `from projects.<snake>.settings import X`
   - `from config.settings import X` → `from projects.<snake>.settings import X`
   - `from comm...` / `from modbus_config import ...` → **保持不变**（`comm/` 与 `modbus_config.py` 都在仓库根，可直接导入；连接参数读 `configs/global.yaml`）
   - `from test_case.<旧项目>...` → 改为 `from projects.<本项目>...`（拷入 `projects/` 后旧的 `test_case.` 绝对导入需逐一改到新路径）
   - `from Protocols...` → `from tools.Protocols...`（协议引擎已移入 `tools/`，脚本统一 `python tools/Protocols/...`）
   - 迁移后路径变深，注意修正测试内 `sys.path` 的 `parents[N]` 深度
6. 改造项目 `conftest.py`：去掉自带 `load_dotenv`（settings 适配层已加载），保留失败截图钩子并让其读 `SCREENSHOT_DIR`
7. 截图/报告只入 `reports/`；旧 `screenshots/` 用 `git rm --cached` 后物理删
8. **验证收集**：`python -m pytest projects/<snake> --collect-only -q` —— 数量须等于旧树基线、0 error
   - ⚠️ 注意 Windows 上 `norecursedirs` 无分隔符模式按 basename **大小写不敏感** 匹配：若项目内有名为 `protocols`/`reports` 等的子包，确认未被根 `pytest.ini` 的 norecursedirs 误伤（用带分隔符的锚定模式，如 `autotest/tools/Protocols`）
9. 入口冒烟：`python run.py <snake> --collect-only -q`
10. 删旧树空壳；更新项目 README（目录树 + 用例表 + 命令 + 前置条件）
11. Commit
