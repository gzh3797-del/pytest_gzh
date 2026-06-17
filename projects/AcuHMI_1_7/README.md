# AcuHMI-1-7 测试项目

本文件只做**模块索引 + 运行命令**；各模块的详细说明（用例清单、前置条件、报告）见对应模块目录下的独立 README。

- **配置**：`configs/global.yaml`（全局）+ 本项目 `config.yaml`（项目差异、显示名）+ `configs/.env`（敏感值，模板见 `configs/.env.example`）
- **报告**：`reports/acuhmi_1_7/<时间戳>/`
- `pages/` Page Object ｜ `helpers/` 项目内客户端/匹配器 ｜ `data/` 测试数据/模板

## 模块

| 模块 | 路径 | 内容 | 详细 README |
|---|---|---|---|
| BACnet/IP UI | `tests/bacnet/` | BACnet/IP 配置页功能 + 协议端到端验证（45 用例，自动化 36） | [`tests/bacnet/README.md`](tests/bacnet/README.md) |
| 接线检查 Wiring Check | `tests/wiring_check/` | 接线状态推导与实测核验 | 待补 |
| Web UI | `tests/ui/` | 各页面功能用例（about / datalog / systemsettings / physicaldevices / webdevices / security / usermanagement / ... 多子模块） | 待补 |

## 运行命令（仓库根目录）

```bash
# 整项目（走 run.py：注入报告目录 + pytest-html，报告落 reports/AcuHMI_1_7/<时间戳>/）
python run.py AcuHMI_1_7
python run.py AcuHMI_1_7 -m smoke          # 按标记过滤

# 单个模块（直接 pytest）
pytest projects/AcuHMI_1_7/tests/wiring_check/ -v
pytest projects/AcuHMI_1_7/tests/ui/datalog/ -v        # ui 下指定子模块
pytest projects/AcuHMI_1_7/tests/ui/ -v                # 全部 Web UI

# 单条用例
pytest projects/AcuHMI_1_7/tests/ui/ -k "AcuHMI_009_08_case01" -v
```

> BACnet 模块的额外依赖（BAC0）、用例清单与运行细节见 [`tests/bacnet/README.md`](tests/bacnet/README.md)。
