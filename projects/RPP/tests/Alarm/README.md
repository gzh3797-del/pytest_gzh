  # Alarm（Alarm Config + Alarm Logs 用例组）自动化测试

RPP 项目「接入设备日志管理」下 **Alarm Config（22 条，TestCase_ACUHMI17_BZ_xxx）**
与 **Alarm Logs（8 条，TestCase_AcuHMI_002_02_casexx）** 手工用例的自动化实现
（用例来源：`RPP项目软件测试方案过程件及用例v1.3.xlsx` → 测试用例(北向) sheet）。

> **当前按 AcuHMI-1-7 真机执行**（RPP 固件尚未提供 Alarm 页面）。
> 目标机 `https://192.168.3.71`，账号走 `configs/.env` 的 `WEB_USERNAME` / `WEB_PASSWORD`。
> RPP 真机就绪后只需改 `config_alarm.py`（或设环境变量 `ALARM_BASE_URL` / `ALARM_TRIGGER_DEVICE`）。

## 目录结构

```
Alarm/
├── README.md            # 本文件
├── config_alarm.py      # 模块配置（目标机地址 / 触发设备 / 轮询等待参数）
├── conftest.py          # fixtures（复用项目根 browser，自建 context + 登录，session 兜底清理）
├── helpers_alarm.py     # 导航 / 告警规则 CRUD / 触发等待 / 确认操作封装
└── test_TestCase_ACUHMI17_BZ_xxx.py   # 22 条用例，一文件一用例，函数名=用例编号
```

## 告警触发方式

在轮询 **Status=ON** 的下挂表（默认 `AcuRev4100_392`，AcuRev-4110-mV）的
`Alarm → Alarm Config` 创建必越限规则（如 System Frequency Min=70/Max=90 →
实际 ~50Hz → UNDERFLOW），网关约 60s 轮询一次，一个周期内触发。
本模块创建的规则统一 `at_` 前缀，用例 finally + session 兜底双重清理。

## 执行

```powershell
# 仓库根目录执行；全量约 25~35 分钟（触发类用例每条需等 1~2 个轮询周期）
python -m pytest projects/RPP/tests/Alarm -v
# 单条
python -m pytest projects/RPP/tests/Alarm/test_TestCase_ACUHMI17_BZ_003_001.py -v
```

## 覆盖情况（2026-07-17 实跑：20 passed + 2 skipped）

| 用例 | 结果 | 说明 |
|------|------|------|
| BZ_001_001 | ✅ | 导航改名 Alarm ✓；数量角标**实装在 Unacknowledged Alarms 二级 tab**，非用例文字所述"左导航 Alarm 括号内"，脚本按实装位置校验数量一致性，位置差异待需求方判定 |
| BZ_001_002 | ✅ | 子页面为 Unacknowledged Alarms + Alarm Logs |
| BZ_001_003 | ✅ | 开关在 System Settings → Alarm Notification，默认 Enable |
| BZ_002_001~003 | ✅ | 蜂鸣器发声/停止为听觉验证，**需人工现场确认**；脚本断言 Web 侧可见效果（Ack Status、告警 ON→OFF 状态迁移） |
| BZ_002_004 | ⏭ skip | 物理面板 UI，需人工目视 |
| BZ_003_001~005 | ✅ | 实测：Disable 时 Unacknowledged Alarms tab **整体隐藏**（等价"无确认入口"）；Ack Status 列随开关显隐 |
| BZ_004_001~004 | ✅ | 8 列展示 / 确认后消失 / 确认日志 / 空列表 |
| BZ_005_001 | ✅ | Ack Status 列 Enable 显示、Disable 隐藏 |
| BZ_005_002 | ✅ | 覆盖"未配置 Trigger → 显示 -"分支；**配置 DO/RO 设备分支需台架接入 AcuIOM 后补充** |
| BZ_005_003 | ✅ | 设备详情页 Alarm Logs 含 Ack Status |
| BZ_006_001~002 | ✅ | 多告警独立确认 / 开关切换历史日志不丢失 |
| BZ_006_003 | ⏭ skip | 需被监控设备真实离线（断电/拔线），需人工配合 |

## Alarm Logs 检索用例组（002_02，2026-07-17 实跑：8/8 passed）

| 用例 | 结果 | 说明 |
|------|------|------|
| case01 | ✅ | Interval 检索（含未来区间检索为空的负向验证） |
| case02~06 | ✅ | Serial Number / Monitor ID 及组合条件检索，断言结果各列全部匹配过滤条件 |
| case07 | ✅ | **Clear Logs 破坏性**：清空全部告警日志并断言立即消失，务必放本组最后执行 |
| case08 | ✅ | Reset 后四个搜索框全部恢复空/占位 |

- 测试数据由 `ensure_alarm_log_data()` 前置保障：日志缺 `at_alog1` 记录时现场
  触发一次（约 60s）→ 确认 → 删规则，日志记录保留供各用例复用；case07 清空后
  其他用例重跑会自动重建。
- Serial Number / Monitor ID 均从日志行实时读取（规则重建后 Monitor ID 会变）。

## 实测得到的关键页面事实（写脚本时依赖）

- Alarm Config 的 **Status 列是图标**：ON=`el-icon warning`、正常=`el-icon success`（无文本）
- Add Alarm 是**内联表单页**（URL `?type=add`），非弹窗；行内主色按钮=编辑、危险色=删除
- Alarm Notification 页：`Alarm Email Enable=Enable` 且收件人为空时 Save 会被必填校验拦下，
  `set_ack_enable` 已做规避（先关 Email）
- 页面常驻隐藏的日期面板（el-picker-panel）内也有 OK 按钮，确认弹窗操作必须过滤可见性
- Alarm Logs 的 Interval（datetimerange）**直接向输入框键入文本不会同步组件内部值**
  （面板一关即回滚、Search 按空条件执行且不报错，极易产生"检索通过"假象）；
  必须通过日期面板点选日期 + footer OK 确认（见 `set_interval_filter`），
  Escape 会把未确认输入整体回滚清空
