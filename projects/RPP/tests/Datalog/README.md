# Datalog（接入设备 Data Log 下载用例组）自动化测试

RPP 项目「接入设备日志管理 / Datalog」子模块 5 条手工用例
（TestCase_AcuHMI_001_05_case01 / 01_1 / 01_2 / 01_3 / 02）的自动化实现。
目标机与账号配置沿用 `../Alarm/config_alarm.py`（当前为 AcuHMI-1-7 真机 192.168.3.71，
触发/数据设备 `AcuRev4100_392`）。

## 页面与实测事实

- 入口：Physical Devices → 设备详情 → Logs → **Data Log**（路由 `/3:1`）
- 控件：数据源 el-select（'Current'）+ 间隔 el-select（1 minute ~ 1 month 共 11 档）
  + daterange 区间（默认 昨天~今天）+ Download / Clear Logs
- **日期面板只允许选择有 datalog 数据的日期**，无数据日（含未来）全部 disabled
- 导出产物为 **gzip 压缩 CSV（\*.csv.gz）**，首列 TimeTag（ISO 时间戳，如
  `2026-07-16T15:29:00+0800`），下载接口 `/api/log/dataLog/export`

## 设计约束（关键，2026-07-17 与需求方确认）

**Log Interval 必须小于数据时长**：所选间隔档位大于设备现有数据跨度时，
后端拒绝导出（`/api/log/dataLog/export` 返回 500，前端无提示），这是设计
行为而非缺陷。因此各下载用例先用 1 minute 档探明当前数据跨度
（`current_data_span_seconds`），只对跨度足够的档位做下载校验，
不足的档位 skip 并注明待数据累计后重跑。

⚠️ 连带效应：case02（Clear Logs）每执行一次数据即归零，之后数小时内
大间隔档位（30 mins 以上逐级恢复）都会因跨度不足而 skip——这是预期行为。

## 执行与结果（2026-07-17 实跑）

```powershell
# case02 破坏性（清空该设备全部 datalog），务必放最后
python -m pytest projects/RPP/tests/Datalog -v
```

| 用例 | 结果 | 说明 |
|------|------|------|
| case01（1/5/10/15/30 mins） | ✅ | 下载解压校验：区间内时间戳 + 相邻间隔=所选档位（允许缺采样点，差值须为间隔整数倍）；数据跨度不足的档位 skip |
| case01_1（1h/6h/12h） | ✅ | 同上 |
| case01_2（1day/7days/1month） | ✅/skip | 同上；用例原文要求累计超 1 个月数据，跨度不足档位 skip 待补 |
| case01_3（部分区间无 datalog） | ✅ | 实装差异：面板禁选无数据日期，用例原文三种"无数据区间"无法构造；脚本改为验证禁选防护 + 有数据区间导出正确 |
| case02（Clear Logs） | ✅ | 清除后导出无数据行；**清空后各档位需重新累计数据** |

## 待跟进

1. 设备数据累计满 1 个月后复跑 case01 系列，所有档位可完整校验
2. 体验建议（非缺陷，可提需求）：跨度不足时前端目前无任何提示（接口 500 静默），
   可建议产品在 UI 层禁选不满足约束的档位或给出提示
