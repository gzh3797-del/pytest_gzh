# System Settings / Configuration Management — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/systemSettings/configurationManagement` |
| 路由名 | `configurationManagement` |
| 面包屑 | AcuHMI-1-7 / System Settings / Configuration Management |
| 顶级模块 | System Settings（...More） |

## 2. 页面用途

配置文件的导入/导出与恢复出厂设置。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| Import Configuration · Browse | file+button | `group('Import Configuration').getByRole('button',{name:'Browse'})` | 选择配置文件 |
| Import | button | `getByRole('button',{name:'Import'})` | 导入所选配置 |
| Export | button | `group('Export Configuration').getByRole('button',{name:'Export'})` | 导出当前配置 |
| Reset (Factory Reset) | button | `group('Factory Reset').getByRole('button',{name:'Reset'})` | 恢复出厂设置（危险操作，通常有二次确认） |

- 提示：`Caution: Importing configuration files between different software versions is not supported.`

## 4. 自动化测试要点

- Import：先 Browse 选文件再 Import；断言成功/失败提示；跨版本导入应报错（提示文案）。
- Export：触发文件下载。
- **Factory Reset：破坏性操作**，自动化中默认不实际执行，仅验证二次确认弹框出现与取消路径。

## 5. 机器可解析摘要

```json
{
  "route": "/systemSettings/configurationManagement",
  "name": "configurationManagement",
  "title": "Configuration Management",
  "module": "System Settings",
  "sections": {
    "Import Configuration": ["Browse(file)","Import"],
    "Export Configuration": ["Export"],
    "Factory Reset": ["Reset(destructive)"]
  },
  "caution": "cross-version import unsupported"
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/systemsettings/`。
> 入口：在 "...More" tooltip 弹出层内（先点 `menuitem "...More"` 再点子项，或直接 hash `#/systemSettings/configurationManagement`）。

### pytest 选择器与控件
- 三个 group：`Import Configuration`（Browse + Import）、`Export Configuration`（Export）、`Factory Reset`（Reset）。
- 页面提示：`Caution: Importing configuration files between different software versions is not supported.`

### 高危（实测确认）
- ⚠️ **配置导入和导出都会触发设备重启**（导入重启有用例记录；导出重启为 2026-07-03 口头确认）。**点击 Import / Export 一律视为重启类操作**：执行前须经用户确认，脚本需带重启等待+重登，禁无人值守连跑。
- ⚠️ **Factory Reset 破坏性**：默认不实际执行，仅验证二次确认弹框与取消路径。
