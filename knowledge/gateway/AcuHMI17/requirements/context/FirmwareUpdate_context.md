# Firmware Update — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/firmwareUpdate` |
| 路由名 | `firmwareUpdate` |
| 面包屑 | AcuHMI-1-7 / Firmware Update |
| 顶级模块 | Firmware Update（顶级，无子页） |

## 2. 页面用途

展示当前固件版本并手动上传固件包升级。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 示例 | 说明 |
|------|------|-----------|------|------|
| Current Firmware Version | 只读文本 | 文本 "Current Firmware Version:" | v1.03p05 | 只读 |
| Firmware Update File · Browse | file+button | `group('Firmware Update File').getByRole('button',{name:'Browse'})` | Choose file | 选择固件文件（必填*） |
| Upload | button | `getByRole('button',{name:'Upload'})` | — | 上传并升级 |

## 4. 自动化测试要点

- 断言当前版本文本；Browse 选文件 → Upload。
- **升级为破坏性/长耗时操作**：自动化中默认不实际执行；验证未选文件时 Upload 的校验提示。

## 5. 机器可解析摘要

```json
{
  "route": "/firmwareUpdate",
  "name": "firmwareUpdate",
  "title": "Firmware Update",
  "module": "Firmware Update",
  "readonly": ["Current Firmware Version"],
  "elements": ["Firmware Update File(Browse, required)","Upload"]
}
```
