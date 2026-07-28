# Templates / Import — 页面上下文

## 1. 页面标识

| 项 | 值 |
|----|----|
| 路由 | `/#/templates/import` |
| 路由名 | `import` |
| 面包屑 | AcuHMI-1-7 / Templates / Import |
| 顶级模块 | Templates |

## 2. 页面用途

导入设备模板文件。

## 3. 交互元素清单

| 元素 | 类型 | 定位策略 1 | 说明 |
|------|------|-----------|------|
| 二级 tab | menuitem | `getByRole('menuitem',{name:'Template List'})` 等 | 切换子页 |
| Template File · Browse | file+button | `group('Template File').getByRole('button',{name:'Browse'})` | 选择模板文件（显示 "Choose file"） |
| Upload | button | `getByRole('button',{name:'Upload'})` | 上传导入 |

## 4. 自动化测试要点

- Browse 选文件 → Upload；断言成功/失败提示；导入后模板出现在 Template List 的 Customized 表。
- 非法文件/空文件的错误处理。

## 5. 机器可解析摘要

```json
{
  "route": "/templates/import",
  "name": "import",
  "title": "Import",
  "module": "Templates",
  "elements": ["Template File(Browse)","Upload"]
}
```

## 实测测试情报（pytest / Element Plus，来源：2026-07-03 联机实测）

> 对应测试目录：`projects/AcuHMI_1_7/tests/ui/templates/`。

### 进入路径
- 顶部子菜单 `.el-menu-item` `.filter(has_text="Import")` 点击；或直达 `#/templates/import`。

### 结果反馈
- Upload 后导入为异步：模板出现在 Template List 的 **Customized 表**（末页/需滚动），**轮询 `page.reload()`** 确认（同 Template List 的异步规律）。
