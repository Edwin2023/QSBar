# Office Ribbon 图标 (imageMso) 兼容性与避坑指南

在开发 Excel/WPS 插件时，`imageMso`（内置图标）的兼容性是一个臭名昭著的“深坑”。不同版本的 Office 甚至是 WPS 对图标 ID 的支持各不相同，一旦引用了不存在的 ID，会导致整个自定义 UI 加载失败并弹出极其不友好的错误。

## 1. 为什么会报错？
- **版本差异**：微软在 Office 2013、2016、2019/365 中不断增加新的图标。如果你在代码中使用了 2019 才有的图标，而用户使用的是 2010 或 WPS，就会报错。
- **WPS 限制**：WPS 虽然兼容 Office 的 Ribbon XML 语法，但它内置的图标库版本较旧（通常接近 Office 2007/2010），许多带“Warning”、“Sync”、“Status”字样的现代图标都不支持。
- **严苛检查**：Office/WPS 会在解析 XML 时校验 `imageMso` 是否存在，不存在即抛出 `0x80004005` 运行时错误。

## 2. 避坑准则
### A. 优先使用“黄金兼容”图标
以下图标自 Office 2007 以来从未改变，且在 WPS 中也 100% 存在：
- **基础操作**：`FileSave` (保存), `FileSaveAs` (另存为), `Help` (帮助), `FileOpen` (打开), `Refresh` (刷新), `Delete` (删除)。
- **警告/提示**：
    - `DialogExclamation` (最标准的**黄色三角感叹号**，推荐用于更新提示)
    - `HighImportance` (红色感叹号)
    - `Note` (黄色便签本)
    - `Information` (蓝色 i 字母)
- **其他安全图标**：`HappyFace`, `Heart`, `CalculateNow`, `PasteValues`, `ViewSheetGridlines`。

### B. 绝对不要使用的图标
- 带有 `Sync` 前缀的（如 `SyncStatusWarning`）
- 带有版本号或现代风格的（如 `AccessibilityChecker`）
- 过于具体的 Office 365 专有图标

## 3. 终极解决方案
如果你对图标美观度有极高要求，或者 `imageMso` 实在无法满足：
1. **使用 getImage 回调**：在 XML 中使用 `getImage='GetMyImage'`，然后在 C# 代码中动态返回一个 `System.Drawing.Bitmap`。
2. **嵌入资源**：将 `.png` 图片作为资源嵌入到 DLL 中。这样图标就是你自己的，永远不会因为环境不同而消失或报错。

## 4. 如何查找可用图标？
- **官方画廊**：搜索下载 "Office 2010 Add-In: Icons Gallery"。
- **在线查看器**：[BERT | ImageMSO List Reference](https://bert-toolkit.com/imagemso-list.html)
- **实时预览**：在 Excel 的“自定义快速访问工具栏”设置中，将鼠标悬停在任何命令上，括号里显示的（如 `FileSave`）就是它的 `imageMso` 名称。

---
*最后更新：2026-01-18*
*存放位置：`e:\Code\QSBar\0memory\Office_Ribbon图标指南.md`*
