# Office Ribbon 图标 (imageMso) 避坑

`imageMso` ID 引用错误会导致整个 UI 加载失败 (0x80004005)。

## 1. 黄金兼容图标 (WPS & Office 2007+ 均安全)
- **通用**：`FileSave`, `FileSaveAs`, `Help`, `FileOpen`, `Refresh`, `Delete`
- **提示**：`DialogExclamation` (黄三角感叹号), `HighImportance` (红感叹号), `Information` (蓝色 i)
- **其他**：`HappyFace`, `Heart`, `CalculateNow`, `PasteValues`

## 2. 绝对禁区 (必报报错)
- 带有 `Sync` 前缀的 (如 `SyncStatusWarning`)
- 带有版本号或现代风格的 (如 `AccessibilityChecker`)
- WPS 不支持的 Office 2016+ 新图标

## 3. 终极方案：getImage
如果不想赌兼容性：
1. XML 中使用 `getImage="GetMyImage"`。
2. C# 中动态返回嵌入资源的 `.png` (System.Drawing.Bitmap)。
3. 这样图标永久打包在 DLL 中，环境无关。

## 4. 快速查找
- 在 Excel “自定义快速访问工具栏”设置中，悬停在命令上，括号里的名称（如 `FileSave`）即为 `imageMso`。
- 在线参考：[BERT imageMso List](https://bert-toolkit.com/imagemso-list.html)
