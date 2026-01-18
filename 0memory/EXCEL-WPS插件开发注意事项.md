# EXCEL-WPS 插件开发与加载注意事项

本指南总结了 QSBar 项目在适配 Excel 2021 和 WPS 过程中积累的核心技术要点，确保插件在双平台下均能稳定加载。

## 1. 核心技术架构：统一使用 COM
本项目放弃了 VSTO 模式，统一采用 **纯 COM 插件** 架构。
- **接口要求**：必须实现 `IDTExtensibility2` 和 `IRibbonExtensibility`。
- **注册方式**：使用 `RegAsm.exe` 注册，并手动配置注册表以兼容 WPS。
- **优点**：彻底解决 Excel 与 WPS 之间的插件冲突，实现一套 DLL 双平台运行。

## 2. Excel 2021 图标与 XML 兼容性（重要）
Excel 2021 的 Ribbon 引擎比以往版本（及 WPS）更加严苛，任何不规范的 XML 都会导致加载失败。
- **严禁使用模糊 ID**：例如 `Group`、`CalculationOptions` 等在 2021 中被识别为容器或保留字。
- **推荐替换方案**：
  - `Group` -> `ObjectsGroup`
  - `Ungroup` -> `OutlineUngroup`
  - `CalculationOptions` -> `CalculateNow`
  - `Currency` -> `CommaStyle`
  - `NumberFormatDialog` -> `NumberFormatCells`
- **属性限制**：
  - `size="small"`：在 Office 2009 Schema 中是非法的，应使用 `size="normal"` 或不写。
  - `supertip` 放置位置：在 `splitButton` 元素上直接使用 `supertip` 属性会导致 `0x80004005` 错误。**修复方法**：将 `screentip` 和 `supertip` 移动到 `splitButton` 下的第一个默认 `button` 元素上。

## 3. 开发效率优化：一键调试 (F5 流程)
为了模拟 VBA “即改即见效”的体验，项目已在 `.csproj` 中配置了深度集成：
- **自动启动配置**：在 `QSBar.csproj` 的 `Debug` 节点配置了 `<StartAction>Program</StartAction>` 和 `<StartProgram>`。点击 **F5** 会自动启动 Excel 并挂载调试器。
- **自动注册联动**：配置 `AfterBuild` 目标。每次生成成功后，会自动触发 `scripts/dev_register.ps1`。
- **功能**：自动执行 32/64 位 RegAsm 注册、写入 HKCU 注册表、配置 WPS 白名单。

## 4. 开发常见坑位与解决方法 (Troubleshooting)
### 4.1 文件锁定导致生成失败
- **现象**：报错 `MSB3021/MSB3027`，提示 `bin\Debug\QSBar.dll` 正在被另一进程使用。
- **原因**：Excel/WPS 窗口未关闭，或者后台残留了调试进程（如 PID 锁定的 PowerShell）。
- **解决**：必须先**关闭所有 Excel/WPS 窗口**再点击生成或 F5。如果依然报错，需在任务管理器中强制结束 `EXCEL.EXE` 或对应的 `powershell.exe`。

### 4.2 脚本编码导致的语法错误
- **现象**：报错 `字符串缺少终止符` 或 `TerminatorExpectedAtEndOfString`。
- **原因**：PowerShell 5.1 对非 UTF-8 with BOM 编码的中文字符支持较差，乱码可能导致引号解析失败。
- **解决**：`scripts/*.ps1` 脚本中的 `Write-Host` 尽量使用**纯英文**，或确保文件保存为 `UTF-8 with BOM` 编码。

### 4.3 颜色偏差与 TintAndShade
- **现象**：设置了固定 RGB，但在 Excel 中显示颜色偏亮或偏暗（例如深蓝变亮蓝）。
- **原因**：Excel 单元格残留了 `TintAndShade`（亮度偏移）属性。
- **解决**：在设置 `Color` 的同时，必须显式设置 `Interior.TintAndShade = 0` 和 `Font.TintAndShade = 0`。

## 5. 自动化打包与部署 (Inno Setup)
为了实现一键安装，本项目使用了 Inno Setup。
- **递归调用陷阱**：在 `[Code]` 段中编写 Pascal 脚本时，**严禁**定义与内置函数完全同名的自定义函数。例如，自定义 `function Is64BitInstallMode` 会导致安装程序在“Saving uninstall information”阶段进入死循环。
- **RegAsm 注册**：
  - 必须同时运行 32 位和 64 位注册（`{dotnet40}\RegAsm.exe` 和 `{dotnet4064}\RegAsm.exe`）。
  - 使用 `/codebase` 参数以支持未签名的程序集。

## 5. 解决加载失败的“韧性”清理
如果插件不显示，请检查以下注册表项：
- **禁用项**：`HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems` (WPS 也有类似路径)。
- **白名单 (WPS)**：ProgID 必须存在于 `HKCU\Software\Kingsoft\Office\ET\AddinsWL` 中。
- **LoadBehavior**：确保值为 `3`。

## 6. 环境要求
- **管理员权限**：VS 生成和安装包运行均需管理员权限。
- **日志调试**：设置系统环境变量 `VSTO_LOGALERTS = 1`，Excel 加载失败时会弹出具体的 XML 错误提示框，这是调试 Ribbon XML 的最快手段。

---
*最后更新日期：2026-01-18*
*所属项目：QSBar*
