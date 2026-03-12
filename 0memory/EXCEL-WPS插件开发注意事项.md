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
- **自动注册联动**：配置 `AfterBuild` 目标。每次生成成功后，会自动触发 `scripts/Register-QSBar.ps1`。
- **功能**：自动执行 RegAsm 注册、写入 HKCU 注册表、配置 WPS 白名单。

## 4. 开发常见坑位与解决方法 (Troubleshooting)
### 4.1 文件锁定导致生成失败
- **现象**：报错 `MSB3021/MSB3027`，提示 `bin\Debug\QSBar.dll` 正在被另一进程使用。
- **原因**：Excel/WPS 窗口未关闭。
- **解决**：项目 `QSBar.csproj` 已配置 `AfterBuild` 自动执行 `taskkill` 关闭 Excel/WPS 进程。如果依然报错，请确保没有以管理员权限运行的僵尸进程。

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
- **RegAsm 注册的坑与终极方案**：
  - **无法定位程序集 (RA0000)**：这是由于 `RegAsm` 找不到 DLL 文件。在 Inno Setup 的 `[Run]` 段中，必须显式设置 `WorkingDir: "{app}"`，确保执行上下文在插件目录。
  - **注册表冗余冲突**：严禁在 `[Registry]` 中手动写入 `CLSID`、`InprocServer32` 等核心 COM 项。这些项应由 `RegAsm /codebase` 动态生成。手动写入如果路径格式（如斜杠方向）有误，会直接阻塞 WPS 的加载。
  - **参数格式匹配**：建议使用 `/codebase /tlb`。不要在 `/tlb` 后接具体路径，让其在当前目录下自动生成，兼容性最强。
  - **静默运行卡死**：如果使用 `.bat` 辅助注册，严禁包含 `pause` 命令。在 `Flags: runhidden` 模式下，`pause` 会导致安装程序永久卡死在后台。
  - **终极必杀技：外部 BAT 脚本替代 ISS 指令**：
    - **现状**：直接在 ISS 的 `[Run]` 中写 `Filename: "{dotnet40}\RegAsm.exe"` 虽然看起来正规，但由于 Inno Setup 宏解析、权限隔离以及 working directory 的细微差异，经常出现“手动执行好使，安装包执行无效”的玄学问题。
    - **对策**：将注册逻辑封装进一个独立的 `Register_QSBar.bat`。在 ISS 中只负责把这个 `bat` 拷贝到安装目录并调用它。
    - **优点**：`.bat` 文件在磁盘上真实存在，环境上下文与手动双击完全一致。如果安装失败，还可以直接去安装目录双击 `bat` 看到报错信息，极大降低了调试成本。
- **环境要求**：必须同时运行 32 位和 64 位注册（`{dotnet40}\RegAsm.exe` 和 `{dotnet4064}\RegAsm.exe`）。

## 6. 解决加载失败的“韧性”清理
如果插件不显示，请检查以下注册表项：
- **禁用项**：`HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems` (WPS 也有类似路径)。
- **白名单 (WPS)**：ProgID 必须存在于 `HKCU\Software\Kingsoft\Office\ET\AddinsWL` 中。
- **LoadBehavior**：确保值为 `3`。

## 7. 环境要求
- **管理员权限**：VS 生成和安装包运行均需管理员权限。
- **日志调试**：设置系统环境变量 `VSTO_LOGALERTS = 1`，Excel 加载失败时会弹出具体的 XML 错误提示框，这是调试 Ribbon XML 的最快手段。

## 8. 快捷键覆盖与拦截方案 (Shortcuts Solution)
在 C# 插件开发中，快捷键拦截（尤其是覆盖 Excel 原生 `Ctrl+9`）是公认的难点。本项目经历了一个从“盲目探测”到“精准破局”的过程。

### 7.1 破局点：Application.OnKey 的“侦察”作用
**这是整个快捷键方案成功的关键。** 
在初期尝试各种钩子均无反应时，通过 `Application.OnKey("^9", "C#方法名")` 强制触发了一个 Excel 报错（提示找不到宏）。
- **核心价值**：这个报错虽然看起来是失败，但它第一次**证实了 Excel 确实可以截获按键并尝试交回给插件处理**。
- **结论**：它排除了“按键未到达”的疑虑，将问题从“如何捕获”缩小到了“如何更优雅地执行”，为后续转向 `WH_KEYBOARD` 本地钩子提供了坚实的理论依据。

### 7.2 方案进化史
| 方案 | 评价 | 关键教训 |
| :--- | :--- | :--- |
| **Timer 轮询** | 失败 | 响应慢，且无法阻止 Excel 原生功能的触发。 |
| **Application.OnKey** | **破局点** | 成功捕获按键并触发逻辑，但因 COM 协议层对“宏名”的解析限制，会导致弹窗报错。 |
| **全局钩子 (LL)** | 失败 | 调试模式下会导致 Excel 界面严重卡死。 |
| **本地钩子 (Hook)** | **最终方案** | 继承了 OnKey 的捕获成功率，同时解决了报错和卡死问题。 |

### 7.3 本地钩子 (WH_KEYBOARD) 实现细节
1. **精准挂载**：通过 `GetCurrentThreadId()` 将钩子仅挂载在 Excel 主线程，避免全局污染。
2. **强制消费**：匹配到快捷键后立即执行业务逻辑并 `return (IntPtr)1`，从消息队列中彻底移除该按键，防止 Excel 执行原生动作。
3. **同步执行**：由于是在主线程钩子内，C# 代码直接拥有 Excel 的 UI 线程访问权，执行效率最高且最稳定。

---
*最后更新日期：2026-01-21*
*所属项目：QSBar*
