# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 构建命令

**不要用 `dotnet build`** — 本项目的 .NET 10 SDK 有 CET 兼容性问题（Windows 10 IoT LTSC 19044 不支持），且项目实际只需要 .NET Framework 4.8。

使用 VS 2026 自带的 MSBuild：

```bash
# Debug 完整编译
"C:/Program Files/Microsoft Visual Studio/18/Community/MSBuild/Current/Bin/MSBuild.exe" "e:/Code/qsbar/QSBar.sln" -nologo -m

# Release 完整编译
"C:/Program Files/Microsoft Visual Studio/18/Community/MSBuild/Current/Bin/MSBuild.exe" "e:/Code/qsbar/QSBar.sln" -nologo -m -p:Configuration=Release
```

仅编译 QSBar 主项目（跳过 Tests，因为 QSBar.Tests 是 SDK 格式会触发 CET 崩溃）：

```bash
"C:/Program Files/Microsoft Visual Studio/18/Community/MSBuild/Current/Bin/MSBuild.exe" "e:/Code/qsbar/QSBar/QSBar.csproj" -nologo -m -p:Configuration=Release
```

## 修改后工作流

> **编译前先 `touch` 改过的源文件，或者直接 `-t:Rebuild`。** 工具改文件后 `LastWriteTime` 不一定跟着更新，源文件时间戳一旦比 `obj\Release\QSBar.dll` 旧，MSBuild 就判定「输出比输入新」跳过 `CoreCompile`，照样打印「已成功生成 0 错误」，但 DLL 还是上一版。这种情况下用户测到的是旧代码，会把调试引到完全错误的方向。
>
> 每次编译后核对 `%LOCALAPPDATA%\QSBar\QSBar.dll` 的 `LastWriteTime` 是不是刚才那一刻；文件大小可能一模一样，不能拿大小当依据。要确认某处改动确实进了 DLL，按 UTF-16 搜字符串（.NET 字符串在 DLL 里是 UTF-16LE，`grep` 搜 ASCII 搜不到）：
>
> ```bash
> powershell -NoProfile -Command "$b=[System.IO.File]::ReadAllBytes('C:\Users\Kevin\AppData\Local\QSBar\QSBar.dll'); [System.Text.Encoding]::Unicode.GetString($b).Contains('要找的字符串')"
> ```

每次代码修改后，必须立即执行以下流程：

1. 用 VS 2026 MSBuild 编译 Release 版本（不能用 dotnet build，CET 不兼容）
2. 编译成功后自动运行 Register-QSBar.ps1 注册 COM 组件
3. 告知用户可以在 Excel 中测试

如果不注册 COM 组件，Excel 无法加载插件，用户看不到效果。

如果当前模式拦截 PowerShell 注册脚本，提醒用户切换权限模式。

注册命令：

```powershell
# 用户级注册（无需管理员，推荐日常使用）
powershell -ExecutionPolicy Bypass -File "e:\Code\qsbar\scripts\Register-QSBar.ps1" -DllPath "$env:LOCALAPPDATA\QSBar\QSBar.dll" -UserConfigOnly -RestartApps

# 管理员权限（完整注册）
powershell -ExecutionPolicy Bypass -File "e:\Code\qsbar\scripts\Register-QSBar.ps1" -DllPath "$env:LOCALAPPDATA\QSBar\QSBar.dll" -RestartApps
```

VS 中直接 F5 调试会自动触发 `AfterBuild` 完成编译→注册→启动 Excel。

清理环境并从头构建：`.\scripts\quick_setup.ps1`

## 架构概览

**类型**: COM 插件（非 VSTO），实现 `IDTExtensibility2` + `IRibbonExtensibility`，一套代码同时兼容 Excel 和 WPS。

| 项目 | 目标框架 | 格式 | 职责 |
|------|---------|------|------|
| `QSBar/` | .NET Framework 4.8 | 旧式 csproj | 插件主体：Ribbon UI、命令分发、COM 入口 |
| `QSBar.Core/` | .NET Framework 4.8 | 旧式 csproj | 核心业务逻辑（数据转换、格式策略、导出引擎） |
| `QSBar.Setup/` | .NET Framework 4.7.2 | 旧式 csproj | 安装器（WinForm GUI，含进度条和 COM 注册） |
| `QSBar.Tests/` | net48 | **SDK 格式** | xunit 测试（当前因 CET 问题无法用 dotnet build 编译） |

**入口点**: `WpsExcelAddIn.cs` — 唯一的 COM 入口类，CLSID `{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}`，ProgID `QSBar.WpsAddIn`。`OnConnection` 初始化全局 `App` 引用、启动定时器、注册快捷键、触发后台更新检查。

**Ribbon**: `Ribbon.xml`（嵌入资源）定义 Ribbon 布局，不再使用硬编码标签。所有 label/screentip/supertip 通过 `GetRibbonLabel`/`GetRibbonScreentip`/`GetRibbonSupertip` 回调获取，支持中英双语切换。语言偏好存储在 `HKCU\Software\QSBar\RibbonLanguage`。

**命令组织**: 功能按静态类分组，每个类对应 Ribbon 上的一个功能模块：
- `FormatCommands.cs` — 格式/排版（分级设置、样式、外链清理）
- `DataCommands.cs` — 数值化/文本化/公式锁定
- `ExportCommands.cs` — 导出当前表/标准报表/内部报表
- `SheetCommands.cs` — 合并工作表、生成表目录/文件目录、隐藏/空行清理
- `PhotoCommands.cs` — 图片选择与缩放
- `LegacyAppCommands.cs` — 计算模式切换、公式绝对引用锁定
- `UpdateManager.cs` — 后台更新检查（多镜像源切换）与重启更新流程

**COM 注册逻辑**: `Register-QSBar.ps1` 处理完整注册流程——RegAsm（32/64位）→ HKCU CLSID 手工注册（含 WOW6432Node）→ Excel/WPS Addins 注册表项 → WPS 白名单。用户级注册直接操作 HKCU 即可，无需管理员。

## 发布流程

`scripts/Publish.ps1` 一键发布：更新版本号 → 编译 Release → 生成安装包 → 更新 version.json → git commit & push。**不要手动修改版本号**。

## 注意事项

- **COM 引用不还，Excel 就退不掉进程**：`fullRange.Font.Name = x`、`ws.Cells[r,c].Value = v`、`foreach (Excel.Worksheet ws in wb.Worksheets)` 这类写法，每一次 `a.b.c` 的中间步都留下一个不会自己消失的引用。攒着不放，用户关掉窗口后 `EXCEL.EXE` 还在（`MainWindowHandle=0`、线程全在 Wait），下次启动就进安全模式。最毒的是 `ws.Application.ActiveWindow` —— 攥着 Application 引用等于告诉 Excel「还有客户端在用你」。
  - 写法：凡是接过 Excel 对象的局部变量，都在 `finally` 里走一遍 `ComUtil.Release`；集合遍历改成先用 `ComUtil.GetSheetNames` 取字符串、循环里再 `ComUtil.GetSheet` 拿了就还；`dynamic` 变量要写成 `ComUtil.Release((object)x)`，否则重载解析会跑到 `IEnumerable` 那个版本上
  - 定位手法：关掉 Excel 后 `Get-Process EXCEL`，进程还在且 `MainWindowHandle=0` 就是这个病。想确认是哪个功能，做减法实验——用一个功能、关掉、看进程，比读代码快得多。注意 32 位 Excel 要用 `SysWOW64` 下的 PowerShell 才枚举得到模块，而且托管程序集（QSBar.dll）在 `Process.Modules` 里根本看不到，别拿「没看到 QSBar」当作插件没加载的证据
  - 已排查过但**不是**原因、别再走回头路的方向：构建脚本的 `taskkill`、WinForms 的 `OpenFileDialog`（换成宿主 `GetOpenFilename` 照样残留）、`AutoUpgradeEnabled = false`、残留进程里的 `explorerframe.dll` / 百度网盘 `YunShellExtV1.dll`（Excel 打开过文件就会加载，与插件无关）
  - `sameFileForm.cs` / `otherFileForm.cs` 是没有任何入口的死代码，里面的泄漏没修
- **Excel 重开弹「安全模式」，也可能是构建脚本干的**：`AfterBuild` 会关掉 Excel/WPS 来释放 DLL 锁，用 `/F` 强杀会被 Excel 记成异常终止，下次启动就进恢复/安全模式。症状容易被误认成某个功能把 Excel 搞崩了，但 WER 里查不到 Excel 的崩溃报告、`HKCU\...\Excel\Resiliency\DisabledItems` 也是空的——强杀不留崩溃报告，Excel 也没把加载项判定成故障项。现在改成先温和 `taskkill`（走正常退出流程），三秒后才 `/F` 兜底
- 窗口级设置（`Window.SplitRow` / `FreezePanes` / `DisplayGridlines`）不能放在 `ExcelScope` 里做：屏幕刷新关着时改窗口的拆分/冻结会让 Excel 崩溃，而且目标表没激活时改的是别人那张表。要在作用域退出、`Activate()` 之后再设。另外已冻结的窗口不接受新的 `SplitRow`，得先 `FreezePanes = false`
- `NumberFormatLocal` 使用中文颜色标识 `[红色]`，不要改为 `[Red]`（中国区 Excel/WPS 不支持）。`[红色]` 必须写在格式段最前面，条件也要排在它后面：`_ * [红色]-#,##0.00_ ` 会被挪成 `[红色]_ * -#,##0.00_ `，`[<0][红色]` 会被挪成 `[红色][<0]`，回读与常量不符，直接导致轮换卡档
- 负号的处理在带条件和不带条件时相反：**条件段（如 `[<0]`）不自动加负号，必须手写 `-`**；无条件的「其余」段自动加，手写会变成 `--876.5万` 双负号。忘了手写就会显示成 `876.5万`，红色但没负号
- 会计格式（Ctrl+0）三档轮换：2位 → 0位 → 3位，首档给最常用的 2 位。百分号格式（Ctrl+8）单档 2 位。两者同一套视觉：红字负数、零值空白、`_ * ` 填充对齐。零值空白用 `_ * ""??_ ` 而非空段 `;;`，前者保留小数点对齐宽度
- 快捷键的权威列表在 `HelpForm.cs`（中英文两段）。改 `WpsExcelAddIn.RegisterShortcuts` 后必须同步这两处，否则帮助窗口会说谎。Ctrl+6 目前空闲
- 单位格式（Ctrl+9）是两档轮换：`"万"` → `"亿"`。量级词不带「元」，国际项目币种不定。每档单一量级、零值空白、红字负数。曾有的英文 `"M"`/`"k"` 档已取消：国际场景千分位本身就够读。想看原始数字按 Ctrl+0
- **格式段结构决定了能放什么**：不带条件时是标准的 `正;负;零;文本` 四段，零值段和颜色都放得下；一旦带条件就只有 `条件1;条件2;其余` 三段，没有独立的零值段和文本段，零值只能跟着某个量级段走。所以「零值空白」和「用条件做量级分档」不可兼得——单位格式拆成万、亿两个单档就是为了腾出零值段。Excel 最多 2 个条件，写第 3 个直接抛「不能设置 NumberFormatLocal」
- 两个轮换都不存状态，靠读当前 `NumberFormatLocal` 跟格式串常量逐字符比对反推档位，比不上就落回第一档。**改动任何格式串常量后必须实测 Excel 回读是否逐字符一致**（尾随空格、`_ * ` 都参与比对），否则轮换会永远卡在第一档
- **量级缩放公式**：格式串末尾每个逗号除以 10³，`!.` 从整数部分借 k 位移到字面小数点右边（等效再除以 10ᵏ，同时产生 k 位小数）。于是 `缩放 = 10^(3n+k)`，而**小数位数就是 k，不能自选**。万（10⁴）解 `3n+k=4` 只有 `n=1,k=1` 和 `n=0,k=4`，所以万档小数位只能是 1 或 4，**2 位小数无解**；亿（10⁸）解 `3n+k=8` 有 `n=2,k=2`，所以亿档恰好能有 2 位小数。1000 的整数次幂（千/百万/十亿）不需要 `!.`，小数位数自由且千分符可用
- 万/亿档不能加 `#,##` 千分位：千分位分组以格式串里的小数点为基准往左数三位，而 `!.` 就是那个小数点，于是分组错位——`#,##0!.0,"万"` 被重排成 `##,#0!.0,"万"`，显示 `-456,78.9万`。换成 `\.` 或 `"."` 一样被重排。同理 `#` 占位符也救不了零值——`!.` 是字面字符，`#` 吞不掉，零值会漏出 `.万`
- 单位格式「按量级自动转档」（如超 1 亿自动显示亿）已评估后不做：自动转档、零值空白、负数红字各要一个条件位，共 3 个，超出 Excel 的 2 个上限。而且负数无法跟着转档（要第 4 个条件位），会出现正数显示 `1.23亿`、负数显示 `-45678.9万` 的单位不一致。手动 Ctrl+9 切档反而整列统一
- `Release/` 目录不在 .gitignore 中，二进制安装包会提交到仓库。旧版安装包移到 `Release/旧中文版/` 归档
- 编译后的 DLL 会自动复制到 `%LOCALAPPDATA%\QSBar\`（AfterBuild target）
- 项目依赖 NuGet 包：EPPlus, Newtonsoft.Json, Microsoft.Office.Interop.Excel（本地 lib），都在 `packages/` 和 `lib/` 目录
