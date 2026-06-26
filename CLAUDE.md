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

- `NumberFormatLocal` 使用中文颜色标识 `[红色]`，不要改为 `[Red]`（中国区 Excel/WPS 不支持）
- `亿/万元` 数字格式用 `"亿元"` / `"万元"` 而非 `"B"` / `"M"`（工程造价行业约定）
- `Release/` 目录不在 .gitignore 中，二进制安装包会提交到仓库。旧版安装包移到 `Release/旧中文版/` 归档
- 编译后的 DLL 会自动复制到 `%LOCALAPPDATA%\QSBar\`（AfterBuild target）
- 项目依赖 NuGet 包：EPPlus, Newtonsoft.Json, Microsoft.Office.Interop.Excel（本地 lib），都在 `packages/` 和 `lib/` 目录
