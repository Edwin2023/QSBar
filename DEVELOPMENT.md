# QSBar (COM Version) 开发指南

![Excel Bar](Installinfo/SHOW1_EXCEL_BAR.bmp)
![WPS Bar](Installinfo/SHOW2_WPS_BAR.bmp)

本项目是 QSBar 插件的 COM 版本，旨在提供一套代码同时兼容 **Microsoft Excel** (包括 Excel 2021) 和 **WPS 表格** 的统一解决方案。

## 目录结构
- `QSBar/`: 插件核心源代码（UI、逻辑处理）。
- `QSBar.Core/`: 核心业务逻辑（数据处理、格式转换、导出策略等）。
- `QSBar.sln`: Visual Studio 解决方案文件。
- `Installinfo/`: **打包资源目录**（包含安装脚本、版本信息、更新日志及展示图片）。
- `scripts/Publish.ps1`: **发布脚本**（一键编译 Release、更新版本号、同步 Installinfo 信息、推送 Gitee）。
- `scripts/quick_setup.ps1`: **开发调试脚本**（一键重置环境、编译 Debug、注册插件）。
- `scripts/Register-QSBar.ps1`: **核心注册脚本**（底层工具，处理注册表清理与写入）。
- `0memory/`: 项目开发记忆与技术文档。

## 快速开始

### 1. 开发环境配置
- **Visual Studio 2022**: 推荐安装 .NET Desktop Development 工作负载。
- **.NET Framework 4.8**: 项目运行的基础环境。

### 2. 开发与调试 (Development)

#### 方式 A：Visual Studio 直接调试 (推荐)
1.  **修改代码**：在 VS 中编辑。
2.  **启动调试**：直接按 **F5**。
    *   VS 会自动编译 Debug 版本。
    *   触发 `AfterBuild` 事件自动调用注册脚本。
    *   自动启动 Excel 并附加调试器。

#### 方式 B：使用脚本重置环境
如果遇到插件不加载、F5 报错或需要彻底清理环境，请运行：
```powershell
.\scripts\quick_setup.ps1
```
此脚本会：
1.  强制关闭所有 Excel/WPS 进程。
2.  清理所有旧的注册表项（包括禁用项）。
3.  重新编译 Debug 版本并注册。

### 3. 正式发布 (Production)

本项目使用 `Publish.ps1` 脚本进行一键发布。**请勿手动修改 AssemblyInfo.cs 中的版本号**，脚本会自动处理。

#### 发布步骤：
1.  打开 PowerShell (建议在 VS 的终端中)。
2.  运行发布命令，指定**新版本号**和**更新日志**：

```powershell
.\scripts\Publish.ps1 -Version "1.0.0.5" -Log "修复了版本号显示问题，优化了注册逻辑"
```

#### 脚本执行流程：
1.  **版本更新**：自动修改 `AssemblyInfo.cs` 和 `QSBar.csproj` 为新版本号。
2.  **同步 Installinfo**：自动更新 `Installinfo/version.json` 和 `Installinfo/QSBar_Installer.iss` 中的版本号。
3.  **编译 Release**：调用 MSBuild 重新编译 Release 版本。
4.  **构建发布包**：将 DLL 复制到 `Release/` 目录。
5.  **更新元数据**：更新根目录 `version.json` 供客户端检查更新。
6.  **推送代码**：自动提交 git commit 并推送到 Gitee `master` 分支。

## 常见问题
- **插件未显示**：通常是因为 Excel 将插件加入了“禁用项”。运行 `scripts/quick_setup.ps1` 可自动修复。
- **版本号未更新**：请确保使用 `Publish.ps1` 发布，它会处理 AssemblyInfo 的编码和版本写入。
- **权限问题**：脚本默认注册到 HKCU (当前用户)，无需管理员权限。如果遇到 HKLM 冲突，脚本会提示你使用管理员权限运行 `-CleanHKLM` 参数。

---
*所属项目：QSBar*