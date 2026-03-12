# QSBar (COM Version)

![Excel Bar](SHOW1_EXCEL%20BAR.png)
![WPS Bar](SHOW2_WPS_BAR.png)

本项目是 QSBar 插件的 COM 版本，旨在提供一套代码同时兼容 **Microsoft Excel** (包括 Excel 2021) 和 **WPS 表格** 的统一解决方案。

## 目录结构
- `QSBar/`: 插件核心源代码（UI、逻辑处理）。
- `QSBar.Core/`: 核心业务逻辑（数据处理、格式转换、导出策略等）。
- `QSBar.sln`: Visual Studio 解决方案文件。
- `QSBar_Installer.iss`: Inno Setup 打包脚本，用于生成一键安装 EXE。
- `scripts/Register-QSBar.ps1`: 核心注册脚本（统一支持开发注册、用户安装、卸载）。
- `scripts/Register.bat`: 一键注册脚本（用户模式）。
- `scripts/Unregister.bat`: 一键卸载脚本。
- `0memory/`: 项目开发记忆与技术文档（包含 VBA 源码参考）。

## 快速开始

### 1. 开发环境配置
- **Visual Studio 2022**: 建议以**管理员身份**运行（非必须，但推荐）。
- **Inno Setup**: 用于生成最终的安装程序。
- **.NET Framework 4.8**: 项目运行的基础环境。

### 2. 开发阶段的调试（极简流程）
为了模拟 VBA “即改即见效”的体验，项目已配置全自动开发流：
1. **修改代码**：在 Visual Studio 中进行逻辑或 UI 修改。
2. **一键调试**：直接按 **F5**。
   - **自动化操作**：
     - 自动检测并强制关闭残留的 Excel/WPS 进程（解决文件占用问题）。
     - 编译并部署最新 DLL 到 `%LOCALAPPDATA%\QSBar`。
     - 调用 `scripts/Register-QSBar.ps1` 更新注册信息。
     - **自动启动 Excel** 并挂载调试器。
   - **效果**：Excel 启动后即可直接测试新功能。
3. **循环开发**：测试完后，可以直接停止调试（Shift+F5），VS 会自动处理进程清理（如下次启动时）。

> **常见问题排查**：
> - **生成失败**：虽然已配置自动杀进程，但如果文件被非 Excel 进程占用，仍可能失败。
> - **脚本错误**：如果注册脚本报错，请检查 PowerShell 执行策略 (`Set-ExecutionPolicy RemoteSigned`)。

### 3. 发布安装包
1. 在 Visual Studio 中切换到 `Release` 模式并生成。
2. 使用 Inno Setup 打开 `QSBar_Installer.iss`。
3. 点击 `Compile`，生成的安装包将存放在 `Installer/QSBar_Setup.exe`。

## 注意事项
- **图标兼容性**: Excel 2021 对 `imageMso` 校验非常严格，已在 `0memory/` 文档中详细记录。
- **管理员权限**: 注册 COM 组件需要写入系统注册表，VS 生成时如果报错，请检查是否以管理员身份运行。
- **进程占用**: 如果生成失败，请确保关闭所有 `EXCEL.EXE`、`WPS.EXE` 和 `ET.EXE`。

---
*所属项目：QSBar*
