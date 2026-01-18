# QSBar (COM Version)

本项目是 QSBar 插件的 COM 版本，旨在提供一套代码同时兼容 **Microsoft Excel** (包括 Excel 2021) 和 **WPS 表格** 的统一解决方案。

## 目录结构
- `QSBar/`: 插件核心源代码（UI、逻辑处理）。
- `QSBar.Core/`: 核心业务逻辑（数据处理、格式转换、导出策略等）。
- `QSBar.sln`: Visual Studio 解决方案文件。
- `QSBar_Installer.iss`: Inno Setup 打包脚本，用于生成一键安装 EXE。
- `scripts/dev_register.ps1`: 由 VS 自动调用的开发注册脚本（支持 Excel & WPS）。
- `quick_setup.ps1`: 备用的全自动环境初始化脚本。
- `0memory/`: 项目开发记忆与技术文档（包含 VBA 源码参考）。

## 快速开始

### 1. 开发环境配置
- **Visual Studio 2022**: 建议以**管理员身份**运行。
- **Inno Setup**: 用于生成最终的安装程序。
- **.NET Framework 4.8**: 项目运行的基础环境。

### 2. 开发阶段的调试（极简流程）
为了模拟 VBA “即改即见效”的体验，项目已配置全自动开发流：
1. **修改代码**：在 Visual Studio 中进行逻辑或 UI 修改。
2. **一键调试**：直接按 **F5**。
   - **自动化操作**：VS 会自动编译 -> 调用 `scripts/dev_register.ps1` 注册 COM -> **自动启动 Excel** 并挂载调试器。
   - **效果**：Excel 启动后即可直接测试新功能。
3. **循环开发**：测试完后，**关闭 Excel**，再次修改代码并按 **F5** 即可。

> **常见问题排查**：
> - **生成失败**：通常是因为 Excel 窗口没关，导致 `QSBar.dll` 被占用。请关闭所有 Excel/WPS 进程。
> - **脚本错误**：如果注册脚本报错，请确保 `scripts/dev_register.ps1` 中不包含导致编码问题的中文字符。

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
