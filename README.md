# QSBar (COM Version)

本项目是 QSBar 插件的 COM 版本，旨在提供一套代码同时兼容 **Microsoft Excel** (包括 Excel 2021) 和 **WPS 表格** 的统一解决方案。

## 目录结构
- `QSBar/`: 插件核心源代码。
- `QSBar.sln`: Visual Studio 解决方案文件。
- `QSBar_Installer.iss`: Inno Setup 打包脚本，用于生成一键安装 EXE。
- `quick_setup.ps1`: 开发阶段的一键式编译与注册脚本（支持管理员权限自动提权）。
- `0memory/`: 项目开发记忆与技术文档。

## 快速开始

### 1. 开发环境配置
- **Visual Studio 2022**: 建议以管理员身份运行。
- **Inno Setup**: 用于生成最终的安装程序。
- **.NET Framework 4.8**: 项目运行的基础环境。

### 2. 开发阶段的调试
右键点击 `quick_setup.ps1` 选择 "使用 PowerShell 运行"。该脚本会自动：
- 恢复 NuGet 依赖。
- 使用 MSBuild 编译项目。
- 自动完成 COM 注册并配置 WPS 白名单。
- 运行自检测试，验证 COM 对象是否创建成功。

### 3. 发布安装包
1. 确保在 Visual Studio 中以 `Release` 模式生成项目。
2. 使用 Inno Setup 打开 `QSBar_Installer.iss`。
3. 点击 `Compile`，生成的安装包将存放在 `Installer/QSBar_Setup.exe`。

## 注意事项
- **图标兼容性**: Excel 2021 对 `imageMso` 校验非常严格，请务必使用官方推荐的图标 ID（如 `ObjectsGroup` 而非 `Group`）。
- **管理员权限**: 注册 COM 组件需要写入系统注册表，开发调试和安装过程均需管理员权限。
- **进程占用**: 编译或安装前，请确保关闭所有 Excel 和 WPS 进程。
