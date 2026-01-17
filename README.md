# BMToolkits (Clean Version)

这个文件夹包含了 BMToolkits 项目的所有核心组件，已精简掉所有冗余的参考文件和临时脚本。

## 目录结构
- `BMToolkits/`: 项目源代码和资源。
- `BMToolkits.sln`: 解决方案文件。
- `nuget.exe`: 用于修复依赖项。
- `quick_setup.ps1`: 一键式编译与注册脚本（支持 Excel 和 WPS）。

## 如何使用

### 1. 快速编译与注册
右键点击 `quick_setup.ps1` 选择 "使用 PowerShell 运行"。该脚本会自动：
- 恢复 NuGet 依赖（如 EPPlus）。
- 编译项目。
- 自动注册到 WPS 的 COM 列表和白名单。

### 2. 在 Excel 中使用
编译完成后，双击 `BMToolkits\bin\Debug\BMToolkits.vsto` 即可安装。

### 3. 在 WPS 中使用
运行完 `quick_setup.ps1` 后，直接打开 WPS 表格即可在 COM 加载项中看到 `BMToolkits (WPS)`。

## 注意事项
- 编译前请确保关闭所有 Excel 和 WPS 进程。
- 脚本需要管理员权限来运行 `RegAsm`。
