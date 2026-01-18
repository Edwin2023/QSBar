# EXCEL-WPS 插件开发与加载注意事项

本指南总结了 QSBar 项目在适配 Excel 2021 和 WPS 过程中积累的核心技术要点，确保插件在双平台下均能稳定加载。

## 1. 核心技术架构：统一使用 COM
本项目放弃了 VSTO 模式，统一采用 **纯 COM 插件** 架构。
- **接口要求**：必须实现 `IDTExtensibility2` 和 `IRibbonExtensibility`。
- **注册方式**：使用 `RegAsm.exe` 注册，并手动配置注册表以兼容 WPS。
- **优点**：彻底解决 Excel 与 WPS 之间的插件冲突，实现一套 DLL 双平台运行。

## 2. Excel 2021 图标兼容性（重要）
Excel 2021 的 Ribbon 引擎比以往版本（及 WPS）更加严苛。
- **严禁使用模糊 ID**：例如 `Group`、`CalculationOptions` 等在 2021 中可能被识别为容器或保留字，导致 Ribbon 加载失败（错误码 `0x80004005`）。
- **推荐替换方案**：
  - `Group` -> `ObjectsGroup`
  - `Ungroup` -> `ObjectsUngroup`
  - `CalculationOptions` -> `CalculateNow`
  - `Currency` -> `AccountingNumberFormat` 或 `CommaStyle`
- **校验机制**：Excel 2021 会在加载时实时校验 XML 中的每一个 `imageMso`。若有一个无效，整个 Tab 可能会消失或报错。

## 3. Ribbon XML Schema 限制
- **supertip 放置位置**：在 `splitButton` 元素上直接使用 `supertip` 属性在某些 Schema 下是非法的。
- **修复方法**：将 `screentip` 和 `supertip` 移动到 `splitButton` 下的第一个默认 `button` 元素上。

## 4. 自动化打包与部署 (Inno Setup)
为了实现一键安装，本项目使用了 Inno Setup。
- **递归调用陷阱**：在 `[Code]` 段中编写 Pascal 脚本时，避免定义与内置函数同名的函数（如 `Is64BitInstallMode`），否则会导致安装程序在“保存卸载信息”阶段死循环卡死。
- **RegAsm 注册**：
  - 必须同时运行 32 位和 64 位注册（`dotnet40` 和 `dotnet4064`）。
  - 使用 `/codebase` 参数以支持未签名的程序集。

## 5. 解决加载失败的“韧性”清理
如果插件不显示，请检查以下注册表项并清理：
- **禁用项**：`HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems`
- **白名单 (WPS)**：必须确保 ProgID 存在于 `HKCU\Software\Kingsoft\Office\ET\AddinsWL` 中。
- **LoadBehavior**：确保值为 `3`。

## 6. 开发环境建议
- **VS 提权**：始终以管理员身份运行 Visual Studio，以便在生成时能够成功写入 `RegisterForComInterop` 相关的注册表项。
- **日志调试**：设置系统环境变量 `VSTO_LOGALERTS = 1`，Excel 加载失败时会弹出具体的 XML 错误提示。

---
*最后更新日期：2026-01-18*
*所属项目：QSBar (COM Version)*
