# EXCEL-WPS 插件开发与加载注意事项

本指南总结了 QSBar 项目在适配 Excel 和 WPS 过程中积累的核心技术要点，旨在确保插件在双平台下均能稳定加载。

## 1. 核心技术架构：统一使用 COM
为了避免 Excel VSTO 与 WPS COM 注册之间的冲突，本项目统一采用 **纯 COM 插件** 架构。
- **接口要求**：必须实现 `IDTExtensibility2`（连接管理）和 `IRibbonExtensibility`（菜单自定义）。
- **注册方式**：使用 `RegAsm.exe` 进行 COM 注册，而不是使用 VSTO 的 `.vsto` 清单模式。
- **优点**：一套 DLL 兼容双平台，消除了 VSTO 带来的“已安装另一版本”冲突问题。

## 2. 注册表关键路径
插件必须在注册表的特定位置声明才能被 Office/WPS 识别。

### 2.1 Excel 加载路径
- `HKCU:\Software\Microsoft\Office\Excel\Addins\[ProgID]`
- **关键键值**：
  - `LoadBehavior` (DWORD): 必须设为 `3`（启动时加载）。
  - `FriendlyName` & `Description`: 显示在“COM 加载项”列表中的名称。

### 2.2 WPS 加载路径
- `HKCU:\Software\Kingsoft\Office\ET\Addins\[ProgID]`
- `HKCU:\Software\Kingsoft\Office\ET\AddinsData\[ProgID]`
- `HKCU:\Software\Kingsoft\Office\WPS\Addins\[ProgID]`

### 2.3 WPS 白名单（至关重要）
WPS 对第三方插件有严格限制，必须将 ProgID 加入白名单：
- 路径：`HKCU:\Software\Kingsoft\Office\[ET/WPS/Common/6.0]\AddinsWL`
- 操作：在该路径下新建一个以插件 **ProgID** 为名称的字符串值（值可为空）。

## 3. 解决 Excel 加载失败（韧性处理）
Excel 有时会自动禁用崩溃或加载缓慢的插件，需通过脚本自动修复：
- **清理禁用列表**：删除 `HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems` 下的所有项。
- **强制启用**：在 `HKCU:\Software\Microsoft\Office\Excel\Addins\[ProgID]` 中设置 `LoadBehavior = 3`。
- **诊断日志**：设置环境变量 `VSTO_LOGALERTS = 1`。这样当 Excel 加载失败时，会弹出详细的错误窗口。

## 4. 彻底移除 VSTO 残留（冲突预防）
如果项目之前尝试过 VSTO 模式，残留的注册表项会导致 COM 模式失效：
- **必须清理的路径**：
  - `HKCU:\Software\Microsoft\VSTO\SolutionMetadata`（最常见的冲突源）
  - `HKCU:\Software\Microsoft\Office\Excel\Addins\[ProgID]` 下的 `Manifest` 键值（COM 模式不需要此键）。
- **缓存清理**：运行 `rundll32 dfshim CleanOnlineAppCache` 清理 ClickOnce 缓存。

## 5. 开发代码要点
- **ProgID 唯一性**：确保 `[ProgId("QSBar.WpsAddIn")]` 在整个系统中唯一。
- **CLSID 固定**：使用固定的 `[Guid("...")]`，方便在安装脚本中手动补齐注册表项。
- **静态 App 引用**：由于不再使用 `Globals.ThisAddIn`，需在 `OnConnection` 中将 `Application` 对象保存到静态变量（如 `WpsExcelAddIn.App`），供其他工具类调用。
- **权限管理**：`RegAsm` 注册 64 位和 32 位 COM 通常需要管理员权限，但在脚本中可以通过强制写入 `HKCU:\Software\Classes\CLSID` 来实现用户级别的 COM 注册。

## 6. 安装脚本逻辑 (`quick_setup.ps1`)
一个完美的安装脚本应包含：
1. **停止相关进程**：确保 DLL 不被锁定。
2. **清理旧项**：清理禁用列表和旧的 VSTO 记录。
3. **编译项目**：生成最新的 DLL。
4. **COM 注册**：调用 `RegAsm` 并手动补齐 `InprocServer32` 注册表项（指向 `mscoree.dll`）。
5. **WPS 白名单**：自动补齐 WPS 所需的所有白名单路径。
6. **自检测试**：通过 `New-Object -ComObject [ProgID]` 验证注册是否成功。

---
*最后更新日期：2026-01-17*
*所属项目：QSBar (COM Version)*
