# WPS 插件安装与排错指南

本文档专门针对 **WPS Office** (尤其是在 Windows 下的个人版与企业版) 的插件加载问题进行汇总。WPS 对 COM 插件的加载机制比 Excel 更为严格，且存在 32 位与 64 位版本混用的复杂情况，导致插件经常出现“装了没反应”的现象。

## 1. 核心痛点：为什么 WPS 看不到插件？

最常见的三大原因：
1.  **位数不匹配**：系统是 64 位，但装了 32 位 WPS，导致找不到 64 位注册表项。
2.  **默认禁用 (白名单机制)**：WPS 出于安全考虑，默认不加载未在白名单 (`AddinsWL`) 中的插件。
3.  **残留禁用项**：如果插件崩溃过一次，WPS 会将其永久拉黑，即使重装也不会恢复。

---

## 2. 解决方案详解

### 2.1 必须注册 32 位兼容项 (Wow6432Node)

大多数用户的 Windows 是 64 位的，但为了兼容旧宏，安装的 WPS 往往是 **32 位** 版本。
- **现象**：Excel (64位) 能看到插件，WPS (32位) 看不到。
- **原理**：32 位程序读取注册表时，会自动重定向到 `HKLM\Software\Wow6432Node`。如果你的插件只注册到了 `HKLM\Software\Classes`，32 位 WPS 是“瞎”的。
- **解决**：安装脚本必须同时写入以下两个位置：
    *   `HKLM\Software\Classes\CLSID\{GUID}` (64位程序用)
    *   `HKLM\Software\Wow6432Node\Classes\CLSID\{GUID}` (32位程序用)
    *   `HKLM\Software\Wow6432Node\Kingsoft\Office\WPS\Addins\QSBar.WpsAddIn` (WPS 专用)

### 2.2 强制开启白名单 (AddinsWL)

这是 WPS 特有的安全机制。如果注册表中没有这个“通行证”，插件可能被静默拦截。
- **位置**：
    *   `HKCU\Software\Kingsoft\Office\WPS\AddinsWL` (文字)
    *   `HKCU\Software\Kingsoft\Office\ET\AddinsWL` (表格)
- **操作**：新建一个 **字符串值**，名称为插件的 `ProgID` (例如 `QSBar.WpsAddIn`)，数值数据设为 `1`。

### 2.3 清理“禁用项目” (Resiliency)

如果 WPS 启动时插件报错，WPS 会询问“是否禁用此插件”。一旦用户点了“是”，WPS 就会在注册表中生成一个二进制值，永久屏蔽该插件。
- **检查路径**：
    *   `HKCU\Software\Kingsoft\Office\ET\Resiliency\DisabledItems`
- **操作**：删除该路径下的所有二进制值（或右键删除整个 `DisabledItems` 文件夹）。

---

## 3. 排错流程 (Troubleshooting)

如果安装后 WPS 表格中没有出现 `QS 工具箱` 选项卡，请按以下步骤操作：

### 第一步：使用诊断工具
运行项目根目录下的 `scripts\Debug-QSBar-State.ps1` 脚本。
- **检查输出**：
  - 是否显示 `[WPS Addin] Found`？
  - 是否显示 `[WPS Addin] Found (32-bit)`？ (如果你的 WPS 是 32 位的，这一项必须有)
  - `LoadBehavior` 是否为 `3`？(如果是 `2` 表示被软禁用，如果是 `0` 表示未配置加载)

### 第二步：手动检查 WPS 界面
1.  打开 WPS 表格。
2.  点击 **开发工具** -> **COM 加载项**。
3.  查看列表中是否有 `QSBar`：
    *   **没有**：说明注册表完全没写进去，或者 32/64 位没对上。请以管理员权限重新运行 `quick_setup.ps1`。
    *   **有，但没勾选**：尝试勾选它。
        *   如果勾选后立即自动取消勾选，说明插件加载时报错了（缺少依赖或代码错误）。
        *   如果提示“未加载。加载 COM 加载项时出现运行错误”，通常是 DLL 版本不对或依赖库缺失。
    *   **有，且已勾选，但界面没显示**：说明 Ribbon XML 加载失败（常见于 XML 格式错误）。

### 第三步：强制启用 (暴力法)
如果以上都不行，可以尝试手动导入以下注册表文件 (`fix_wps.reg`)：

```registry
Windows Registry Editor Version 5.00

; --- 32位 WPS 兼容项 ---
[HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn]
"FriendlyName"="QSBar (COM)"
"Description"="QSBar COM Add-in for Excel and WPS"
"LoadBehavior"=dword:00000003
"CommandLineSafe"=dword:00000001

; --- 白名单 (用户层) ---
[HKEY_CURRENT_USER\Software\Kingsoft\Office\ET\AddinsWL]
"QSBar.WpsAddIn"="1"

; --- 确保加载行为 ---
[HKEY_CURRENT_USER\Software\Kingsoft\Office\ET\Addins\QSBar.WpsAddIn]
"LoadBehavior"=dword:00000003
```

## 4. 附录：技术细节
- **ProgID**: `QSBar.WpsAddIn`
- **CLSID**: `{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}`
- **DLL名称**: `QSBar.dll`
