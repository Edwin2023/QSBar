# QSBar 撤销功能设计

日期：2026-08-11

## 背景

COM 加载项通过 Interop 写单元格（`Range.Value2`、`NumberFormatLocal`、`Font` 等）会清空 Excel 的原生撤销栈。这是 Office 的既定行为，VSTO 和 COM 加载项都一样，没有绕过的办法。用户点了 QSBar 的数值化之后按 Ctrl+Z 毫无反应，只能靠没保存前关掉重开。

结论：撤销只能由插件自己实现一套。

## 范围

只覆盖**修改单元格值/格式**的命令。这类命令作用于选中区域，快照成本低、还原精确。

接入撤销的命令：

| 文件 | 命令 | 快照方式 |
|---|---|---|
| `DataCommands.cs` | `NormalizeNumbers` | `Capture` |
| `DataCommands.cs` | `Textify` | `Capture` |
| `FormatCommands.cs` | `ApplyAccountingNumberFormatLocal`（覆盖 `Accounting0/2/3`，Ctrl+6/7/8） | `Capture` |
| `FormatCommands.cs` | `YiWanFormat`（Ctrl+9） | `Capture` |
| `FormatCommands.cs` | `WrapText` | `Capture` |
| `FormatCommands.cs` | `SetGradingStyle` | `Capture` |
| `FormatCommands.cs` | `SetGrading` | `CaptureOutline` |
| `LegacyAppCommands.cs` | `LockFormula` | `Capture`（见下文特例） |
| `SheetCommands.cs` | `DeleteHyperlinks` | `Capture` |

**不接入**的命令（结构变更类）：删除空行、合并工作表、全表数值化、断开外链、生成表目录/文件目录、显示隐藏表。这些执行后一律清空 QSBar 撤销栈，Ctrl+Z 行为回到现状。

**不需要**撤销的命令：选中可见单元格、选中非空单元格、全选图片、图片缩放、三个导出、切换计算模式。

`ClearStyle`、`RowStyle`、`MultiAreaGroup`、`MultiAreaUngroup`、`PasteExternalLinks`、`LockValidationCells`、`ExpandPivotTable`、`CollapsePivotTable`、`MergeWorkbooks`、`FileDirectory` 有方法实现但 `Ribbon.xml` 里没有对应按钮、也没绑快捷键，用户当前触发不到，本次不接入。将来给它们加按钮时按同样规则决定是否 `Capture`。

## 架构

新增 `QSBar/UndoManager.cs`，静态类，与现有 `FormatCommands` 等命令类同级。

### 对外接口

```csharp
static void Capture(Excel.Range target, string descriptionCn, string descriptionEn)
static void CaptureOutline(Excel.Range target, string descriptionCn, string descriptionEn)
static void Undo()
static bool CanUndo { get; }
static void Clear()
static void Shutdown()
```

`CanUndo` 只读托管字段（栈的 `Count`），不触碰 COM——它会在键盘钩子回调里被同步调用，任何 COM 调用都会拖慢按键响应甚至死锁。

描述文本按 Ribbon 现有双语约定分中英两个参数，与 `WpsExcelAddIn.UseChineseRibbon` 配合。

### 快照存储：隐藏备份工作簿

懒创建一个不可见工作簿作为快照容器：

```csharp
_backupBook = app.Workbooks.Add();
_backupBook.Windows[1].Visible = false;
```

创建期间置 `app.ScreenUpdating = false`、`app.EnableEvents = false`，避免闪屏和事件回灌。

每次 `Capture`：在备份工作簿里新建一张 sheet，把源区域逐 Area 复制过去，**保持原始地址**：

```csharp
foreach (Excel.Range area in target.Areas)
    area.Copy(backupSheet.Range[area.Address]);
```

地址原样保持，撤销时按同地址反向 Copy 即可，不需要额外的坐标映射。

选这个方案而不是内存数组的理由：`Range.Copy` 一次带回值、公式、数字格式、字体、填充、边框、对齐、合并单元格、条件格式，不需要逐格枚举格式属性。逐格读 `Interior.Color` / `Font.*` 在几万行的 BOQ 表上会卡到无法接受——`SmartClearAndRestyle` 现有的逐格读取已经是性能痛点。

### 多重选定区域

`Range.Copy` 不支持多重选定区域，对 `Areas.Count > 1` 的 Range 直接调用会抛"该命令不能用于多重选定区域"。而多区域在本插件里是常态：

- `SetGradingStyle` 的实现明确在遍历 `rng.Areas`
- `LockFormula` 的 `SpecialCells(xlCellTypeFormulas)` 返回的必然是不连续区域
- 用户按 Ctrl 手动多选很常见

因此 `UndoEntry` 存一个地址列表，快照和还原都逐 Area 进行。

**不能**用外接矩形（`Range[topLeft, bottomRight]`）代替：还原时会把用户当初没选中的单元格一起覆盖回旧值，这是真实的数据丢失。

### SetGrading 特例

`SetGrading` 改的是 `Rows[i].OutlineLevel`，这是行属性。跨工作簿复制单元格区域带不回 `OutlineLevel`，而复制整行又会在还原时覆盖选中列以外的数据。

所以这条单独走 `CaptureOutline`：只记录一个 `int[]` 行大纲级别数组（连同起始行号），撤销时逐行写回 `OutlineLevel`。内存和耗时都可忽略，不占用备份工作簿的 sheet。

### LockFormula 特例

`LockFormula` 直接用 `Selection` 且没有 `Intersect(UsedRange)`，用户选整列时是百万单元格。它的快照必须建在 `SpecialCells(xlCellTypeFormulas)` 的结果上，而不是 `Selection` 上，否则会被规模阈值一律拒绝。

即：先取 `formulas`，再 `Capture(formulas, ...)`，再执行改写。

### 撤销栈

`List<UndoEntry>`，上限 10 条。超出上限时移除最老的一条并删掉它对应的备份 sheet（`DisplayAlerts = false` 后 `backupSheet.Delete()`）。

`UndoEntry` 字段：

- 源工作簿 `Name`（用于关闭时匹配）
- 源工作表 `Name`
- 区域地址列表 `List<string>`
- 备份 sheet 名（`CaptureOutline` 的条目为 `null`）
- 行大纲级别数组与起始行（仅 `CaptureOutline` 条目）
- 中英文描述

### 规模阈值

`Capture` 前累加各 Area 的 `Cells.Count`，超过 300000 就不记录快照，弹 Toast 提示"区域过大，本次操作不可撤销 / Range too large, this action cannot be undone"，然后命令照常执行。

这是为了防止在大 BOQ 表上快照本身把 Excel 拖死。阈值定为常量，后续可调。

## Ctrl+Z 劫持

### KeyboardHook 扩展

`KeyboardHook.AddShortcut` 增加一个 `Func<bool> shouldIntercept` 参数，默认 `null` 表示无条件拦截，现有 Ctrl+3~0 的注册行为完全不变。

`HookCallback` 里，命中快捷键后先同步调用 `shouldIntercept()`：

- 返回 `true`：走现有路径——启动 10ms `Timer` 延迟执行 action，返回 `(IntPtr)1` 吃掉按键
- 返回 `false`：不执行 action，直接 `CallNextHookEx` 放行给 Excel

Ctrl+Z 注册为 `AddShortcut(true, false, Keys.Z, UndoManager.Undo, () => UndoManager.CanUndo)`。

延迟执行的机制必须保留：现有注释已说明，在钩子回调里同步做重活会导致系统判定钩子超时并把按键透传给 Excel。

### 安全网一：栈空即透传

`CanUndo == false` 时 Ctrl+Z 完全属于 Excel。这是最基本的一条——没有 QSBar 快照时插件不该出现在这条链路上。

### 安全网二：手动编辑即弃栈

插件目前一个 Excel 事件都没订阅。在 `OnConnection` 里新增：

```csharp
_application.SheetChange += OnSheetChange;
_application.WorkbookBeforeClose += OnWorkbookBeforeClose;
```

`OnSheetChange`：若不是 QSBar 自己引发的（`UndoManager` 内部一个 `_suppressChangeTracking` 标志位在 `Capture`/`Undo` 期间置位），就 `UndoManager.Clear()`。

没有这一条的话，"手动改两个格 → 按 Ctrl+Z"会莫名其妙撤销掉十分钟前的数值化，比没有撤销功能更糟。

`OnWorkbookBeforeClose`：移除该工作簿相关的所有快照条目，避免撤销到已关闭的工作簿上。

`OnDisconnection` / `OnBeginShutdown` 里调 `UndoManager.Shutdown()`：`DisplayAlerts = false` 后 `_backupBook.Close(false)`，释放 COM 引用。

### 还原前校验

`Undo()` 执行前校验源工作簿和工作表仍存在（按名字查找，`try/catch` 兜住 COM 异常）。找不到就丢弃该条目、提示用户、不做任何写入。

工作表被改名的情况按"找不到"处理——这是保守选择，比猜错工作表写坏数据好。

## Ribbon

`groupFormat` 里新增按钮：

```xml
<button id='btnUndo' getLabel='GetRibbonLabel' onAction='OnUndo' imageMso='Undo'
  size='normal' getEnabled='GetUndoEnabled'
  getScreentip='GetRibbonScreentip' getSupertip='GetRibbonSupertip' />
```

标签"撤销 QSBar / Undo QSBar"，加入 `WpsExcelAddIn` 的 `RibbonLabels` / `RibbonScreentips` / `RibbonSupertips` 三个字典，supertip 里写明快捷键 Ctrl+Z 与"只能撤销 QSBar 自己的操作"。

作用有两个：灰显状态让用户看得出当前有没有可撤销的东西；键盘钩子万一失效时的兜底入口。

`Capture` 和 `Undo` 之后调 `WpsExcelAddIn` 新增的一个静态方法 `InvalidateUndoButton()`（照 `InvalidateCalcButtons` 的写法）刷新灰显。

## 已知代价

1. **`Range.Copy` 会清空系统剪贴板。** 用户复制了内容 → 点数值化 → 剪贴板没了。已确认接受。用 Win32 备份恢复剪贴板在 Excel 进程里对大对象和多格式并不可靠，不做。
2. **备份工作簿常驻进程。** 内存随快照区域增长，靠 10 条上限和 30 万单元格阈值控制。
3. **WPS 兼容性待实测。** 隐藏工作簿、跨工作簿 `Range.Copy`、`OutlineLevel` 写回三项需要在 WPS 和 Excel 上各验一遍。

## 测试

`QSBar.Tests` 目前因 .NET 10 SDK 的 CET 问题无法用 `dotnet build` 编译，且 `UndoManager` 强依赖 Excel COM，单元测试价值有限。以手工验证为主，Excel 和 WPS 各跑一遍：

1. 数值化一个区域 → Ctrl+Z → 值、数字格式、`ShrinkToFit` 全部还原
2. 连续 3 个不同命令 → 连按 3 次 Ctrl+Z → 逐步回退，第 4 次 Ctrl+Z 落到 Excel 自己
3. 撤销栈为空时 Ctrl+Z → Excel 原生撤销正常工作
4. QSBar 操作后手动改一格 → Ctrl+Z → 撤销的是手动改动，QSBar 栈已弃
5. Ctrl 多选 3 个不连续区域 → 会计格式 → Ctrl+Z → 三个区域都还原，区域外单元格未被触碰
6. `SetGrading` 设分级 → Ctrl+Z → 行大纲级别还原
7. `SetGradingStyle` → Ctrl+Z → 填充、字体、边框还原
8. 选整列 `LockFormula` → Ctrl+Z → 公式还原为相对引用
9. 超过 30 万单元格的区域执行数值化 → Toast 提示不可撤销 → Ctrl+Z 落到 Excel
10. 快照后关闭源工作簿 → Ctrl+Z → 不崩溃，提示条目失效
11. 连做 12 次可撤销操作 → 只能回退 10 步，备份工作簿里的 sheet 数不超过 10
12. 关闭 Excel → 进程干净退出，无残留 EXCEL.EXE
