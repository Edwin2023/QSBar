using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using AddInDesignerObjects;

namespace QSBar
{
    [ComVisible(true)]
    [Guid("D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF")]
    [ProgId("QSBar.WpsAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class WpsExcelAddIn : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private Excel.Application _application;
        public static Excel.Application App { get; private set; }

        private static Office.IRibbonUI _ribbon;
        private static int _lastCalcMode = -1;
        private static Timer _calcTimer;
        private static Control _uiInvoker;
        private static KeyboardHook _keyboardHook;
        private const string LanguageRegistryPath = @"Software\QSBar";
        private const string LanguageRegistryName = "RibbonLanguage";
        private static bool _useChineseRibbon = LoadRibbonLanguage();
        internal static bool UseChineseRibbon => _useChineseRibbon;

        private static readonly Dictionary<string, string[]> RibbonLabels = new Dictionary<string, string[]>
        {
            { "tabQSBar", new string[] { "QS工具箱", "QS Toolbox" } },
            { "groupFormat", new string[] { "格式/排版", "Format / Layout" } },
            { "groupBQ", new string[] { "批量处理", "Batch Tools" } },
            { "groupGrading", new string[] { "分级/筛选", "Outline / Index" } },
            { "groupPhotoTools", new string[] { "对象选择", "Selection" } },
            { "groupSheets", new string[] { "工作表", "Worksheet" } },
            { "groupCalc", new string[] { "计算模式", "Calculation" } },
            { "groupHelpUpdate", new string[] { "帮助更新", "Help / Update" } },
            { "btnBatch", new string[] { "批量处理", "Batch Process" } },
            { "btnNormalize", new string[] { "数值化", "Normalize" } },
            { "btnTextify", new string[] { "文本化", "Textify" } },
            { "btnLock", new string[] { "锁定公式", "Lock Formulas" } },
            { "btnSetGrading", new string[] { "设置分级", "Set Outline" } },
            { "btnSetGradingStyle", new string[] { "分级样式", "Outline Style" } },
            { "btnOutSheet", new string[] { "导出当前表", "Export Sheet" } },
            { "btnOutStd", new string[] { "导出标准报表", "Export Standard" } },
            { "btnOutInt", new string[] { "导出内部报表", "Export Internal" } },
            { "btnMergeSheets", new string[] { "合并工作表", "Merge Sheets" } },
            { "btnSheetIndex", new string[] { "生成表目录", "Sheet Index" } },
            { "btnFileDir", new string[] { "生成文件目录", "File Index" } },
            { "btnSelVisible", new string[] { "选中可见单元格", "Visible Cells" } },
            { "btnSelectAllPictures", new string[] { "全选图片", "Select Pictures" } },
            { "btnResize1", new string[] { "1倍大小", "1x" } },
            { "btnResize2", new string[] { "2倍大小", "2x" } },
            { "btnResize4", new string[] { "4倍大小", "4x" } },
            { "btnForceRefresh", new string[] { "强制刷新", "Force Refresh" } },
            { "btnValOnly", new string[] { "全表数值化", "Values Only" } },
            { "btnUnhideSheets", new string[] { "显示隐藏表", "Unhide Sheets" } },
            { "btnBreakLinks", new string[] { "删除外链接", "Break Links" } },
            { "btnDeleteLinks", new string[] { "删除超链接", "Delete Hyperlinks" } },
            { "btnDeleteEmptyRows", new string[] { "删除表空行", "Delete Empty Rows" } },
            { "btnCalcAuto", new string[] { "切到手动计算", "Switch to Manual" } },
            { "btnCalcManual", new string[] { "切到自动计算", "Switch to Automatic" } },
            { "btnHelp", new string[] { "使用帮助", "Help" } },
            { "btnCheckUpdate", new string[] { "检查更新", "Check Updates" } },
            { "btnUpdate", new string[] { "重启更新", "Restart Update" } }
        };

        private static readonly Dictionary<string, string[]> RibbonScreentips = new Dictionary<string, string[]>
        {
            { "btnBatch", new string[] { "批量处理 / 公式计算 (Ctrl+3)", "Batch process / formula calculation (Ctrl+3)" } },
            { "btnNormalize", new string[] { "规范数值", "Normalize data" } },
            { "btnTextify", new string[] { "转换为文本", "Convert to text" } },
            { "btnLock", new string[] { "锁定 / 解锁公式", "Lock / unlock formulas" } },
            { "btnSetGrading", new string[] { "设置大纲级别", "Set outline levels" } },
            { "btnSetGradingStyle", new string[] { "设置分级样式", "Set outline styles" } },
            { "btnOutSheet", new string[] { "导出当前工作表", "Export current sheet" } },
            { "btnOutStd", new string[] { "导出标准报表", "Export standard report" } },
            { "btnOutInt", new string[] { "导出内部报表", "Export internal report" } },
            { "btnMergeSheets", new string[] { "合并工作表", "Merge worksheets" } },
            { "btnSheetIndex", new string[] { "生成工作表目录", "Create sheet index" } },
            { "btnFileDir", new string[] { "生成文件目录", "Create file index" } },
            { "btnSelVisible", new string[] { "选中可见单元格 (Ctrl+5)", "Select visible cells (Ctrl+5)" } },
            { "btnSelectAllPictures", new string[] { "选择全部图片", "Select all pictures" } },
            { "btnForceRefresh", new string[] { "强制刷新", "Force refresh" } },
            { "btnValOnly", new string[] { "整表转换为数值", "Convert entire sheet to values" } },
            { "btnUnhideSheets", new string[] { "显示全部隐藏工作表", "Unhide all sheets" } },
            { "btnBreakLinks", new string[] { "断开外部链接", "Break external links" } },
            { "btnDeleteLinks", new string[] { "删除超链接", "Delete hyperlinks" } },
            { "btnDeleteEmptyRows", new string[] { "删除空行", "Delete empty rows" } },
            { "btnCalcAuto", new string[] { "当前：自动计算 (Ctrl+0 / F10)", "Current: Automatic (Ctrl+0 / F10)" } },
            { "btnCalcManual", new string[] { "当前：手动计算 (Ctrl+0 / F10)", "Current: Manual (Ctrl+0 / F10)" } },
            { "btnHelp", new string[] { "帮助", "Help" } },
            { "btnCheckUpdate", new string[] { "检查更新", "Check updates" } },
            { "btnUpdate", new string[] { "重启更新", "Restart update" } }
        };

        private static readonly Dictionary<string, string[]> RibbonSupertips = new Dictionary<string, string[]>
        {
            { "btnBatch", new string[] { "批量处理或公式计算，通常用于大数据量表格。快捷键 Ctrl+3。", "Batch processing or formula calculation, usually for large data sets. Shortcut: Ctrl+3." } },
            { "btnNormalize", new string[] { "规范化选中区域中的数据。", "Normalize the data in the selected range." } },
            { "btnTextify", new string[] { "强制把选中单元格转换为文本格式。", "Force the selected cells to text format." } },
            { "btnLock", new string[] { "锁定或解锁选中区域中的公式。", "Lock or unlock formulas in the selected range." } },
            { "btnSetGrading", new string[] { "根据内容自动设置大纲级别；支持 【 / [[、《 / <<、{。", "Automatically set outline levels based on content; supports 【 / [[, 《 / <<, and {." } },
            { "btnSetGradingStyle", new string[] { "按大纲级别应用不同单元格样式。", "Apply different cell styles for outline levels." } },
            { "btnOutSheet", new string[] { "把当前工作表导出为数值。", "Export the current worksheet as values." } },
            { "btnOutStd", new string[] { "删除隐藏行列，然后把所有工作表导出为数值。", "Remove hidden rows and columns, then export all sheets as values." } },
            { "btnOutInt", new string[] { "把所有工作表导出为数值。", "Export all sheets as values." } },
            { "btnMergeSheets", new string[] { "把多个工作表合并到一个工作表。", "Merge multiple worksheets into one worksheet." } },
            { "btnSheetIndex", new string[] { "为当前工作簿生成所有工作表目录。", "Generate an index of all worksheets in the current workbook." } },
            { "btnFileDir", new string[] { "选择文件，并在当前工作簿中生成目录。", "Select files and generate an index for them in the current workbook." } },
            { "btnSelVisible", new string[] { "选择当前工作表中的可见单元格。", "Select the visible cells in the current worksheet." } },
            { "btnSelectAllPictures", new string[] { "选择当前工作表中的全部图片。", "Select all pictures in the current worksheet." } },
            { "btnForceRefresh", new string[] { "强制刷新选中区域中的全部公式。", "Force refresh all formulas in the selected range." } },
            { "btnValOnly", new string[] { "把当前工作表中的全部公式转换为数值。", "Convert all formulas in the current worksheet to values." } },
            { "btnUnhideSheets", new string[] { "显示当前工作簿中的全部隐藏工作表。", "Unhide all worksheets in the current workbook." } },
            { "btnBreakLinks", new string[] { "断开全部外部工作簿引用，并清理异常名称。", "Break all external workbook references and clean invalid names." } },
            { "btnDeleteLinks", new string[] { "删除选中区域中的全部超链接。", "Delete all hyperlinks in the selected range." } },
            { "btnDeleteEmptyRows", new string[] { "删除选中区域中的全部空行。", "Delete all empty rows in the selected range." } },
            { "btnCalcAuto", new string[] { "把计算模式切换为手动。快捷键 Ctrl+0 / F10。", "Switch calculation mode to manual. Shortcut: Ctrl+0 / F10." } },
            { "btnCalcManual", new string[] { "把计算模式切换为自动。快捷键 Ctrl+0 / F10。", "Switch calculation mode to automatic. Shortcut: Ctrl+0 / F10." } },
            { "btnHelp", new string[] { "查看 QS 工具箱版本、快捷键和更新检查。", "View QS Toolbox version info, shortcuts, and update checks." } },
            { "btnCheckUpdate", new string[] { "手动检查是否有新版本。", "Manually check whether a new version is available." } },
            { "btnUpdate", new string[] { "已有新版本；点击确认重启并完成更新。", "A new version is available; click to confirm restart and finish updating." } }
        };

        public static void RefreshRibbon()
        {
            if (_ribbon != null) _ribbon.Invalidate();
        }

        private static bool LoadRibbonLanguage()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(LanguageRegistryPath))
                {
                    object value = key == null ? null : key.GetValue(LanguageRegistryName);
                    string language = value == null ? null : value.ToString();
                    if (string.Equals(language, "zh-CN", StringComparison.OrdinalIgnoreCase)) return true;
                    if (string.Equals(language, "en-US", StringComparison.OrdinalIgnoreCase)) return false;
                }
            }
            catch { }

            try
            {
                return CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.Equals("zh", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static void SaveRibbonLanguage()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(LanguageRegistryPath))
                {
                    if (key != null) key.SetValue(LanguageRegistryName, _useChineseRibbon ? "zh-CN" : "en-US");
                }
            }
            catch { }
        }

        private static string GetLocalizedText(Dictionary<string, string[]> source, string controlId)
        {
            if (string.IsNullOrEmpty(controlId)) return string.Empty;
            string[] values;
            if (source != null && source.TryGetValue(controlId, out values) && values != null && values.Length >= 2)
            {
                return _useChineseRibbon ? values[0] : values[1];
            }
            return string.Empty;
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _application = (Excel.Application)Application;
                App = _application;

                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                // StartCalcTimer();  // 临时禁用：排查 Ribbon 黑条问题
                RegisterShortcuts();

                // Run one silent update check at startup, fully off the UI thread.
                var _ = Task.Run(() => UpdateManager.CheckForUpdateAsync(true, false));
                var __ = Task.Run(() => AnalyticsHelper.Ping());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during QSBar connection: " + ex.Message);
            }
        }

        private void AddShortcuts(bool ctrl, bool alt, Keys mainKey, Keys numKey, Action action)
        {
            _keyboardHook.AddShortcut(ctrl, alt, mainKey, action);
            if (numKey != Keys.None) _keyboardHook.AddShortcut(ctrl, alt, numKey, action);
        }

        private void RegisterShortcuts()
        {
            if (_keyboardHook != null) return;
            try
            {
                _keyboardHook = new KeyboardHook();

                // Ctrl + Key
                AddShortcuts(true, false, Keys.D3, Keys.NumPad3, () => DataCommands.BatchProcess());
                AddShortcuts(true, false, Keys.D4, Keys.NumPad4, () => FormatCommands.SelectNonEmptyCells());
                AddShortcuts(true, false, Keys.D5, Keys.NumPad5, () => FormatCommands.SelectVisibleCells());
                AddShortcuts(true, false, Keys.D6, Keys.NumPad6, () => FormatCommands.Accounting0());
                AddShortcuts(true, false, Keys.D7, Keys.NumPad7, () => FormatCommands.Accounting2());
                AddShortcuts(true, false, Keys.D8, Keys.NumPad8, () => FormatCommands.Accounting3());
                AddShortcuts(true, false, Keys.D9, Keys.NumPad9, () => FormatCommands.YiWanFormat());
                AddShortcuts(true, false, Keys.D0, Keys.NumPad0, () => LegacyAppCommands.ToggleCalculation());

                // Ctrl + Alt + Key
                AddShortcuts(true, true, Keys.D1, Keys.NumPad1, () => DataCommands.ExpandPivotTable());
                AddShortcuts(true, true, Keys.D2, Keys.NumPad2, () => DataCommands.CollapsePivotTable());
            }
            catch { }
        }

        private void UnregisterShortcuts()
        {
            if (_keyboardHook != null)
            {
                _keyboardHook.Dispose();
                _keyboardHook = null;
            }
        }

        private void StartCalcTimer()
        {
            try
            {
                if (_calcTimer == null)
                {
                    _calcTimer = new Timer();
                    _calcTimer.Interval = 1000;
                    _calcTimer.Tick += (s, e) =>
                    {
                        try
                        {
                            if (App == null) return;
                            int mode = (int)App.Calculation;
                            if (mode == _lastCalcMode) return;
                            _lastCalcMode = mode;
                            if (_ribbon != null)
                            {
                                _ribbon.InvalidateControl("btnCalcAuto");
                                _ribbon.InvalidateControl("btnCalcManual");
                            }
                        }
                        catch { }
                    };
                    _calcTimer.Start();
                }
            }
            catch { }
        }

        private void StopCalcTimer()
        {
            if (_calcTimer != null)
            {
                _calcTimer.Stop();
                _calcTimer.Dispose();
                _calcTimer = null;
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            UnregisterShortcuts();
            StopCalcTimer();
            if (_uiInvoker != null)
            {
                try { _uiInvoker.Dispose(); } catch { }
                _uiInvoker = null;
            }
            _application = null;
            App = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public void OnLoad(Office.IRibbonUI ribbon)
        {
            try
            {
                _ribbon = ribbon;
                if (_uiInvoker == null || _uiInvoker.IsDisposed)
                {
                    _uiInvoker = new Control();
                    var _ = _uiInvoker.Handle;
                }
                RegisterShortcuts();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("OnLoad Error: " + ex.ToString());
            }
        }

        public string GetCustomUI(string RibbonID)
        {
            return GetResourceText("QSBar.Ribbon.xml");
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #region Ribbon Callbacks

        public string GetRibbonLabel(Office.IRibbonControl control)
        {
            if (control != null && string.Equals(control.Id, "btnToggleLanguage", StringComparison.OrdinalIgnoreCase))
            {
                return _useChineseRibbon ? "English" : "中文";
            }

            string label = GetLocalizedText(RibbonLabels, control == null ? null : control.Id);
            return string.IsNullOrEmpty(label) && control != null ? control.Id : label;
        }

        public string GetRibbonScreentip(Office.IRibbonControl control)
        {
            if (control != null && string.Equals(control.Id, "btnToggleLanguage", StringComparison.OrdinalIgnoreCase))
            {
                return _useChineseRibbon ? "切换到英文界面" : "Switch to Chinese UI";
            }

            string screentip = GetLocalizedText(RibbonScreentips, control == null ? null : control.Id);
            if (!string.IsNullOrEmpty(screentip)) return screentip;
            return GetRibbonLabel(control);
        }

        public string GetRibbonSupertip(Office.IRibbonControl control)
        {
            if (control != null && string.Equals(control.Id, "btnToggleLanguage", StringComparison.OrdinalIgnoreCase))
            {
                return _useChineseRibbon ? "将 QS 工具箱 Ribbon 切换为英文。此选择会保存到当前 Windows 用户。" : "Switch the QS Toolbox ribbon to Chinese. This preference is saved for the current Windows user.";
            }

            string supertip = GetLocalizedText(RibbonSupertips, control == null ? null : control.Id);
            return string.IsNullOrEmpty(supertip) ? GetRibbonScreentip(control) : supertip;
        }

        public void OnToggleLanguage(Office.IRibbonControl control)
        {
            _useChineseRibbon = !_useChineseRibbon;
            SaveRibbonLanguage();
            RefreshRibbon();
        }

        public bool GetHelpNormalVisible(Office.IRibbonControl control)
        {
            return !UpdateManager.HasNewVersion;
        }

        public bool GetHelpUpdateVisible(Office.IRibbonControl control)
        {
            return UpdateManager.HasNewVersion;
        }

        public async void OnCheckUpdate(Office.IRibbonControl control)
        {
            try
            {
                await UpdateManager.CheckForUpdateAsync(false, false);
                if (UpdateManager.HasNewVersion)
                {
                    await UpdateManager.StartUpdateFlow();
                }
                RefreshRibbon();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while checking for updates: " + ex.Message, "Update Error");
            }
        }

        public async void OnRestartUpdate(Office.IRibbonControl control)
        {
            try
            {
                await UpdateManager.StartUpdateFlow();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while starting the update: " + ex.Message, "Update Error");
            }
        }

        public void OnBatchProcess(Office.IRibbonControl control) { DataCommands.BatchProcess(); }
        public void OnExportCurrentSheet(Office.IRibbonControl control) { ExportCommands.ExportCurrentSheet(); }
        public void OnExportStandardReport(Office.IRibbonControl control) { ExportCommands.ExportStandardReport(); }
        public void OnExportInternalReport(Office.IRibbonControl control) { ExportCommands.ExportInternalReport(); }
        public void OnConvertAllToValues(Office.IRibbonControl control) { ExportCommands.ConvertAllToValues(); }
        public void OnShowHelp(Office.IRibbonControl control)
        {
            LegacyAppCommands.ShowHelp();
        }
        public void OnNormalizeNumbers(Office.IRibbonControl control) { DataCommands.NormalizeNumbers(); }
        public void OnTextify(Office.IRibbonControl control) { DataCommands.Textify(); }
        public void OnLockFormula(Office.IRibbonControl control) { LegacyAppCommands.LockFormula(); }
        public void OnWrapText(Office.IRibbonControl control) { FormatCommands.WrapText(); }
        public void OnAccounting0(Office.IRibbonControl control) { FormatCommands.Accounting0(); }
        public void OnAccounting2(Office.IRibbonControl control) { FormatCommands.Accounting2(); }
        public void OnAccounting3(Office.IRibbonControl control) { FormatCommands.Accounting3(); }
        public void OnYiWanFormat(Office.IRibbonControl control) { FormatCommands.YiWanFormat(); }
        public void OnSetGrading(Office.IRibbonControl control) { FormatCommands.SetGrading(); }
        public void OnSetGradingStyle(Office.IRibbonControl control) { FormatCommands.SetGradingStyle(); }
        public void OnClearStyle(Office.IRibbonControl control) { FormatCommands.ClearStyle(); }
        public void OnBreakLinks(Office.IRibbonControl control) { FormatCommands.BreakExternalLinks(); }
        public void OnMultiAreaGroup(Office.IRibbonControl control) { FormatCommands.MultiAreaGroup(); }
        public void OnMultiAreaUngroup(Office.IRibbonControl control) { FormatCommands.MultiAreaUngroup(); }
        public void OnSelectVisibleCells(Office.IRibbonControl control) { FormatCommands.SelectVisibleCells(); }
        public void OnResizePictures(Office.IRibbonControl control)
        {
            int factor = 1;
            int f;
            if (control.Tag != null && int.TryParse(control.Tag, out f)) factor = f;
            PhotoCommands.ResizePictures(factor);
        }
        public void OnSelectAllPictures(Office.IRibbonControl control) { PhotoCommands.SelectAllPictures(); }
        public void OnCreateSheetIndex(Office.IRibbonControl control) { SheetCommands.CreateSheetIndex(); }
        public void OnCreateFileIndex(Office.IRibbonControl control) { SheetCommands.CreateFileIndex(); }
        public void OnMergeSheets(Office.IRibbonControl control) { SheetCommands.MergeSheets(); }
        public void OnDeleteHyperlinks(Office.IRibbonControl control) { SheetCommands.DeleteHyperlinks(); }
        public void OnForceRefresh(Office.IRibbonControl control) { DataCommands.ForceRefresh(); }
        public void OnDeleteEmptyRows(Office.IRibbonControl control) { SheetCommands.DeleteEmptyRows(); }
        public void OnUnhideAllSheets(Office.IRibbonControl control) { SheetCommands.UnhideAllSheets(); }
        public void OnFileDirectory(Office.IRibbonControl control) { SheetCommands.FileDirectory(); }

        public bool GetCalcAutoVisible(Office.IRibbonControl control)
        {
            try { return App != null && App.Calculation == Excel.XlCalculation.xlCalculationAutomatic; }
            catch { return true; }
        }

        public bool GetCalcManualVisible(Office.IRibbonControl control)
        {
            try { return App != null && App.Calculation != Excel.XlCalculation.xlCalculationAutomatic; }
            catch { return false; }
        }

        public void OnToggleCalculation(Office.IRibbonControl control)
        {
            LegacyAppCommands.ToggleCalculation();
        }

        public static void InvalidateCalcButtons()
        {
            if (_ribbon != null)
            {
                try
                {
                    _ribbon.InvalidateControl("btnCalcAuto");
                    _ribbon.InvalidateControl("btnCalcManual");
                }
                catch { }
            }
        }

        #endregion
    }
}
