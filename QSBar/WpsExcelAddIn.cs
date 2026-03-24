using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
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
        private static KeyboardHook _keyboardHook;
        private static uint _currentProcessId;

        public static void RefreshRibbon()
        {
            if (_ribbon != null) _ribbon.Invalidate();
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _application = (Excel.Application)Application;
                App = _application;
                _currentProcessId = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
                
                // EPPlus License
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                StartCalcTimer();
                RegisterShortcuts();
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
                RegisterShortcuts(); // 确保 Ribbon 加载后也尝试注册快捷键
            
                // 启动时静默检查更新（不打扰用户）
                var _ = Task.Run(() => UpdateManager.CheckForUpdateAsync(true));
            }
            catch (Exception ex)
            {
                // 记录日志或忽略，避免阻断加载
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

        public bool GetHelpNormalVisible(Office.IRibbonControl control)
        {
            return !UpdateManager.HasNewVersion;
        }

        public bool GetHelpUpdateVisible(Office.IRibbonControl control)
        {
            return UpdateManager.HasNewVersion;
        }

        public void OnRestartUpdate(Office.IRibbonControl control)
        {
            UpdateManager.StartUpdateFlow();
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
