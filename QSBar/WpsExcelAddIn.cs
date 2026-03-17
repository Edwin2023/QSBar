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
                AddShortcuts(true, false, Keys.D7, Keys.NumPad7, () => FormatCommands.WrapText());
                AddShortcuts(true, false, Keys.D8, Keys.NumPad8, () => FormatCommands.Accounting0());
                AddShortcuts(true, false, Keys.D9, Keys.NumPad9, () => FormatCommands.Accounting2());
                AddShortcuts(true, false, Keys.D0, Keys.NumPad0, () => FormatCommands.Accounting3());
                AddShortcuts(true, false, Keys.D6, Keys.NumPad6, () => FormatCommands.YiWanFormat());
                AddShortcuts(true, false, Keys.D1, Keys.NumPad1, () => ExportCommands.ShowLevel1());
                AddShortcuts(true, false, Keys.D2, Keys.NumPad2, () => ExportCommands.ShowLevel2());
                AddShortcuts(true, false, Keys.D3, Keys.NumPad3, () => ExportCommands.ShowLevel3());
                AddShortcuts(true, false, Keys.D4, Keys.NumPad4, () => ExportCommands.ShowLevel4());
                AddShortcuts(true, false, Keys.D5, Keys.NumPad5, () => FormatCommands.SelectVisibleCells());

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

        public void OnLoad(object ribbon)
        {
            try
            {
                _ribbon = ribbon as Office.IRibbonUI;
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

        public bool GetHelpNormalVisible(object control)
        {
            return !UpdateManager.HasNewVersion;
        }

        public bool GetHelpUpdateVisible(object control)
        {
            return UpdateManager.HasNewVersion;
        }

        public void OnRestartUpdate(object control)
        {
            UpdateManager.StartUpdateFlow();
        }

        public void OnBatchProcess(object control) { DataCommands.BatchProcess(); }
        public void OnExportCurrentSheet(object control) { ExportCommands.ExportCurrentSheet(); }
        public void OnExportStandardReport(object control) { ExportCommands.ExportStandardReport(); }
        public void OnExportInternalReport(object control) { ExportCommands.ExportInternalReport(); }
        public void OnConvertAllToValues(object control) { ExportCommands.ConvertAllToValues(); }
        public void OnShowHelp(object control) 
        { 
            LegacyAppCommands.ShowHelp(); 
        }
        public void OnNormalizeNumbers(object control) { DataCommands.NormalizeNumbers(); }
        public void OnTextify(object control) { DataCommands.Textify(); }
        public void OnLockFormula(object control) { LegacyAppCommands.LockFormula(); }
        public void OnWrapText(object control) { FormatCommands.WrapText(); }
        public void OnAccounting0(object control) { FormatCommands.Accounting0(); }
        public void OnAccounting2(object control) { FormatCommands.Accounting2(); }
        public void OnAccounting3(object control) { FormatCommands.Accounting3(); }
        public void OnYiWanFormat(object control) { FormatCommands.YiWanFormat(); }
        public void OnSetGrading(object control) { FormatCommands.SetGrading(); }
        public void OnSetGradingStyle(object control) { FormatCommands.SetGradingStyle(); }
        public void OnClearStyle(object control) { FormatCommands.ClearStyle(); }
        public void OnBreakLinks(object control) { FormatCommands.BreakExternalLinks(); }
        public void OnMultiAreaGroup(object control) { FormatCommands.MultiAreaGroup(); }
        public void OnMultiAreaUngroup(object control) { FormatCommands.MultiAreaUngroup(); }
        public void OnSelectVisibleCells(object control) { FormatCommands.SelectVisibleCells(); }
        public void OnResizePictures(object control) 
        { 
            int factor = 1;
            int f;
            var ribbonControl = control as Office.IRibbonControl;
            if (ribbonControl != null && ribbonControl.Tag != null && int.TryParse(ribbonControl.Tag, out f)) factor = f;
            PhotoCommands.ResizePictures(factor); 
        }
        public void OnSelectAllPictures(object control) { PhotoCommands.SelectAllPictures(); }
        public void OnCreateSheetIndex(object control) { SheetCommands.CreateSheetIndex(); }
        public void OnCreateFileIndex(object control) { SheetCommands.CreateFileIndex(); }
        public void OnMergeSheets(object control) { SheetCommands.MergeSheets(); }
        public void OnDeleteHyperlinks(object control) { SheetCommands.DeleteHyperlinks(); }
        public void OnForceRefresh(object control) { DataCommands.ForceRefresh(); }
        public void OnDeleteEmptyRows(object control) { SheetCommands.DeleteEmptyRows(); }
        public void OnUnhideAllSheets(object control) { SheetCommands.UnhideAllSheets(); }
        public void OnFileDirectory(object control) { SheetCommands.FileDirectory(); }

        public bool GetCalcAutoVisible(object control)
        {
            try { return App != null && App.Calculation == Excel.XlCalculation.xlCalculationAutomatic; }
            catch { return true; }
        }

        public bool GetCalcManualVisible(object control)
        {
            try { return App != null && App.Calculation != Excel.XlCalculation.xlCalculationAutomatic; }
            catch { return false; }
        }

        public void OnToggleCalculation(object control)
        {
            try
            {
                if (App == null) return;
                if (App.Calculation == Excel.XlCalculation.xlCalculationAutomatic)
                    App.Calculation = Excel.XlCalculation.xlCalculationManual;
                else
                    App.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                
                if (_ribbon != null) _ribbon.InvalidateControl("btnCalcAuto");
                if (_ribbon != null) _ribbon.InvalidateControl("btnCalcManual");
            }
            catch { }
        }

        #endregion
    }
}
