using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading.Tasks;
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

        public static void RefreshRibbon()
        {
            _ribbon?.Invalidate();
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _application = (Excel.Application)Application;
                App = _application;
                
                // EPPlus License
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                StartCalcTimer();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during QSBar connection: " + ex.Message);
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
                            _ribbon?.InvalidateControl("btnCalcAuto");
                            _ribbon?.InvalidateControl("btnCalcManual");
                        }
                        catch { }
                    };
                    _calcTimer.Start();
                }
            }
            catch { }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _calcTimer?.Stop();
            _calcTimer = null;
            _application = null;
            App = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public void OnLoad(Office.IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            
            // 启动时静默检查更新（不打扰用户）
            Task.Run(() => UpdateManager.CheckForUpdateAsync(true));
        }

        public string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='tabQSBar' label='QS工具箱'>
        <group id='groupFormat' label='格式/排版'>
          <button id='btnNormalize' label='数值化' onAction='OnNormalizeNumbers' imageMso='NumberFormat' size='large' screentip='数据规范化' supertip='对选中区域的数据进行规范化处理' />
          <button id='btnLock' label='锁定公式' onAction='OnLockFormula' imageMso='Lock' size='large' screentip='锁定/解锁公式' supertip='锁定或解锁选中区域的单元格公式' />
          <splitButton id='spFormat' size='large'>
            <button id='btnFormatDefault' label='单元格式' onAction='OnWrapText' imageMso='WrapText' screentip='单元格式' supertip='设置单元格格式' />
            <menu id='menuFormat' itemSize='large'>
              <button id='btnWrapText' label='文本格式 (Ctrl+7)' onAction='OnWrapText' imageMso='WrapText' />
              <button id='btnAcct0' label='会计格式0位 (Ctrl+8)' onAction='OnAccounting0' imageMso='CommaStyle' />
              <button id='btnAcct2' label='会计格式2位 (Ctrl+9)' onAction='OnAccounting2' imageMso='CommaStyle' />
              <button id='btnAcct3' label='会计格式3位 (Ctrl+0)' onAction='OnAccounting3' imageMso='CommaStyle' />
              <button id='btnYiWan' label='亿/万位格式 (Ctrl+6)' onAction='OnYiWanFormat' imageMso='CommaStyle' />
            </menu>
          </splitButton>
        </group>

        <group id='groupBQ' label='批量处理'>
          <button id='btnSetGrading' label='设置分级' onAction='OnSetGrading' imageMso='ObjectsGroup' size='large' screentip='设置分级显示' supertip='根据内容自动设置工作表的分级显示' />
          <button id='btnSetGradingStyle' label='分级样式' onAction='OnSetGradingStyle' imageMso='FormatPainter' size='large' screentip='设置分级样式' supertip='为分级显示设置不同的单元格样式' />
        </group>

        <group id='groupGrading' label='分级/筛选'>
          <button id='btnBatch' label='批量处理' onAction='OnBatchProcess' imageMso='ViewSheetGridlines' size='large' screentip='批量处理/公式计算' supertip='批量处理或公式计算，通常用于大数据量处理' />
          <button id='btnClearStyle' label='清除样式' onAction='OnClearStyle' imageMso='Clear' size='large' screentip='清除样式' supertip='清除选中区域的所有单元格样式' />
          <splitButton id='spExport' size='large'>
            <button id='btnExportDefault' label='导出报表' onAction='OnExportCurrentSheet' imageMso='WindowNew' screentip='导出报表' supertip='以数值形式导出报表数据' />
            <menu id='menuExport' itemSize='large'>
              <button id='btnOutSheet' label='导出当前表' onAction='OnExportCurrentSheet' imageMso='WindowNew' />
              <button id='btnOutStd' label='导出标准报表' onAction='OnExportStandardReport' imageMso='FileSaveAs' />
              <button id='btnOutInt' label='导出内部报表' onAction='OnExportInternalReport' imageMso='FileSaveAs' />
            </menu>
          </splitButton>
        </group>

        <group id='groupPhotoTools' label='图片'>
          <splitButton id='spPhotoResize' size='large'>
            <button id='btnPhotoResizeDefault' label='调整图片大小' onAction='OnResizePictures' tag='1' imageMso='ObjectsGroup' screentip='调整图片大小' supertip='批量调整工作表中图片的大小' />
            <menu id='menuPhotoResize' itemSize='large'>
              <button id='btnResize1' label='1倍大小' onAction='OnResizePictures' tag='1' imageMso='PictureInsertFromFile' />
              <button id='btnResize2' label='2倍大小' onAction='OnResizePictures' tag='2' imageMso='PictureInsertFromFile' />
              <button id='btnResize3' label='3倍大小' onAction='OnResizePictures' tag='3' imageMso='PictureInsertFromFile' />
              <button id='btnResize4' label='4倍大小' onAction='OnResizePictures' tag='4' imageMso='PictureInsertFromFile' />
            </menu>
          </splitButton>
          <button id='btnSelectAllPictures' label='全选图片' onAction='OnSelectAllPictures' imageMso='SelectAll' size='large' screentip='选中全图' supertip='选中当前工作表中的所有图片' />
          <button id='btnSelVisible' label='选中可见单元格' onAction='OnSelectVisibleCells' imageMso='SelectAll' size='large' screentip='选中可见单元格' supertip='选中选择区域的可见单元格 (Ctrl+5)' />/>
        </group>

        <group id='groupSheets' label='工作表/文件'>
          <button id='btnMergeSheets' label='合并工作表' onAction='OnMergeSheets' imageMso='Consolidate' size='large' screentip='合并工作表' supertip='将多个工作表合并为一个工作表' />
          <button id='btnSheetIndex' label='生成表目录' onAction='OnCreateSheetIndex' imageMso='Numbering' size='large' screentip='生成表目录' supertip='在当前工作簿中生成所有工作表的目录' />
          <button id='btnForceRefresh' label='强制刷新' onAction='OnForceRefresh' imageMso='Refresh' size='normal' screentip='强制刷新' supertip='强制刷新选中区域的所有公式' />
          <button id='btnValOnly' label='全表粘死' onAction='OnConvertAllToValues' imageMso='PasteValues' size='normal' screentip='全表粘死' supertip='将当前工作表的所有公式转换为数值' />
          <button id='btnDeleteLinks' label='删除超链接' onAction='OnDeleteHyperlinks' imageMso='Delete' size='normal' screentip='删除超链接' supertip='删除选中区域的所有超链接' />
          <button id='btnDeleteEmptyRows' label='删除空行' onAction='OnDeleteEmptyRows' imageMso='TableRowsDelete' size='normal' screentip='删除空行' supertip='删除选中区域内的所有空行' />
          <button id='btnUnhideSheets' label='取消隐藏所有表' onAction='OnUnhideAllSheets' imageMso='FillRight' size='normal' screentip='取消隐藏所有表' supertip='取消隐藏当前工作簿中的所有工作表' />
          <button id='btnFileDir' label='文件目录' onAction='OnFileDirectory' imageMso='FileOpen' size='normal' screentip='文件目录' supertip='查看或管理当前文件所在的目录' />
        </group>

        <group id='groupCalc' label='计算模式'>
          <button id='btnCalcAuto' label='切换自动计算' getVisible='GetCalcAutoVisible' onAction='OnToggleCalculation' imageMso='CalculateNow' size='large' screentip='状态: 自动' supertip='将 Excel 计算模式切换为自动' />
          <button id='btnCalcManual' label='切换手动计算' getVisible='GetCalcManualVisible' onAction='OnToggleCalculation' imageMso='CalculateFull' size='large' screentip='状态: 手动' supertip='将 Excel 计算模式切换为手动' />
        </group>

        <group id='groupHelpUpdate' label='帮助更新'>
          <button id='btnHelp' label='使用帮助' onAction='OnShowHelp' imageMso='Help' size='large' screentip='使用帮助' supertip='查看 QS 工具箱的版本信息、快捷键及检查更新' />
          <button id='btnUpdate' label='重启更新' getVisible='GetHelpUpdateVisible' onAction='OnRestartUpdate' imageMso='Refresh' size='large' screentip='重启更新' supertip='检测到新版本，请点击并确认重启以完成更新' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
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
        public void OnLockFormula(Office.IRibbonControl control) { LegacyAppCommands.LockFormula(); }
        public void OnWrapText(Office.IRibbonControl control) { FormatCommands.WrapText(); }
        public void OnAccounting0(Office.IRibbonControl control) { FormatCommands.Accounting0(); }
        public void OnAccounting2(Office.IRibbonControl control) { FormatCommands.Accounting2(); }
        public void OnAccounting3(Office.IRibbonControl control) { FormatCommands.Accounting3(); }
        public void OnYiWanFormat(Office.IRibbonControl control) { FormatCommands.YiWanFormat(); }
        public void OnSetGrading(Office.IRibbonControl control) { FormatCommands.SetGrading(); }
        public void OnSetGradingStyle(Office.IRibbonControl control) { FormatCommands.SetGradingStyle(); }
        public void OnClearStyle(Office.IRibbonControl control) { FormatCommands.ClearStyle(); }
        public void OnMultiAreaGroup(Office.IRibbonControl control) { FormatCommands.MultiAreaGroup(); }
        public void OnMultiAreaUngroup(Office.IRibbonControl control) { FormatCommands.MultiAreaUngroup(); }
        public void OnSelectVisibleCells(Office.IRibbonControl control) { FormatCommands.SelectVisibleCells(); }
        public void OnResizePictures(Office.IRibbonControl control) 
        { 
            int factor = 1;
            if (control.Tag != null && int.TryParse(control.Tag, out int f)) factor = f;
            PhotoCommands.ResizePictures(factor); 
        }
        public void OnSelectAllPictures(Office.IRibbonControl control) { PhotoCommands.SelectAllPictures(); }
        public void OnCreateSheetIndex(Office.IRibbonControl control) { SheetCommands.CreateSheetIndex(); }
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
            try
            {
                if (App == null) return;
                if (App.Calculation == Excel.XlCalculation.xlCalculationAutomatic)
                    App.Calculation = Excel.XlCalculation.xlCalculationManual;
                else
                    App.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                
                _ribbon?.InvalidateControl("btnCalcAuto");
                _ribbon?.InvalidateControl("btnCalcManual");
            }
            catch { }
        }

        #endregion
    }
}
