using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _application = (Excel.Application)Application;
                App = _application;
                // MessageBox.Show($"QSBar connecting to {(_application.Name)}...");
                
                // EPPlus License
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during QSBar connection: " + ex.Message);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            // Use 2009/07 namespace for better compatibility with modern Excel
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabQSBar' label='QSBar'>
        <group id='groupInvoice' label='Invoice Summary'>
          <button id='btnOtherFiles' label='Other Files' size='large' onAction='OnOtherFilesClick' imageMso='FileSaveAsExcelXlsx' />
          <button id='btnSameFiles' label='Same File' size='large' onAction='OnSameFilesClick' imageMso='TableInsert' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnOtherFilesClick(Office.IRibbonControl control)
        {
            if (_application != null)
            {
                otherFileForm form = new otherFileForm(_application.ActiveWorkbook);
                form.Show();
            }
        }

        public void OnSameFilesClick(Office.IRibbonControl control)
        {
            if (_application != null)
            {
                sameFileForm form = new sameFileForm(_application.ActiveWorkbook);
                form.Show();
            }
        }
    }
}
