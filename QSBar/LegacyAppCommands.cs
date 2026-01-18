using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public static class LegacyAppCommands
    {
        
        public static void LockFormula()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range selection = app.Selection as Excel.Range;
            if (selection == null) return;

            app.ScreenUpdating = false;
            try
            {
                foreach (Excel.Range cell in selection)
                {
                    try
                    {
                        if (cell.HasFormula)
                        {
                            string formula = (string)cell.Formula;
                            if (formula.StartsWith("="))
                            {
                                string content = formula.Substring(1);
                                try
                                {
                                    Excel.Range target = app.Range[content];
                                    if (target != null)
                                    {
                                        string sheetName = target.Worksheet.Name;
                                        string address = target.Address;
                                        cell.Formula = "='" + sheetName + "'!" + address;
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                    catch
                    {
                        // Ignore errors
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        
        public static void OpenFileDirectory()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            try
            {
                string path = app.ActiveWorkbook.Path;
                if (!string.IsNullOrEmpty(path))
                {
                    System.Diagnostics.Process.Start("explorer.exe", path);
                }
                else
                {
                    MessageBox.Show("当前工作簿尚未保存！");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开目录失败: " + ex.Message);
            }
        }

        public static void ShowHelp()
        {
            MessageBox.Show("QS工具箱帮助信息:\n\n1. 快捷键:\n   Ctrl+5: 选择可见/全选\n   Ctrl+6: 亿元/万元格式\n   Ctrl+7: 文本换行\n   Ctrl+8: 会计格式0位\n   Ctrl+9: 会计格式2位\n   Ctrl+0: 会计格式3位\n\n更多功能请查看Ribbon工具栏。", "帮助");
        }

        public static void ToggleCalculation()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            try
            {
                if (app.Calculation == Excel.XlCalculation.xlCalculationAutomatic)
                {
                    app.Calculation = Excel.XlCalculation.xlCalculationManual;
                }
                else
                {
                    app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                }
                // QSRibbon.InvalidateCalcButtons();
            }
            catch { }
        }

        
        public static void LockValidationCells()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range selection = app.Selection as Excel.Range;
            if (selection == null) return;

            try
            {
                foreach (Excel.Range cell in selection)
                {
                    try
                    {
                        if (cell.Validation.Type != (int)Excel.XlDVType.xlValidateInputOnly) // xlValidateInputOnly = 0
                        {
                            cell.Locked = true;
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}

