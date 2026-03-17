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
                // Find all cells with formulas in the selection
                Excel.Range formulas = null;
                try
                {
                    formulas = selection.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                }
                catch 
                {
                    // No formulas found
                    return;
                }

                if (formulas != null)
                {
                    foreach (Excel.Range area in formulas.Areas)
                    {
                        // Convert formulas to absolute references
                        foreach (Excel.Range cell in area)
                        {
                            try
                            {
                                string f = (string)cell.Formula;
                                if (!string.IsNullOrEmpty(f))
                                {
                                    // Convert A1 style references to absolute A1 style references
                                    object newFormula = app.ConvertFormula(
                                        f,
                                        Excel.XlReferenceStyle.xlA1,
                                        Excel.XlReferenceStyle.xlA1,
                                        Excel.XlReferenceType.xlAbsolute
                                    );
                                    
                                    if (newFormula is string)
                                    {
                                        string s = (string)newFormula;
                                        if (!string.IsNullOrEmpty(s))
                                        {
                                            cell.Formula = s;
                                        }
                                    }
                                }
                            }
                            catch { /* Ignore individual conversion errors */ }
                        }
                    }
                }
            }
            catch
            {
                // Ignore general errors
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
            using (var form = new HelpForm())
            {
                form.ShowDialog();
            }
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

