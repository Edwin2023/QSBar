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
                Excel.Range formulas = null;
                try
                {
                    formulas = selection.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                }
                catch
                {
                    return;
                }

                if (formulas != null)
                {
                    foreach (Excel.Range area in formulas.Areas)
                    {
                        foreach (Excel.Range cell in area)
                        {
                            try
                            {
                                string f = (string)cell.Formula;
                                if (!string.IsNullOrEmpty(f))
                                {
                                    object newFormula = app.ConvertFormula(
                                        f,
                                        Excel.XlReferenceStyle.xlA1,
                                        Excel.XlReferenceStyle.xlA1,
                                        Excel.XlReferenceType.xlAbsolute
                                    );

                                    string s = newFormula as string;
                                    if (!string.IsNullOrEmpty(s))
                                    {
                                        cell.Formula = s;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            catch
            {
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
                    MessageBox.Show("The current workbook has not been saved yet.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open folder: " + ex.Message);
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
                    ToastForm.ShowToast("Calculation mode switched to: Manual");
                }
                else
                {
                    app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    ToastForm.ShowToast("Calculation mode switched to: Automatic");
                }
                WpsExcelAddIn.InvalidateCalcButtons();
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
                        if (cell.Validation.Type != (int)Excel.XlDVType.xlValidateInputOnly)
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
