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

            // 逐格改写公式，自动计算开着的话每改一格就重算一遍
            Excel.Range formulas = null;
            Excel.Areas areas = null;

            using (ExcelScope.Begin(app))
            try
            {
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
                    areas = formulas.Areas;
                    int areaCount = areas.Count;
                    for (int a = 1; a <= areaCount; a++)
                    {
                        Excel.Range area = null;
                        try
                        {
                            area = areas[a];
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
                                finally { ComUtil.Release(cell); }
                            }
                        }
                        finally { ComUtil.Release(area); }
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ComUtil.Release(areas, formulas, selection);
            }
        }

        public static void OpenFileDirectory()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Workbook wb = null;
            try
            {
                wb = app.ActiveWorkbook;
                if (wb == null) return;

                string path = wb.Path;
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
            finally { ComUtil.Release(wb); }
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
                    Excel.Validation validation = null;
                    try
                    {
                        validation = cell.Validation;
                        if (validation.Type != (int)Excel.XlDVType.xlValidateInputOnly)
                        {
                            cell.Locked = true;
                        }
                    }
                    catch { }
                    finally { ComUtil.Release(validation, cell); }
                }
            }
            catch { }
            finally { ComUtil.Release(selection); }
        }
    }
}
