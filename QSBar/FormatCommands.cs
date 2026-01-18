using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public static class FormatCommands
    {
        
        public static void WrapText()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Worksheet sheet = app.ActiveSheet as Excel.Worksheet;
            Excel.Range sel = app.Selection as Excel.Range;
            if (sheet == null || sel == null) return;

            Excel.Range used = sheet.UsedRange;
            Excel.Range target = app.Intersect(sel, used);
            if (target == null) return;

            target.NumberFormatLocal = "@";
            target.WrapText = true;
        }

        
        public static void Accounting0()
        {
            ApplyAccountingNumberFormatLocal(" #,##0_ ;[红色] -#,##0_ ;_ \"\"\"\"?_ ;@");
        }

        
        public static void Accounting2()
        {
            ApplyAccountingNumberFormatLocal("_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * \"-\"??_ ;_ @ ");
        }

        
        public static void Accounting3()
        {
            ApplyAccountingNumberFormatLocal("_ * #,##0.000_ ;_ * -#,##0.000_ ;_ * \"-\"???_ ;_ @ ");
        }

        static void ApplyAccountingNumberFormatLocal(string fmt)
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Worksheet sheet = app.ActiveSheet as Excel.Worksheet;
            Excel.Range sel = app.Selection as Excel.Range;
            if (sheet == null || sel == null) return;

            Excel.Range used = sheet.UsedRange;
            Excel.Range target = app.Intersect(sel, used);
            if (target == null) return;

            target.ShrinkToFit = true;
            target.NumberFormatLocal = fmt;
        }

        
        public static void YiWanFormat()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range sel = app.Selection as Excel.Range;
            if (sel == null) return;

            sel.ShrinkToFit = true;
            sel.NumberFormatLocal = @"[<=-100000000]-0!.00,,""亿元"";[>=100000000]0!.00,,""亿元"";0!.0,""万元""";
        }

        // =========================================================================================
        // M_Format.bas Migration
        // =========================================================================================

        public static void RowGroup(Excel.Range rng, int col, string cond, int lev)
        {
            if (rng == null) return;
            if (lev < 1) lev = 1;
            if (lev > 8) lev = 8; // Excel max outline level is 8

            int rowCount = rng.Rows.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                try
                {
                    string val = Convert.ToString(((Excel.Range)rng.Cells[i, col]).Value2);
                    if (Like(val, cond))
                    {
                        ((Excel.Range)rng.Rows[i]).OutlineLevel = lev;
                    }
                }
                catch { }
            }
        }

        
        public static void RowStyle()
        {
             Excel.Application app = WpsExcelAddIn.App;
             if (app == null) return;
             Excel.Range rng = app.Selection as Excel.Range;
             RowStyleInternal(rng);
        }

        private static void RowStyleInternal(Excel.Range rng)
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            bool originalUpdating = app.ScreenUpdating;
            Excel.XlCalculation originalCalc = app.Calculation;
            bool originalEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                if (rng == null) return;

                foreach (Excel.Range area in rng.Areas)
                {
                    int rowCount = area.Rows.Count;
                    if (rowCount <= 0) continue;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        Excel.Range row = area.Rows[i] as Excel.Range;
                        if (row == null) continue;

                        int level = 0;
                        try { level = (int)row.OutlineLevel; } catch { }
                        if (level == 0) 
                        {
                            try { level = (int)row.EntireRow.OutlineLevel; } catch { }
                        }

                        row.Font.Name = "Microsoft YaHei UI";

                        if (level == 1)
                        {
                            // 截图配色：深蓝灰色 (#333F4F) - 顶层大纲
                            row.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            row.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(51, 63, 79)); 
                            row.Interior.TintAndShade = 0;
                            row.Font.Color = ColorTranslator.ToOle(Color.White);
                            row.Font.TintAndShade = 0;
                            row.Font.Bold = true;
                        }
                        else if (level == 2)
                        {
                            // 截图配色：淡蓝色 (Excel 风格) - 二级分类
                            row.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            row.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(217, 225, 242));
                            row.Interior.TintAndShade = 0;
                            row.Font.Color = Color.Black.ToArgb(); // 黑色文字
                            row.Font.TintAndShade = 0;
                            row.Font.Bold = true;
                        }
                        else if (level == 3)
                        {
                            // 截图配色：淡橙色 (截图底部效果) - 三级明细
                            row.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            row.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(252, 228, 214));
                            row.Interior.TintAndShade = 0;
                            row.Font.Color = Color.Black.ToArgb();
                            row.Font.TintAndShade = 0;
                            row.Font.Bold = false;
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Optionally log error
            }
            finally
            {
                app.ScreenUpdating = originalUpdating;
                app.Calculation = originalCalc;
                app.EnableEvents = originalEvents;
            }
        }

        
        public static void ClearStyle()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            ClearStyleInternal(app.Selection as Excel.Range);
        }

        private static void ClearStyleInternal(Excel.Range rng)
        {
            if (rng == null) return;
            try
            {
                rng.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                rng.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                rng.Font.Bold = false;
            }
            catch { }
        }

        public static bool IsStyled(Excel.Range rng)
        {
             if (rng == null) return false;
             try
             {
                 foreach (Excel.Range cell in rng)
                 {
                     if ((int)cell.Font.ColorIndex != (int)Excel.Constants.xlAutomatic || (int)cell.Interior.ColorIndex != (int)Excel.Constants.xlNone) 
                     {
                         return true;
                     }
                 }
             }
             catch { }
             return false;
        }

        
        public static void SetGrading()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range selection = app.Selection as Excel.Range;
            Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
            if (selection == null || activeSheet == null) return;

            Excel.Range usedRange = activeSheet.UsedRange;
            Excel.Range rng = app.Intersect(selection, usedRange);
            if (rng == null) return;

            int col = 2; // Hardcoded in VBA
            
            try
            {
                foreach (Excel.Range row in rng.Rows)
                {
                    string v = Convert.ToString(((Excel.Range)row.Cells[1, col]).Value2);
                    
                    int level;
                    if ( Like(v, "*【*") ) level = 1;
                    else if ( Like(v, "*《*") || Like(v, "*<*")) level = 2;
                    else if (Like(v, "*{*") || Like(v, "*｛*") ) level = 3;
                    else level = 4;
                    
                    row.OutlineLevel = level;
                }
            }
            catch { }
        }

        
        public static void SetGradingStyle()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range rng = app.Selection as Excel.Range;
            if (rng == null) return;
            
            ClearStyleInternal(rng);
            RowStyleInternal(rng);
        }

        
        public static void MultiAreaGroup()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            try
            {
                Excel.Range selection = app.Selection as Excel.Range;
                if (selection == null) return;
                Excel.Range visible = selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                foreach (Excel.Range area in visible.Areas)
                {
                    foreach (Excel.Range row in area.Rows)
                    {
                        row.Group();
                    }
                }
            }
            catch { }
        }

        
        public static void MultiAreaUngroup()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            try
            {
                Excel.Range selection = app.Selection as Excel.Range;
                if (selection == null) return;
                Excel.Range visible = selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                foreach (Excel.Range area in visible.Areas)
                {
                    foreach (Excel.Range row in area.Rows)
                    {
                        row.Ungroup();
                    }
                }
            }
            catch { }
        }

        private static bool Like(string s, string pattern)
        {
            if (s == null) return false;
            // Escape regex characters
            string regexPattern = "^" + System.Text.RegularExpressions.Regex.Escape(pattern)
                .Replace("\\*", ".*")
                .Replace("\\?", ".") + "$";
            return System.Text.RegularExpressions.Regex.IsMatch(s, regexPattern, System.Text.RegularExpressions.RegexOptions.Singleline);
        }

        
        public static void SelectVisibleCells()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            try
            {
                Excel.Range selection = app.Selection as Excel.Range;
                if (selection == null) return;
                selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Select();
            }
            catch { }
        }
    }
}

