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
            
            try
            {
                Excel.Range selection = app.Selection as Excel.Range;
                Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
                if (selection == null || activeSheet == null) return;

                // 1) 确定范围 - Intersect 确保只处理有数据的区域
                Excel.Range rng = app.Intersect(selection, activeSheet.UsedRange);
                if (rng == null) return;

                // 2) 智能判断列索引
                // 如果只选了 1 列，或者选区是从 B 列开始的，则判断选区内的第 1 列
                // 否则默认判断选区内的第 2 列
                int colInRng = 2;
                if (selection.Columns.Count == 1 || selection.Column == 2)
                {
                    colInRng = 1;
                }
                
                // 性能优化：一次性将整个区域的数据读入内存数组
                // 这比循环数千次读取单元格要快 100 倍以上
                object[,] dataValues = null;
                if (rng.Rows.Count > 1 || rng.Columns.Count > 1)
                {
                    dataValues = rng.Value2 as object[,];
                }
                else
                {
                    // 单个单元格处理
                    dataValues = new object[2, 2];
                    dataValues[1, 1] = rng.Value2;
                }

                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual; // 暂时关闭自动计算

                // 3) 循环数组并设置等级
                int rowCount = rng.Rows.Count;
                for (int i = 1; i <= rowCount; i++)
                {
                    // 从数组中取值，无需访问 Excel 接口
                    object val = dataValues[i, colInRng];
                    string v = val == null ? "" : val.ToString();
                    
                    int level = 4;
                    if (Like(v, "*【*")) level = 1;
                    else if (Like(v, "*《*")) level = 2;
                    else if (Like(v, "*{*") || Like(v, "*｛*")) level = 3;
                    
                    // 只有在等级不同时才设置，进一步优化
                    Excel.Range row = (Excel.Range)rng.Rows[i];
                    if ((int)row.OutlineLevel != level)
                    {
                        row.OutlineLevel = level;
                    }
                }
            }
            catch (Exception)
            {
                // 可以按需记录异常
            }
            finally
            {
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                app.ScreenUpdating = true;
            }
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

        public static void BreakExternalLinks()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            try
            {
                app.ScreenUpdating = false;

                // 1. 断开外部工作簿链接 (将公式转为数值)
                object linksObj = workbook.LinkSources(Excel.XlLink.xlExcelLinks);
                Array links = linksObj as Array;
                if (links != null)
                {
                    for (int i = links.GetLowerBound(0); i <= links.GetUpperBound(0); i++)
                    {
                        try
                        {
                            object link = links.GetValue(i);
                            if (link != null)
                            {
                                workbook.BreakLink(link.ToString(), Excel.XlLinkType.xlLinkTypeExcelLinks);
                            }
                        }
                        catch { }
                    }
                }

                // 2. 清理定义名称 (Names)
                Excel.Names names = workbook.Names;
                for (int i = names.Count; i >= 1; i--)
                {
                    Excel.Name name = names.Item(i);
                    string refersTo = "";
                    try { refersTo = name.RefersTo; } catch { }

                    // 删除原则：
                    // 1) 包含报错 #REF!
                    // 2) 包含外部工作簿引用标记 [ 
                    // 3) 包含绝对路径标识 (:\ 或 \\) 且有工作表关联 !
                    if (refersTo.Contains("#REF!") || 
                        refersTo.Contains("[") || 
                        (refersTo.Contains("!") && (refersTo.Contains(":\\") || refersTo.Contains("\\\\"))))
                    {
                        try { name.Delete(); } catch { }
                    }
                }

                MessageBox.Show("外部链接已全部断开为数值，且已清理异常的定义名称。", "处理完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("删除外部链接时出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }
    }
}

