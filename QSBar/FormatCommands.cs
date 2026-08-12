using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
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

        private static void ApplyAccountingNumberFormatLocal(string fmt)
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

                // 全局一次性设置字体，极大减少 COM 调用
                try { rng.Font.Name = "Microsoft YaHei UI"; } catch { }

                foreach (Excel.Range area in rng.Areas)
                {
                    int rowCount = area.Rows.Count;
                    if (rowCount <= 0) continue;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        try
                        {
                            Excel.Range row = (Excel.Range)area.Rows[i];
                            int level = (int)row.OutlineLevel;
                            ApplyStyleToRow(row, level);
                        }
                        catch { }
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

        private static void ApplyStyleToRow(Excel.Range row, int level)
        {
            try
            {
                if (row == null) return;

                if (level == 1)
                {
                    row.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    row.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(51, 63, 79)); // #333F4F 深蓝灰
                    row.Interior.TintAndShade = 0;
                    row.Font.Color = ColorTranslator.ToOle(Color.FromArgb(255, 255, 255)); // #FFFFFF 白色
                    row.Font.TintAndShade = 0;
                    row.Font.Bold = true;
                    row.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
                else if (level == 2)
                {
                    row.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    row.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(217, 225, 242)); // #D9E1F2 浅蓝
                    row.Interior.TintAndShade = 0;
                    row.Font.Color = ColorTranslator.ToOle(Color.FromArgb(26, 26, 26)); // #1A1A1A
                    row.Font.TintAndShade = 0;
                    row.Font.Bold = true;
                    row.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
                else if (level == 3)
                {
                    row.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    row.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(251, 229, 214)); // #FBE5D6
                    row.Interior.TintAndShade = 0;
                    row.Font.Color = ColorTranslator.ToOle(Color.FromArgb(26, 26, 26)); // #1A1A1A
                    row.Font.TintAndShade = 0;
                    row.Font.Bold = true;
                }
            }
            catch { }
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
                int colCount = rng.Columns.Count;
                for (int i = 1; i <= rowCount; i++)
                {
                    int level = 4;
                    // 遍历该行的所有选中列，寻找分级标志
                    for (int c = 1; c <= colCount; c++)
                    {
                        object val = dataValues[i, c];
                        string v = val == null ? "" : val.ToString();
                        
                        if (Like(v, "*【*") || Like(v, "*[[*")) { level = 1; break; }
                        else if (Like(v, "*《*") || Like(v, "*<<*")) { level = 2; break; }
                        else if (Like(v, "*{*") || Like(v, "*｛*")) { level = 3; break; }
                    }
                    
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
            Excel.Range selection = app.Selection as Excel.Range;
            Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
            if (selection == null || activeSheet == null) return;

            Excel.Range rng = app.Intersect(selection, activeSheet.UsedRange);
            if (rng == null) return;

            SmartClearAndRestyle(app, rng);
        }

        private struct ManualCellFormat
        {
            public int Col;
            public int InteriorColor;
            public int Pattern;
            public int FontColor;
            public bool FontBold;
            public double FontSize;
            public string FontName;
        }

        private static void SmartClearAndRestyle(Excel.Application app, Excel.Range rng)
        {
            if (rng == null) return;

            bool originalUpdating = app.ScreenUpdating;
            Excel.XlCalculation originalCalc = app.Calculation;
            bool originalEvents = app.EnableEvents;

            int colorL1Bg = ColorTranslator.ToOle(Color.FromArgb(51, 63, 79));
            int colorL2Bg = ColorTranslator.ToOle(Color.FromArgb(217, 225, 242));
            int colorL3Bg = ColorTranslator.ToOle(Color.FromArgb(251, 229, 214));
            int colorL1Font = ColorTranslator.ToOle(Color.FromArgb(255, 255, 255));
            int colorL23Font = ColorTranslator.ToOle(Color.FromArgb(26, 26, 26));

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                foreach (Excel.Range area in rng.Areas)
                {
                    int rowCount = area.Rows.Count;
                    if (rowCount <= 0) continue;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        Excel.Range row = null;
                        try
                        {
                            row = (Excel.Range)area.Rows[i];
                            int level = (int)row.OutlineLevel;
                            if (level < 1 || level > 3) continue;

                            int gradingBg = level == 1 ? colorL1Bg : (level == 2 ? colorL2Bg : colorL3Bg);
                            int gradingFont = level == 1 ? colorL1Font : colorL23Font;

                            var manualCells = new List<ManualCellFormat>();
                            int colCount = row.Columns.Count;

                            for (int c = 1; c <= colCount; c++)
                            {
                                try
                                {
                                    Excel.Range cell = (Excel.Range)row.Cells[1, c];
                                    int cellBg = (int)cell.Interior.Color;
                                    int cellPattern = (int)cell.Interior.Pattern;

                                    if (cellPattern == (int)Excel.XlPattern.xlPatternSolid &&
                                        cellBg == gradingBg &&
                                        (int)cell.Font.Color == gradingFont &&
                                        (bool)cell.Font.Bold)
                                        continue;
                                    if (cellPattern == (int)Excel.XlPattern.xlPatternNone)
                                        continue;

                                    manualCells.Add(new ManualCellFormat
                                    {
                                        Col = c,
                                        InteriorColor = cellBg,
                                        Pattern = cellPattern,
                                        FontColor = (int)cell.Font.Color,
                                        FontBold = (bool)cell.Font.Bold,
                                        FontSize = (double)cell.Font.Size,
                                        FontName = cell.Font.Name as string
                                    });
                                }
                                catch { }
                            }

                            ClearStyleInternal(row);
                            ApplyStyleToRow(row, level);

                            foreach (var mc in manualCells)
                            {
                                try
                                {
                                    Excel.Range cell = (Excel.Range)row.Cells[1, mc.Col];
                                    cell.Interior.Pattern = (Excel.XlPattern)mc.Pattern;
                                    cell.Interior.Color = mc.InteriorColor;
                                    cell.Font.Color = mc.FontColor;
                                    cell.Font.Bold = mc.FontBold;
                                    cell.Font.Size = mc.FontSize;
                                    cell.Font.Name = mc.FontName;
                                }
                                catch { }
                            }
                        }
                        catch { }
                        finally
                        {
                            if (row != null) Marshal.ReleaseComObject(row);
                        }
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = originalUpdating;
                app.Calculation = originalCalc;
                app.EnableEvents = originalEvents;
            }
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

        public static void SelectNonEmptyCells()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            try
            {
                Excel.Range selection = app.Selection as Excel.Range;
                if (selection == null) return;
                
                // SpecialCells can only find one type at a time.
                // We need to combine constants (xlCellTypeConstants = 2) and formulas (xlCellTypeFormulas = -4123)
                Excel.Range constants = null;
                Excel.Range formulas = null;
                
                try { constants = selection.SpecialCells(Excel.XlCellType.xlCellTypeConstants); } catch { }
                try { formulas = selection.SpecialCells(Excel.XlCellType.xlCellTypeFormulas); } catch { }
                
                if (constants != null && formulas != null)
                {
                    app.Union(constants, formulas).Select();
                }
                else if (constants != null)
                {
                    constants.Select();
                }
                else if (formulas != null)
                {
                    formulas.Select();
                }
            }
            catch { }
        }

        private struct CleanStats
        {
            public int ExtLinksFound;
            public int CellsConverted;
            public int NamesDeleted;
            public int LinksBroken;
            public int Errors;
        }

        private static readonly string[] ErrTokens = { "#REF!", "#VALUE!", "#N/A", "#NAME?", "#DIV/0!", "#NULL!", "#NUM!" };

        public static void BreakExternalLinks()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            var stats = new CleanStats();
            var sw = System.Diagnostics.Stopwatch.StartNew();

            bool oldScreenUpdating = app.ScreenUpdating;
            Excel.XlCalculation oldCalc = app.Calculation;
            bool oldEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // Step 1: Scan external link sources and broken names
                app.StatusBar = "Break Links: Scanning external links...";
                System.Windows.Forms.Application.DoEvents();

                ScanExternalLinks(workbook, ref stats);

                if (stats.ExtLinksFound == 0)
                {
                    app.StatusBar = false;
                    MessageBox.Show("No external links detected. Cleanup is not needed.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Step 2: Convert formulas with external references to values
                app.StatusBar = "Break Links: Converting formulas with external references to values...";
                System.Windows.Forms.Application.DoEvents();
                ConvertExternalFormulas(workbook, app, ref stats);

                // Step 3: Delete broken defined names
                app.StatusBar = "Break Links: Deleting broken names...";
                System.Windows.Forms.Application.DoEvents();
                DeleteBadNames(workbook, app, ref stats);

                // Step 4: BreakLink all external link sources
                app.StatusBar = "Break Links: Disconnecting link sources...";
                System.Windows.Forms.Application.DoEvents();
                BreakAllLinks(workbook, app, ref stats);

                // Step 5: Clean hidden names
                app.StatusBar = "Break Links: Cleaning hidden names...";
                System.Windows.Forms.Application.DoEvents();
                CleanHiddenNames(workbook, app, ref stats);

                app.StatusBar = false;
                sw.Stop();

                MessageBox.Show(
                    "All external links have been broken into values, and invalid defined names have been cleaned up.\n" +
                    "---\n" +
                    "External link sources / broken names detected: " + stats.ExtLinksFound + "\n" +
                    "Formulas converted to values: " + stats.CellsConverted + "\n" +
                    "Defined names deleted: " + stats.NamesDeleted + "\n" +
                    "Links broken: " + stats.LinksBroken + "\n" +
                    "Errors: " + stats.Errors + "\n" +
                    "Time: " + sw.Elapsed.TotalSeconds.ToString("0.00") + " seconds\n" +
                    "\nSave the workbook and reopen it. The \"update links\" prompt should no longer appear.",
                    "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                app.StatusBar = false;
                MessageBox.Show("Error while removing external links: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.EnableEvents = oldEvents;
                app.Calculation = oldCalc;
                app.ScreenUpdating = oldScreenUpdating;
                app.StatusBar = false;
            }
        }

        private static void ScanExternalLinks(Excel.Workbook wb, ref CleanStats stats)
        {
            Array links = wb.LinkSources(Excel.XlLink.xlExcelLinks) as Array;
            if (links != null)
                stats.ExtLinksFound += links.Length;

            foreach (Excel.Name nm in wb.Names)
            {
                try
                {
                    if (IsBadNameRefersTo(nm.RefersTo as string))
                        stats.ExtLinksFound++;
                }
                catch { }
            }
        }

        private static void ConvertExternalFormulas(Excel.Workbook wb, Excel.Application app, ref CleanStats stats)
        {
            const int doEventsInterval = 200;

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (ws.Type != Excel.XlSheetType.xlWorksheet) continue;
                Excel.Range usedRange = null;
                try { usedRange = ws.UsedRange; } catch { continue; }
                if (usedRange == null) continue;

                Excel.Range formulaCells = null;
                try { formulaCells = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas); } catch { }
                if (formulaCells == null) continue;

                app.StatusBar = "删除外部链接: 处理工作表 " + ws.Name + " ...";
                System.Windows.Forms.Application.DoEvents();

                foreach (Excel.Range area in formulaCells.Areas)
                {
                    int rowCount = area.Rows.Count;
                    int colCount = area.Columns.Count;

                    if (rowCount == 1 && colCount == 1)
                    {
                        string f = "";
                        try { f = area.Formula as string; } catch { }
                        if (HasExternalRef(f))
                        {
                            try { area.Value = area.Value; stats.CellsConverted++; }
                            catch { stats.Errors++; }
                        }
                        continue;
                    }

                    object[,] formulas = null;
                    try { formulas = area.Formula as object[,]; } catch { }
                    if (formulas == null) continue;

                    bool hasExt = false;
                    int lbR = formulas.GetLowerBound(0), ubR = formulas.GetUpperBound(0);
                    int lbC = formulas.GetLowerBound(1), ubC = formulas.GetUpperBound(1);

                    for (int r = lbR; r <= ubR && !hasExt; r++)
                        for (int c = lbC; c <= ubC && !hasExt; c++)
                            if (HasExternalRef(formulas[r, c] as string))
                                hasExt = true;

                    if (!hasExt) continue;

                    object[,] values = null;
                    try { values = area.Value as object[,]; } catch { }
                    if (values == null) continue;

                    int converted = 0;
                    for (int r = lbR; r <= ubR; r++)
                    {
                        for (int c = lbC; c <= ubC; c++)
                        {
                            if (HasExternalRef(formulas[r, c] as string))
                            {
                                formulas[r, c] = values[r, c];
                                converted++;
                            }
                        }
                        if (r % doEventsInterval == 0)
                            System.Windows.Forms.Application.DoEvents();
                    }

                    if (converted > 0)
                    {
                        area.Formula = formulas;
                        stats.CellsConverted += converted;
                    }
                }
            }
        }

        private static void DeleteBadNames(Excel.Workbook wb, Excel.Application app, ref CleanStats stats)
        {
            var badNames = new System.Collections.Generic.List<string>();

            foreach (Excel.Name nm in wb.Names)
            {
                try
                {
                    if (IsBadNameRefersTo(nm.RefersTo as string))
                        badNames.Add(nm.Name);
                }
                catch { }
            }

            for (int i = 0; i < badNames.Count; i++)
            {
                if (i % 50 == 0)
                {
                    app.StatusBar = "删除外部链接: 删除损坏名称 " + (i + 1) + "/" + badNames.Count + " ...";
                    System.Windows.Forms.Application.DoEvents();
                }

                try { wb.Names.Item(badNames[i]).Delete(); stats.NamesDeleted++; }
                catch { stats.Errors++; }
            }
        }

        private static void BreakAllLinks(Excel.Workbook wb, Excel.Application app, ref CleanStats stats)
        {
            Array links = wb.LinkSources(Excel.XlLink.xlExcelLinks) as Array;
            if (links == null) return;

            for (int i = links.GetLowerBound(0); i <= links.GetUpperBound(0); i++)
            {
                app.StatusBar = "删除外部链接: 断开链接 " + (i + 1) + "/" + links.Length;
                System.Windows.Forms.Application.DoEvents();

                try
                {
                    object link = links.GetValue(i);
                    if (link != null)
                    {
                        wb.BreakLink(link.ToString(), Excel.XlLinkType.xlLinkTypeExcelLinks);
                        stats.LinksBroken++;
                    }
                }
                catch { stats.Errors++; }
            }
        }

        private static void CleanHiddenNames(Excel.Workbook wb, Excel.Application app, ref CleanStats stats)
        {
            var badNames = new System.Collections.Generic.List<string>();

            foreach (Excel.Name nm in wb.Names)
            {
                try
                {
                    string n = nm.Name;
                    if (n.Contains("!")) continue;
                    if (n.StartsWith("_xlfn")) continue;
                    if (n.StartsWith("_xlpm")) continue;
                    if (n.StartsWith("_FilterDatabase")) continue;

                    if (IsBadNameRefersTo(nm.RefersTo as string))
                        badNames.Add(n);
                }
                catch { }
            }

            for (int i = 0; i < badNames.Count; i++)
            {
                if (i % 50 == 0)
                {
                    app.StatusBar = "删除外部链接: 清理隐藏名称 " + (i + 1) + "/" + badNames.Count + " ...";
                    System.Windows.Forms.Application.DoEvents();
                }

                try { wb.Names.Item(badNames[i]).Delete(); stats.NamesDeleted++; }
                catch { }
            }
        }

        private static bool HasExternalRef(string formulaText)
        {
            if (string.IsNullOrEmpty(formulaText)) return false;

            if (formulaText.Contains("[")) return true;

            foreach (string t in ErrTokens)
                if (formulaText.Contains(t)) return true;

            if (formulaText.StartsWith("file://")) return true;

            if (formulaText.StartsWith("\\\\") && formulaText.Length > 2 && char.IsLetter(formulaText[2]))
                return true;

            return false;
        }

        private static bool IsBadNameRefersTo(string refersTo)
        {
            if (string.IsNullOrEmpty(refersTo)) return true;

            foreach (string t in ErrTokens)
                if (refersTo.Contains(t)) return true;

            if (refersTo.Contains("[")) return true;
            if (refersTo.StartsWith("file://")) return true;

            if (refersTo.StartsWith("\\\\") && refersTo.Length > 2 && char.IsLetter(refersTo[2]))
                return true;

            if (refersTo.Length >= 3 && refersTo[1] == ':' && refersTo[2] == '\\')
                return true;

            return false;
        }
    }
}

