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

            Excel.Worksheet sheet = null;
            Excel.Range sel = null;
            Excel.Range used = null;
            Excel.Range target = null;
            try
            {
                sheet = app.ActiveSheet as Excel.Worksheet;
                sel = app.Selection as Excel.Range;
                if (sheet == null || sel == null) return;

                used = sheet.UsedRange;
                target = app.Intersect(sel, used);
                if (target == null) return;

                target.NumberFormatLocal = "@";
                target.WrapText = true;
            }
            finally { ComUtil.Release(target, used, sel, sheet); }
        }

        
        // Ctrl+0 三档轮换：2位 → 0位 → 3位 → 回到2位。
        // 三档风格统一（红字负数、零值空白、_ * 填充对齐），轮换时只有小数位数在变。
        // [红色] 必须写在段首：写成 _ * [红色]-#,##0.00_ 的话 Excel 会静默把它挪到段首，
        // 回读就跟常量不相等，轮换会卡在第一档
        private const string AcctFormat0 = "_ * #,##0_ ;[红色]_ * -#,##0_ ;_ * \"\"_ ;_ @ ";
        private const string AcctFormat2 = "_ * #,##0.00_ ;[红色]_ * -#,##0.00_ ;_ * \"\"??_ ;_ @ ";
        private const string AcctFormat3 = "_ * #,##0.000_ ;[红色]_ * -#,##0.000_ ;_ * \"\"???_ ;_ @ ";

        // Ctrl+8 百分号格式。风格跟会计格式一条线：红字负数、零值空白、_ * 填充对齐
        private const string PercentFmt = "_ * #,##0.00%_ ;[红色]_ * -#,##0.00%_ ;_ * \"\"??_ ;_ @ ";

        // 调用方负责 Release 返回的 Range
        private static Excel.Range GetUsedSelection()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return null;

            Excel.Worksheet sheet = null;
            Excel.Range sel = null;
            Excel.Range used = null;
            try
            {
                sheet = app.ActiveSheet as Excel.Worksheet;
                sel = app.Selection as Excel.Range;
                if (sheet == null || sel == null) return null;
                used = sheet.UsedRange;
                return app.Intersect(sel, used);
            }
            finally { ComUtil.Release(used, sel, sheet); }
        }

        public static void PercentFormat()
        {
            Excel.Range target = GetUsedSelection();
            if (target == null) return;

            try
            {
                target.ShrinkToFit = true;
                target.NumberFormatLocal = PercentFmt;
            }
            finally { ComUtil.Release(target); }
        }

        public static void AccountingFormat()
        {
            Excel.Range target = GetUsedSelection();
            if (target == null) return;

            try
            {
                // 同 Ctrl+9，档位从当前格式串反推，不存状态。首档给最常用的 2 位小数
                string current = null;
                try { current = target.NumberFormatLocal as string; }
                catch { }

                string next;
                if (current == AcctFormat2) next = AcctFormat0;
                else if (current == AcctFormat0) next = AcctFormat3;
                else next = AcctFormat2;

                target.ShrinkToFit = true;
                target.NumberFormatLocal = next;
            }
            finally { ComUtil.Release(target); }
        }

        
        // Ctrl+9 两档轮换：万 → 亿 → 回到万。零值空白、红字负数、文本原样。
        // 量级词不带「元」，国际项目币种不定。想看原始数字按 Ctrl+0 的千分位会计格式，
        // 英文 M/k 缩写档已取消：国际场景千分位本身就够读，缩写反而多一层心算。
        //
        // 单一量级是零值空白的前提。带条件的格式只有 条件1;条件2;其余 三段，没有独立的
        // 零值段和文本段，零值只能跟着量级段显示成 0.0万；不带条件才是标准的
        // 正;负;零;文本 四段，零值和颜色才放得下。原先亿/万合一档要吃掉两个条件位
        // 做量级分档，正是它没法让零值空白的原因。Excel 最多 2 个条件，写 3 个直接报错。
        //
        // 不能写成 #,##0!.0 —— !. 硬插小数点的技巧和 #,## 千分位冲突，格式会失效并
        // 漏出字面的 ".," 字符
        private const string UnitFormatWan = @"0!.0,""万"";[红色]-0!.0,""万"";;@";
        private const string UnitFormatYi = @"0!.00,,""亿"";[红色]-0!.00,,""亿"";;@";

        public static void YiWanFormat()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range sel = app.Selection as Excel.Range;
            if (sel == null) return;

            try
            {
                // 当前处于哪一档直接从格式串反推，不存状态。换选区、换工作簿都能自己对上；
                // 选区内格式不统一时读不到统一格式串，和任何一档都不相等，落到第一档
                string current = null;
                try { current = sel.NumberFormatLocal as string; }
                catch { }

                string next = current == UnitFormatWan ? UnitFormatYi : UnitFormatWan;

                sel.ShrinkToFit = true;
                sel.NumberFormatLocal = next;
            }
            finally { ComUtil.Release(sel); }
        }

        // =========================================================================================
        // M_Format.bas Migration
        // =========================================================================================

        public static void RowGroup(Excel.Range rng, int col, string cond, int lev)
        {
            if (rng == null) return;
            if (lev < 1) lev = 1;
            if (lev > 8) lev = 8; // Excel max outline level is 8

            Excel.Range rows = null;
            try
            {
                rows = rng.Rows;
                int rowCount = rows.Count;
                for (int i = 1; i <= rowCount; i++)
                {
                    Excel.Range cell = null;
                    Excel.Range row = null;
                    try
                    {
                        cell = rng.Cells[i, col] as Excel.Range;
                        string val = Convert.ToString(cell.Value2);
                        if (Like(val, cond))
                        {
                            row = rows[i] as Excel.Range;
                            row.OutlineLevel = lev;
                        }
                    }
                    catch { }
                    finally { ComUtil.Release(row, cell); }
                }
            }
            finally { ComUtil.Release(rows); }
        }


        public static void RowStyle()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Range rng = null;
            try
            {
                rng = app.Selection as Excel.Range;
                RowStyleInternal(rng);
            }
            finally { ComUtil.Release(rng); }
        }

        private static void RowStyleInternal(Excel.Range rng)
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Font font = null;
            Excel.Areas areas = null;

            using (ExcelScope.Begin(app))
            try
            {
                if (rng == null) return;

                // 全局一次性设置字体，极大减少 COM 调用
                try
                {
                    font = rng.Font;
                    font.Name = "Microsoft YaHei UI";
                }
                catch { }

                areas = rng.Areas;
                int areaCount = areas.Count;
                for (int a = 1; a <= areaCount; a++)
                {
                    Excel.Range area = null;
                    Excel.Range rows = null;
                    try
                    {
                        area = areas[a];
                        rows = area.Rows;
                        int rowCount = rows.Count;
                        if (rowCount <= 0) continue;

                        for (int i = 1; i <= rowCount; i++)
                        {
                            Excel.Range row = null;
                            try
                            {
                                row = rows[i] as Excel.Range;
                                int level = (int)row.OutlineLevel;
                                ApplyStyleToRow(row, level);
                            }
                            catch { }
                            finally { ComUtil.Release(row); }
                        }
                    }
                    finally { ComUtil.Release(rows, area); }
                }
            }
            catch (Exception)
            {
                // Optionally log error
            }
            finally { ComUtil.Release(areas, font); }
        }

        // 分级配色集中定义。判定「这底色是不是本插件刷的」和实际刷色必须用同一份，
        // 两处各写一遍字面量，就会出现换级别时旧底色被当成手动填充保留下来的情况
        private static readonly int GradingBg1 = ColorTranslator.ToOle(Color.FromArgb(51, 63, 79));      // #333F4F 深蓝灰
        private static readonly int GradingBg2 = ColorTranslator.ToOle(Color.FromArgb(217, 225, 242));   // #D9E1F2 浅蓝
        private static readonly int GradingBg3 = ColorTranslator.ToOle(Color.FromArgb(251, 229, 214));   // #FBE5D6 浅橙
        private static readonly int GradingFont1 = ColorTranslator.ToOle(Color.FromArgb(255, 255, 255)); // #FFFFFF 白色
        private static readonly int GradingFont23 = ColorTranslator.ToOle(Color.FromArgb(26, 26, 26));   // #1A1A1A

        private static void ApplyStyleToRow(Excel.Range row, int level)
        {
            if (row == null) return;
            if (level < 1 || level > 3) return;

            int bg = level == 1 ? GradingBg1 : (level == 2 ? GradingBg2 : GradingBg3);
            int fg = level == 1 ? GradingFont1 : GradingFont23;

            Excel.Interior interior = null;
            Excel.Font font = null;
            Excel.Borders borders = null;
            Excel.Border bottom = null;
            try
            {
                interior = row.Interior;
                interior.Pattern = Excel.XlPattern.xlPatternSolid;
                interior.Color = bg;
                interior.TintAndShade = 0;

                font = row.Font;
                font.Color = fg;
                font.TintAndShade = 0;
                font.Bold = true;

                // L3 不动下边框，保持原样
                if (level != 3)
                {
                    borders = row.Borders;
                    bottom = borders[Excel.XlBordersIndex.xlEdgeBottom];
                    bottom.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
            }
            catch { }
            finally { ComUtil.Release(bottom, borders, font, interior); }
        }

        
        public static void ClearStyle()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Range sel = null;
            try
            {
                sel = app.Selection as Excel.Range;
                ClearStyleInternal(sel);
            }
            finally { ComUtil.Release(sel); }
        }

        private static void ClearStyleInternal(Excel.Range rng)
        {
            if (rng == null) return;

            Excel.Interior interior = null;
            Excel.Font font = null;
            try
            {
                interior = rng.Interior;
                interior.Pattern = Excel.XlPattern.xlPatternNone;

                font = rng.Font;
                font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                font.Bold = false;
            }
            catch { }
            finally { ComUtil.Release(font, interior); }
        }

        public static bool IsStyled(Excel.Range rng)
        {
            if (rng == null) return false;
            try
            {
                foreach (Excel.Range cell in rng)
                {
                    Excel.Font font = null;
                    Excel.Interior interior = null;
                    try
                    {
                        font = cell.Font;
                        interior = cell.Interior;
                        if ((int)font.ColorIndex != (int)Excel.Constants.xlAutomatic ||
                            (int)interior.ColorIndex != (int)Excel.Constants.xlNone)
                        {
                            return true;
                        }
                    }
                    finally { ComUtil.Release(interior, font, cell); }
                }
            }
            catch { }
            return false;
        }

        
        public static void SetGrading()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Range selection = null;
            Excel.Worksheet activeSheet = null;
            Excel.Range used = null;
            Excel.Range rng = null;
            Excel.Range rngRows = null;
            Excel.Range rngCols = null;

            using (ExcelScope.Begin(app))
            try
            {
                selection = app.Selection as Excel.Range;
                activeSheet = app.ActiveSheet as Excel.Worksheet;
                if (selection == null || activeSheet == null) return;

                // 1) 确定范围 - Intersect 确保只处理有数据的区域
                used = activeSheet.UsedRange;
                rng = app.Intersect(selection, used);
                if (rng == null) return;

                rngRows = rng.Rows;
                rngCols = rng.Columns;
                int rowCount = rngRows.Count;
                int colCount = rngCols.Count;

                // 性能优化：一次性将整个区域的数据读入内存数组
                // 这比循环数千次读取单元格要快 100 倍以上
                object[,] dataValues = null;
                if (rowCount > 1 || colCount > 1)
                {
                    dataValues = rng.Value2 as object[,];
                }
                else
                {
                    // 单个单元格处理
                    dataValues = new object[2, 2];
                    dataValues[1, 1] = rng.Value2;
                }

                // 3) 循环数组并设置等级
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
                    Excel.Range row = null;
                    try
                    {
                        row = rngRows[i] as Excel.Range;
                        if ((int)row.OutlineLevel != level)
                        {
                            row.OutlineLevel = level;
                        }
                    }
                    finally { ComUtil.Release(row); }
                }
            }
            catch (Exception)
            {
                // 可以按需记录异常
            }
            finally
            {
                ComUtil.Release(rngCols, rngRows, rng, used, activeSheet, selection);
            }
        }

        
        public static void SetGradingStyle()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Range selection = null;
            Excel.Worksheet activeSheet = null;
            Excel.Range used = null;
            Excel.Range rng = null;
            try
            {
                selection = app.Selection as Excel.Range;
                activeSheet = app.ActiveSheet as Excel.Worksheet;
                if (selection == null || activeSheet == null) return;

                used = activeSheet.UsedRange;
                rng = app.Intersect(selection, used);
                if (rng == null) return;

                ClearAndRestyleByLevel(app, rng);
            }
            finally { ComUtil.Release(rng, used, activeSheet, selection); }
        }

        // L1~L3 标题行的底色和字色全由级别决定，手动填充一律覆盖，不做保留
        private static void ClearAndRestyleByLevel(Excel.Application app, Excel.Range rng)
        {
            if (rng == null) return;

            Excel.Areas areas = null;

            using (ExcelScope.Begin(app))
            try
            {
                areas = rng.Areas;
                int areaCount = areas.Count;
                for (int a = 1; a <= areaCount; a++)
                {
                    Excel.Range area = null;
                    Excel.Range rows = null;
                    try
                    {
                        area = areas[a];
                        rows = area.Rows;
                        int rowCount = rows.Count;
                        if (rowCount <= 0) continue;

                        for (int i = 1; i <= rowCount; i++)
                        {
                            Excel.Range row = null;
                            try
                            {
                                row = rows[i] as Excel.Range;
                                int level = (int)row.OutlineLevel;
                                if (level < 1 || level > 3) continue;

                                // 先清成无填充，把图案填充和渐变填充也一并去掉，再刷分级底色
                                ClearStyleInternal(row);
                                ApplyStyleToRow(row, level);
                            }
                            catch { }
                            finally { ComUtil.Release(row); }
                        }
                    }
                    finally { ComUtil.Release(rows, area); }
                }
            }
            catch (Exception)
            {
            }
            finally { ComUtil.Release(areas); }
        }


        public static void MultiAreaGroup()
        {
            MultiAreaOutline(true);
        }


        public static void MultiAreaUngroup()
        {
            MultiAreaOutline(false);
        }

        private static void MultiAreaOutline(bool group)
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Range selection = null;
            Excel.Range visible = null;
            Excel.Areas areas = null;
            try
            {
                selection = app.Selection as Excel.Range;
                if (selection == null) return;

                visible = selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                areas = visible.Areas;
                int areaCount = areas.Count;
                for (int a = 1; a <= areaCount; a++)
                {
                    Excel.Range area = null;
                    Excel.Range rows = null;
                    try
                    {
                        area = areas[a];
                        rows = area.Rows;
                        int rowCount = rows.Count;
                        for (int i = 1; i <= rowCount; i++)
                        {
                            Excel.Range row = null;
                            try
                            {
                                row = rows[i] as Excel.Range;
                                if (group) row.Group(); else row.Ungroup();
                            }
                            finally { ComUtil.Release(row); }
                        }
                    }
                    finally { ComUtil.Release(rows, area); }
                }
            }
            catch { }
            finally { ComUtil.Release(areas, visible, selection); }
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

            Excel.Range selection = null;
            Excel.Range visible = null;
            try
            {
                selection = app.Selection as Excel.Range;
                if (selection == null) return;
                visible = selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                visible.Select();
            }
            catch { }
            finally { ComUtil.Release(visible, selection); }
        }

        public static void SelectNonEmptyCells()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Range selection = null;
            Excel.Range constants = null;
            Excel.Range formulas = null;
            Excel.Range union = null;
            try
            {
                selection = app.Selection as Excel.Range;
                if (selection == null) return;

                // SpecialCells can only find one type at a time.
                // We need to combine constants (xlCellTypeConstants = 2) and formulas (xlCellTypeFormulas = -4123)
                try { constants = selection.SpecialCells(Excel.XlCellType.xlCellTypeConstants); } catch { }
                try { formulas = selection.SpecialCells(Excel.XlCellType.xlCellTypeFormulas); } catch { }

                if (constants != null && formulas != null)
                {
                    union = app.Union(constants, formulas);
                    union.Select();
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
            finally { ComUtil.Release(union, formulas, constants, selection); }
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

        // [1]Sheet1!A1 里方括号中是链接索引，'C:\x\[Book.xlsx]Sheet1'!A1 里是工作簿文件名
        private static readonly string[] WorkbookExtensions =
            { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".csv", ".et", ".ett", ".dbf" };

        // 单次从 Excel 取回的格子数上限，按行切块用
        private const int ChunkCellBudget = 400000;

        private static bool Cn()
        {
            return WpsExcelAddIn.UseChineseRibbon;
        }

        public static void BreakExternalLinks()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            var stats = new CleanStats();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            List<NameCommands.NameEntry> doomed = null;

            using (ExcelScope.Begin(app))
            try
            {
                // Step 1: 名称只遍历一次，同时拿到要删的 Name 对象和它们的短名
                app.StatusBar = Cn() ? "删除外部链接: 扫描名称..." : "Break Links: scanning names...";
                System.Windows.Forms.Application.DoEvents();

                doomed = CollectDoomedNames(workbook);

                Array links = workbook.LinkSources(Excel.XlLink.xlExcelLinks) as Array;
                stats.ExtLinksFound = (links == null ? 0 : links.Length) + doomed.Count;

                if (stats.ExtLinksFound == 0)
                {
                    app.StatusBar = false;
                    MessageBox.Show(
                        Cn() ? "没有检测到外部链接，无需清理。" : "No external links detected. Cleanup is not needed.",
                        Cn() ? "提示" : "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var doomedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (NameCommands.NameEntry d in doomed) doomedNames.Add(d.ShortName);

                // Step 2: 引用外链或引用待删名称的公式转成值。名称一删这些公式就变
                // #NAME?，先转值才留得住单元格里最后一次算出的结果
                ConvertExternalFormulas(app, workbook, doomedNames, ref stats);

                // Step 3: 删名称
                DeleteDoomedNames(app, workbook, doomed, ref stats);

                // Step 4: 断开链接源
                BreakAllLinks(workbook, app, ref stats);

                app.StatusBar = false;
                sw.Stop();

                MessageBox.Show(
                    (Cn()
                        ? "外部链接已断开并粘为数值，无效定义名称已清理。\n---\n"
                        + "检测到的外部链接 / 坏名称: " + stats.ExtLinksFound + "\n"
                        + "公式转为数值: " + stats.CellsConverted + "\n"
                        + "删除定义名称: " + stats.NamesDeleted + "\n"
                        + "断开链接: " + stats.LinksBroken + "\n"
                        + "错误: " + stats.Errors + "\n"
                        + "耗时: " + sw.Elapsed.TotalSeconds.ToString("0.00") + " 秒\n"
                        + "\n保存后重新打开工作簿，应该不会再弹出「更新链接」提示。"
                        : "All external links have been broken into values, and invalid defined names have been cleaned up.\n---\n"
                        + "External link sources / broken names detected: " + stats.ExtLinksFound + "\n"
                        + "Formulas converted to values: " + stats.CellsConverted + "\n"
                        + "Defined names deleted: " + stats.NamesDeleted + "\n"
                        + "Links broken: " + stats.LinksBroken + "\n"
                        + "Errors: " + stats.Errors + "\n"
                        + "Time: " + sw.Elapsed.TotalSeconds.ToString("0.00") + " seconds\n"
                        + "\nSave the workbook and reopen it. The \"update links\" prompt should no longer appear."),
                    Cn() ? "完成" : "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                app.StatusBar = false;
                MessageBox.Show(
                    (Cn() ? "删除外部链接时出错: " : "Error while removing external links: ") + ex.Message,
                    Cn() ? "错误" : "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                NameCommands.ReleaseEntries(doomed);
                ComUtil.Release(workbook);
            }
        }

        /// <summary>
        /// 名称只遍历一次，直接把 Name 对象攥在手里。原来扫一遍、删坏名称再扫一遍、
        /// 清隐藏名称又扫一遍，而且每次删除都回头 wb.Names.Item(名字) 按名线性查找，
        /// 上万个名称时光查找就耗掉几十秒。
        /// </summary>
        private static List<NameCommands.NameEntry> CollectDoomedNames(Excel.Workbook wb)
        {
            var list = new List<NameCommands.NameEntry>();

            // NameEntry.ComName 要留给后面的删除步骤，所以这里遍历到的 Name 对象不能
            // 一律释放 —— 不入列表的那些由 NameCommands 那边统一收尾
            Excel.Names names = null;
            try
            {
                names = wb.Names;
                foreach (Excel.Name nm in names)
                {
                    string full;
                    try { full = nm.Name; }
                    catch { ComUtil.Release(nm); continue; }

                    // _xlfn 是本版本不认识的新函数占位、_xlpm 是 LAMBDA 参数，它们的
                    // RefersTo 正常情况下就是 #NAME?。删掉会直接把公式打坏
                    if (NameCommands.IsNeverDelete(full)) { ComUtil.Release(nm); continue; }

                    // 名称坏到 RefersTo 都读不出来时，读不出来本身就说明该删
                    string refersTo;
                    try { refersTo = nm.RefersTo as string; }
                    catch { refersTo = null; }

                    if (!IsBadNameRefersTo(refersTo)) { ComUtil.Release(nm); continue; }

                    list.Add(new NameCommands.NameEntry
                    {
                        ComName = nm,
                        FullName = full,
                        ShortName = NameCommands.StripSheetPrefix(full),
                        RefersTo = refersTo ?? ""
                    });
                }
            }
            finally { ComUtil.Release(names); }

            return list;
        }

        private static void ConvertExternalFormulas(Excel.Application app, Excel.Workbook wb,
            HashSet<string> doomedNames, ref CleanStats stats)
        {
            List<string> sheetNames = ComUtil.GetSheetNames(wb);
            int sheetTotal = sheetNames.Count;
            int sheetIndex = 0;

            foreach (string sheetName in sheetNames)
            {
                sheetIndex++;
                app.StatusBar = string.Format(
                    Cn() ? "删除外部链接: 处理工作表 ({0}/{1}) {2}..." : "Break Links: sheet ({0}/{1}) {2}...",
                    sheetIndex, sheetTotal, sheetName);
                System.Windows.Forms.Application.DoEvents();

                Excel.Worksheet ws = null;
                Excel.Range used = null;
                Excel.Range usedRowsRange = null;
                Excel.Range usedColsRange = null;
                try
                {
                    ws = ComUtil.GetSheet(wb, sheetName);
                    if (ws == null) continue;

                    try { used = ws.UsedRange; }
                    catch { continue; }
                    if (used == null) continue;

                    int usedRows, usedCols, firstRow, firstCol;
                    try
                    {
                        usedRowsRange = used.Rows;
                        usedColsRange = used.Columns;
                        usedRows = usedRowsRange.Count;
                        usedCols = usedColsRange.Count;
                        firstRow = used.Row;
                        firstCol = used.Column;
                    }
                    catch { continue; }
                    if (usedRows < 1 || usedCols < 1) continue;

                    // 一次取一整块比逐个 Areas 快得多（每个 Area 都要 4 次以上 COM 往返，
                    // 公式分散时 Areas 能有上万个），但 UsedRange 常被脏文件撑到几十万空行，
                    // 整块读会吃光内存。按行切块取一个折中
                    int chunkRows = Math.Max(1, ChunkCellBudget / usedCols);

                    for (int offset = 0; offset < usedRows; offset += chunkRows)
                    {
                        int height = Math.Min(chunkRows, usedRows - offset);
                        ConvertChunk(ws, firstRow + offset, firstCol, height, usedCols, doomedNames, ref stats);
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                finally { ComUtil.Release(usedColsRange, usedRowsRange, used, ws); }
            }
        }

        private static void ConvertChunk(Excel.Worksheet ws, int firstRow, int firstCol,
            int rows, int cols, HashSet<string> doomedNames, ref CleanStats stats)
        {
            Excel.Range block = null;
            try
            {
                try
                {
                    block = ComUtil.GetRange(ws, firstRow, firstCol,
                        firstRow + rows - 1, firstCol + cols - 1);
                }
                catch { return; }

                object formulaObj;
                try { formulaObj = block.Formula; }
                catch { return; }

                object[,] formulas = formulaObj as object[,];
                if (formulas == null)
                {
                    if (NeedsFlatten(formulaObj as string, doomedNames))
                    {
                        try { block.Value2 = block.Value2; stats.CellsConverted++; }
                        catch { stats.Errors++; }
                    }
                    return;
                }

                int rLo = formulas.GetLowerBound(0);
                int cLo = formulas.GetLowerBound(1);

                var hit = new bool[rows, cols];
                int hitCount = 0;

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        if (!NeedsFlatten(formulas[rLo + r, cLo + c] as string, doomedNames)) continue;
                        hit[r, c] = true;
                        hitCount++;
                    }
                    if ((r & 511) == 0) System.Windows.Forms.Application.DoEvents();
                }

                if (hitCount == 0) return;

                object[,] values = null;
                try { values = block.Value2 as object[,]; }
                catch { }
                if (values == null) { stats.Errors++; return; }

                WriteBackValues(ws, hit, values, rLo, cLo, rows, cols, firstRow, firstCol, ref stats);
            }
            finally { ComUtil.Release(block); }
        }

        /// <summary>
        /// 只把命中的单元格写回，并把连续的命中区域拼成矩形整块写。
        /// 原来是把整个区域的 Formula 数组整体赋回去，等于让 Excel 把没问题的
        /// 公式也重新解析一遍，一格外链能拖累十万格。
        /// </summary>
        private static void WriteBackValues(Excel.Worksheet ws,
            bool[,] hit, object[,] values, int rLo, int cLo, int rows, int cols,
            int firstRow, int firstCol, ref CleanStats stats)
        {
            var segments = new List<int[]>();
            int i = 0;

            while (i < rows)
            {
                RowSegments(hit, i, cols, segments);
                if (segments.Count == 0) { i++; continue; }

                // 相邻行的命中形状完全一样就并成一个矩形。整列都是外链公式时，
                // 一次写代替上万次写
                int j = i + 1;
                while (j < rows && SameSegments(hit, j, cols, segments)) j++;

                foreach (int[] seg in segments)
                {
                    int h = j - i;
                    int w = seg[1] - seg[0] + 1;

                    var buffer = new object[h, w];
                    for (int rr = 0; rr < h; rr++)
                        for (int cc = 0; cc < w; cc++)
                            buffer[rr, cc] = values[rLo + i + rr, cLo + seg[0] + cc];

                    Excel.Range target = null;
                    try
                    {
                        target = ComUtil.GetRange(ws,
                            firstRow + i, firstCol + seg[0],
                            firstRow + j - 1, firstCol + seg[1]);
                        target.Value2 = buffer;
                        stats.CellsConverted += h * w;
                    }
                    catch { stats.Errors++; }
                    finally { ComUtil.Release(target); }
                }

                i = j;
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private static void RowSegments(bool[,] hit, int row, int cols, List<int[]> into)
        {
            into.Clear();
            int c = 0;
            while (c < cols)
            {
                if (!hit[row, c]) { c++; continue; }
                int start = c;
                while (c < cols && hit[row, c]) c++;
                into.Add(new int[] { start, c - 1 });
            }
        }

        private static bool SameSegments(bool[,] hit, int row, int cols, List<int[]> segments)
        {
            int index = 0;
            int c = 0;
            while (c < cols)
            {
                if (!hit[row, c]) { c++; continue; }
                int start = c;
                while (c < cols && hit[row, c]) c++;
                if (index >= segments.Count) return false;
                if (segments[index][0] != start || segments[index][1] != c - 1) return false;
                index++;
            }
            return index == segments.Count && index > 0;
        }

        private static void DeleteDoomedNames(Excel.Application app, Excel.Workbook wb,
            List<NameCommands.NameEntry> doomed, ref CleanStats stats)
        {
            for (int i = 0; i < doomed.Count; i++)
            {
                if (i % 100 == 0)
                {
                    app.StatusBar = string.Format(
                        Cn() ? "删除外部链接: 删除名称 {0}/{1} ..." : "Break Links: deleting names {0}/{1} ...",
                        i + 1, doomed.Count);
                    System.Windows.Forms.Application.DoEvents();
                }

                // 走净化名称那套删除：Print_Area 要先清 PageSetup 否则 Excel 立刻重建，
                // RefersTo 已经是 #REF! 的名称 Delete 会静默失效，得先指回合法区域
                if (NameCommands.TryDeleteName(wb, doomed[i])) stats.NamesDeleted++;
                else stats.Errors++;
            }
        }

        private static void BreakAllLinks(Excel.Workbook wb, Excel.Application app, ref CleanStats stats)
        {
            Array links = wb.LinkSources(Excel.XlLink.xlExcelLinks) as Array;
            if (links == null) return;

            for (int i = links.GetLowerBound(0); i <= links.GetUpperBound(0); i++)
            {
                app.StatusBar = string.Format(
                    Cn() ? "删除外部链接: 断开链接 {0}/{1}" : "Break Links: disconnecting {0}/{1}",
                    i + 1, links.Length);
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

        private static bool NeedsFlatten(string formulaText, HashSet<string> doomedNames)
        {
            if (string.IsNullOrEmpty(formulaText)) return false;
            // 常量不是公式。Formula 取回来的公式一定带等号，数组公式是 {=...}
            if (formulaText[0] != '=' && !formulaText.StartsWith("{=")) return false;

            if (HasExternalRef(formulaText)) return true;
            if (doomedNames.Count == 0) return false;

            foreach (string ident in NameCommands.Identifiers(formulaText))
            {
                // 公式里满地都是 AA:AA、$B$3，不排掉就会拿列标去撞同名的待删名称，
                // 把根本不碰外链的公式一起粘死
                if (LooksLikeCellRef(ident)) continue;
                if (doomedNames.Contains(ident)) return true;
            }

            return false;
        }

        private static bool HasExternalRef(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            int i = 0;
            while ((i = text.IndexOf('[', i)) >= 0)
            {
                int close = text.IndexOf(']', i + 1);
                if (close < 0) break;
                if (IsWorkbookToken(text, i + 1, close - i - 1)) return true;
                i = close + 1;
            }

            return false;
        }

        /// <summary>
        /// 方括号里装的是链接索引或工作簿文件名才算外部引用。
        /// Table1[列名] 是同表结构化引用，当成外链会把正常公式粘死。
        /// </summary>
        private static bool IsWorkbookToken(string text, int start, int length)
        {
            if (length <= 0) return false;

            bool allDigits = true;
            for (int i = 0; i < length; i++)
            {
                char ch = text[start + i];
                if (ch == '\\' || ch == ':') return true;
                if (!char.IsDigit(ch)) allDigits = false;
            }
            if (allDigits) return true;

            string token = text.Substring(start, length);
            foreach (string ext in WorkbookExtensions)
                if (token.EndsWith(ext, StringComparison.OrdinalIgnoreCase)) return true;

            return false;
        }

        // Excel 建名称时就不许用单元格引用形式，所以长这样的标识符一定是引用不是名称
        private static bool LooksLikeCellRef(string token)
        {
            if (string.IsNullOrEmpty(token) || token.Length > 10) return false;

            int i = 0;
            int column = 0;
            while (i < token.Length && token[i] < 128 && char.IsLetter(token[i]))
            {
                column = column * 26 + (char.ToUpperInvariant(token[i]) - 'A' + 1);
                i++;
                if (i > 3) return false;
            }

            if (i == 0 || column < 1 || column > 16384) return false;
            if (i == token.Length) return true;

            long row = 0;
            for (; i < token.Length; i++)
            {
                if (!char.IsDigit(token[i])) return false;
                row = row * 10 + (token[i] - '0');
                if (row > 1048576) return false;
            }

            return row >= 1;
        }

        private static bool IsBadNameRefersTo(string refersTo)
        {
            if (string.IsNullOrEmpty(refersTo)) return true;

            foreach (string t in ErrTokens)
                if (refersTo.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0) return true;

            if (refersTo.IndexOf(":\\", StringComparison.Ordinal) >= 0) return true;
            if (refersTo.IndexOf("\\\\", StringComparison.Ordinal) >= 0) return true;
            if (refersTo.IndexOf("://", StringComparison.OrdinalIgnoreCase) >= 0) return true;

            return HasExternalRef(refersTo);
        }
    }
}

