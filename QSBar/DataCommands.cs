﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{

    public static class DataCommands

    {

        

        public static void BatchProcess()

        {

            using (var form = new BatchProcessForm())
            {
                form.ShowDialog();
            }

        }



        

        /// <summary>
        /// 取「选区 ∩ 已用区域」。调用方负责 Release 返回的 Range。
        /// </summary>
        private static Excel.Range GetUsedSelection(Excel.Application app)
        {
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

        public static void Textify()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Range target = GetUsedSelection(app);
            if (target == null) return;

            using (ExcelScope.Begin(app))
            try
            {
                object value2 = target.Value2;
                if (!(value2 is object[,]))
                {
                    object val = value2;
                    if (val != null)
                    {
                        bool isError = false;
                        if (val is int)
                        {
                            if (IsExcelErrorCode((int)val))
                                isError = true;
                        }

                        if (!isError)
                        {
                            string s = val.ToString();
                            if (s.Length > 0 && s[0] != '=')
                            {
                                // 无论是不是数字，都加上单引号前缀强制转为文本
                                val = "'" + s;
                            }
                        }
                    }

                    target.NumberFormatLocal = "@";
                    target.Value2 = val;
                    target.ShrinkToFit = true;
                    return;
                }

                object[,] data = (object[,])value2;
                int rows = data.GetLength(0);
                int cols = data.GetLength(1);
                int rBase = data.GetLowerBound(0);
                int cBase = data.GetLowerBound(1);

                for (int i = rBase; i < rBase + rows; i++)
                {
                    for (int j = cBase; j < cBase + cols; j++)
                    {
                        object val = data[i, j];
                        if (val == null) continue;

                        if (val is int)
                        {
                            if (IsExcelErrorCode((int)val)) continue;
                        }

                        string s = val.ToString();
                        if (s.Length == 0) continue;
                        if (s[0] == '=') continue;

                        data[i, j] = "'" + s;
                    }
                }

                target.NumberFormatLocal = "@";
                target.Value2 = data;
                target.ShrinkToFit = true;
            }
            finally { ComUtil.Release(target); }
        }

        public static void NormalizeNumbers()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Range target = GetUsedSelection(app);
            if (target == null) return;

            using (ExcelScope.Begin(app))
            try
            {
                Excel.Hyperlinks links = null;
                try
                {
                    links = target.Hyperlinks;
                    links.Delete();
                }
                catch
                {
                }
                finally { ComUtil.Release(links); }

                object value2 = target.Value2;
                if (!(value2 is object[,]))
                {
                    object val = value2;
                    if (val != null)
                    {
                        bool isError = false;
                        if (val is int)
                        {
                            if (IsExcelErrorCode((int)val))
                                isError = true;
                        }

                        if (!isError)
                        {
                            string s = val.ToString();
                            if (s.Length > 0 && s[0] != '=')
                            {
                                double p;
                                DateTime dt;
                                double num;

                                if (s.EndsWith("%", StringComparison.Ordinal))
                                {
                                    string inner = s.Substring(0, s.Length - 1);
                                    if (double.TryParse(inner, out p))
                                    {
                                        val = (double)(p / 100.0);
                                    }
                                }
                                else if (double.TryParse(s, out num))
                                {
                                    val = num;
                                }
                                else if (DateTime.TryParse(s, out dt))
                                {
                                    val = dt;
                                }
                            }
                        }
                    }

                    target.Value2 = val;
                    target.ShrinkToFit = true;
                    target.NumberFormatLocal = " #,##0.00_ ;[红色] -#,##0.00_ ;_ \"\"\"\"?_ ;@";
                    return;
                }

                object[,] data = (object[,])value2;
                int rows = data.GetLength(0);
                int cols = data.GetLength(1);
                int rBase = data.GetLowerBound(0);
                int cBase = data.GetLowerBound(1);

                for (int i = rBase; i < rBase + rows; i++)
                {
                    for (int j = cBase; j < cBase + cols; j++)
                    {
                        object val = data[i, j];
                        if (val == null)
                        {
                            continue;
                        }

                        if (val is int)
                        {
                            if (IsExcelErrorCode((int)val))
                                continue;
                        }

                        string s = val.ToString();
                        if (s.Length == 0)
                        {
                            continue;
                        }

                        if (s[0] == '=')
                        {
                            continue;
                        }

                        double p;
                        DateTime dt;
                        double num;

                        if (s.EndsWith("%", StringComparison.Ordinal))
                        {
                            string inner = s.Substring(0, s.Length - 1);
                            if (double.TryParse(inner, out p))
                            {
                                data[i, j] = (double)(p / 100.0);
                                continue;
                            }
                        }

                        if (double.TryParse(s, out num))
                        {
                            data[i, j] = num;
                            continue;
                        }

                        if (DateTime.TryParse(s, out dt))
                        {
                            data[i, j] = dt;
                            continue;
                        }
                    }
                }

                target.Value2 = data;
                target.ShrinkToFit = true;
                target.NumberFormatLocal = " #,##0.00_ ;[红色] -#,##0.00_ ;_ \"\"\"\"?_ ;@";
            }
            finally { ComUtil.Release(target); }
        }

        private static bool IsExcelErrorCode(int iVal)
        {
            return iVal == -2146826246 || iVal == -2146826281 || iVal == -2146826259 || 
                   iVal == -2146826288 || iVal == -2146826252 || iVal == -2146826273 || 
                   iVal == -2146826265;
        }



        

        public static void ForceRefresh()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Worksheet sheet = null;
            Excel.Range rng = null;
            Excel.Range formulas = null;
            Excel.Areas areas = null;
            try
            {
                sheet = app.ActiveSheet as Excel.Worksheet;
                if (sheet == null) return;
                rng = sheet.UsedRange;
                if (rng == null) return;

                // xlCellTypeFormulas = -4123
                formulas = rng.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                if (formulas == null) return;

                areas = formulas.Areas;
                int areaCount = areas.Count;
                for (int a = 1; a <= areaCount; a++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = areas[a];
                        area.Formula = area.Formula;
                    }
                    finally { ComUtil.Release(area); }
                }
            }
            catch { }
            finally { ComUtil.Release(areas, formulas, rng, sheet); }
        }

        public static void ExpandPivotTable()
        {
            SetPivotDetail(true);
        }

        public static void CollapsePivotTable()
        {
            SetPivotDetail(false);
        }

        private static void SetPivotDetail(bool show)
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Range activeCell = null;
            Excel.PivotTable pt = null;
            Excel.PivotField pf = null;
            try
            {
                activeCell = app.ActiveCell;
                if (activeCell == null) return;

                pt = activeCell.PivotTable;
                if (pt == null) return;

                pf = activeCell.PivotField;
                if (pf != null) pf.ShowDetail = show;
            }
            catch { }
            finally { ComUtil.Release(pf, pt, activeCell); }
        }

        public static void PasteExternalLinks()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Worksheet activeSheet = null;
            Excel.Workbook wb = null;
            try
            {
                activeSheet = app.ActiveSheet as Excel.Worksheet;
                if (activeSheet == null) return;

                if (activeSheet.ProtectContents)
                {
                    MessageBox.Show("This worksheet is protected. The operation has been blocked.", "Notice");
                    return;
                }

                wb = app.ActiveWorkbook;
                if (wb == null) return;
            }
            finally { ComUtil.Release(activeSheet); }

            bool oldAlerts = app.DisplayAlerts;
            app.DisplayAlerts = false;

            // 循环里逐格 Value2 = Value2，自动计算开着的话每写一格就重算一遍全表
            using (ExcelScope.Begin(app))
            try
            {
                foreach (string sheetName in ComUtil.GetSheetNames(wb))
                {
                    Excel.Worksheet sht = null;
                    Excel.Range usedRange = null;
                    try
                    {
                        sht = ComUtil.GetSheet(wb, sheetName);
                        if (sht == null) continue;

                        usedRange = sht.UsedRange;
                        if (usedRange == null) continue;

                        // XlFindLookIn.xlFormulas = -4144
                        // XlLookAt.xlPart = 2
                        // XlSearchOrder.xlByRows = 1
                        Excel.Range cell = usedRange.Find("=*[*]*", Type.Missing, Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, true);

                        if (cell != null)
                        {
                            string firstAddress = cell.Address;
                            do
                            {
                                Excel.Range current = cell;
                                try
                                {
                                    current.Value2 = current.Value2;
                                    cell = usedRange.FindNext(current);
                                }
                                finally { ComUtil.Release(current); }
                            }
                            while (cell != null && cell.Address != firstAddress);

                            ComUtil.Release(cell);
                        }
                    }
                    finally { ComUtil.Release(usedRange, sht); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to process external links: " + ex.Message);
            }
            finally
            {
                app.DisplayAlerts = oldAlerts;
                ComUtil.Release(wb);
            }
        }
    }
}

