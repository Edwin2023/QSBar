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

            var form = new BatchProcessForm();

            form.ShowDialog();

        }



        

        public static void Textify()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Worksheet sheet = app.ActiveSheet as Excel.Worksheet;
            Excel.Range sel = app.Selection as Excel.Range;
            if (sheet == null || sel == null) return;

            Excel.Range used = sheet.UsedRange;
            Excel.Range target = app.Intersect(sel, used);
            if (target == null) return;

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

        public static void NormalizeNumbers()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Worksheet sheet = app.ActiveSheet as Excel.Worksheet;
            Excel.Range sel = app.Selection as Excel.Range;
            if (sheet == null || sel == null) return;
            
            Excel.Range used = sheet.UsedRange;
            Excel.Range target = app.Intersect(sel, used);
            if (target == null)
            {
                return;
            }

            try
            {
                target.Hyperlinks.Delete();
            }
            catch
            {
            }

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
            Excel.Worksheet sheet = app.ActiveSheet as Excel.Worksheet;
            if (sheet == null) return;
            Excel.Range rng = sheet.UsedRange;
            if (rng == null) return;

            try
            {
                // xlCellTypeFormulas = -4123
                Excel.Range formulas = rng.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                if (formulas == null) return;

                foreach (Excel.Range area in formulas.Areas)
                {
                    area.Formula = area.Formula;
                }
            }
            catch { }
        }

        public static void ExpandPivotTable()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range activeCell = app.ActiveCell;
            if (activeCell == null) return;

            try
            {
                Excel.PivotTable pt = activeCell.PivotTable;
                if (pt != null)
                {
                    Excel.PivotField pf = activeCell.PivotField;
                    if (pf != null)
                    {
                        pf.ShowDetail = true;
                    }
                }
            }
            catch { }
        }

        public static void CollapsePivotTable()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range activeCell = app.ActiveCell;
            if (activeCell == null) return;

            try
            {
                Excel.PivotTable pt = activeCell.PivotTable;
                if (pt != null)
                {
                    Excel.PivotField pf = activeCell.PivotField;
                    if (pf != null)
                    {
                        pf.ShowDetail = false;
                    }
                }
            }
            catch { }
        }

        public static void PasteExternalLinks()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
            if (activeSheet == null) return;

            if (activeSheet.ProtectContents)
            {
                MessageBox.Show("工作表已保护,本程序拒绝执行！", "提示");
                return;
            }

            app.ScreenUpdating = false;
            app.DisplayAlerts = false;

            try
            {
                foreach (Excel.Worksheet sht in app.ActiveWorkbook.Worksheets)
                {
                    Excel.Range usedRange = sht.UsedRange;
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
                            cell.Value2 = cell.Value2;
                            cell = usedRange.FindNext(cell);
                        }
                        while (cell != null && cell.Address != firstAddress);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("处理外部链接失败: " + ex.Message);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
            }
        }
    }
}

