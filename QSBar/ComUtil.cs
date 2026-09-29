using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    /// <summary>
    /// Excel COM 对象的取用与归还。
    ///
    /// 每一次 `a.b.c` 的中间步都会产生一个引用，攒着不放 Excel 就认为还有客户端在用它 ——
    /// 用户关掉窗口后 EXCEL.EXE 退不掉，下次启动进安全模式。凡是接过 Excel 对象的局部
    /// 变量都要在 finally 里 Release 一遍，尤其是 Application / Window / Workbook 这几层。
    ///
    /// 实测判据：关掉 Excel 后进程还在、`MainWindowHandle=0`，就是这里漏了。
    /// </summary>
    internal static class ComUtil
    {
        public static void Release(params object[] comObjects)
        {
            if (comObjects == null) return;
            foreach (object o in comObjects)
            {
                if (o == null) continue;
                try
                {
                    if (Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
                }
                catch { }
            }
        }

        public static void Release<T>(IEnumerable<T> comObjects)
        {
            if (comObjects == null) return;
            foreach (T o in comObjects)
            {
                if (o == null) continue;
                try
                {
                    if (Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
                }
                catch { }
            }
        }

        /// <summary>
        /// 取 ws 上的矩形区域，顺手把两个角单元格的引用还回去。
        /// </summary>
        public static Excel.Range GetRange(Excel.Worksheet ws, int row1, int col1, int row2, int col2)
        {
            Excel.Range topLeft = null;
            Excel.Range bottomRight = null;
            try
            {
                topLeft = ws.Cells[row1, col1] as Excel.Range;
                bottomRight = ws.Cells[row2, col2] as Excel.Range;
                return ws.Range[topLeft, bottomRight];
            }
            finally { Release(bottomRight, topLeft); }
        }

        public static void SetCellValue(Excel.Worksheet ws, int row, int col, object value)
        {
            Excel.Range cell = null;
            try
            {
                cell = ws.Cells[row, col] as Excel.Range;
                cell.Value = value;
            }
            finally { Release(cell); }
        }

        public static void SetCellFormula(Excel.Worksheet ws, int row, int col, string formula)
        {
            Excel.Range cell = null;
            try
            {
                cell = ws.Cells[row, col] as Excel.Range;
                cell.Formula = formula;
            }
            finally { Release(cell); }
        }

        /// <summary>
        /// 工作簿里所有工作表的名字。先取成字符串，调用方就不用攥着一堆 Worksheet 引用。
        /// </summary>
        public static List<string> GetSheetNames(Excel.Workbook wb, bool visibleOnly = false, string excludeName = null)
        {
            var names = new List<string>();
            Excel.Sheets sheets = null;
            try
            {
                sheets = wb.Worksheets;
                int count = sheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets[i] as Excel.Worksheet;
                        if (sheet == null) continue;
                        if (excludeName != null && sheet.Name == excludeName) continue;
                        if (visibleOnly && sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible) continue;
                        names.Add(sheet.Name);
                    }
                    finally { Release(sheet); }
                }
            }
            finally { Release(sheets); }
            return names;
        }

        /// <summary>
        /// 按名字取工作表，取不到返回 null。调用方负责 Release 拿到的这一个。
        /// </summary>
        public static Excel.Worksheet GetSheet(Excel.Workbook wb, string name)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = wb.Worksheets;
                int count = sheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet sheet = null;
                    bool keep = false;
                    try
                    {
                        sheet = sheets[i] as Excel.Worksheet;
                        if (sheet != null && sheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                        {
                            keep = true;
                            return sheet;
                        }
                    }
                    finally { if (!keep) Release(sheet); }
                }
            }
            finally { Release(sheets); }
            return null;
        }
    }
}
