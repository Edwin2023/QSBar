using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public static class SheetCommands
    {
        private static void Release(params object[] comObjects)
        {
            ComUtil.Release(comObjects);
        }

        private static Excel.Range GetRange(Excel.Worksheet ws, int row1, int col1, int row2, int col2)
        {
            return ComUtil.GetRange(ws, row1, col1, row2, col2);
        }

        private static void ApplyModernStyle(Excel.Worksheet ws, int totalRows, int totalCols)
        {
            if (totalRows < 1 || totalCols < 1) return;

            Excel.Range fullRange = null;
            Excel.Range headerRange = null;
            Excel.Range dataRange = null;
            Excel.Font fullFont = null;
            Excel.Font headerFont = null;
            Excel.Interior headerInterior = null;
            Excel.Range firstColumn = null;
            Excel.Borders borders = null;
            Excel.Range fullColumns = null;

            try
            {
                fullRange = GetRange(ws, 1, 1, totalRows, totalCols);
                headerRange = GetRange(ws, 1, 1, 1, totalCols);
                if (totalRows > 1)
                {
                    dataRange = GetRange(ws, 2, 1, totalRows, totalCols);
                }

                // 1. 基础字体设置
                fullFont = fullRange.Font;
                fullFont.Name = "微软雅黑";
                fullFont.Size = 10;
                fullRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                // 2. 表头样式：深蓝色背景 + 白色粗体
                headerRange.RowHeight = 30;
                headerFont = headerRange.Font;
                headerFont.Size = 11;
                headerFont.Bold = true;
                headerFont.Color = ColorTranslator.ToOle(Color.White);
                headerInterior = headerRange.Interior;
                headerInterior.Color = ColorTranslator.ToOle(Color.FromArgb(44, 62, 80)); // 优雅的深灰蓝
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // 3. 数据行样式
                if (dataRange != null)
                {
                    dataRange.RowHeight = 22;

                    // 隔行变色 (Banding) - 浅灰蓝色
                    for (int i = 2; i <= totalRows; i += 2)
                    {
                        Excel.Range band = null;
                        Excel.Interior bandInterior = null;
                        try
                        {
                            band = GetRange(ws, i, 1, i, totalCols);
                            bandInterior = band.Interior;
                            bandInterior.Color = ColorTranslator.ToOle(Color.FromArgb(242, 244, 247));
                        }
                        finally { Release(bandInterior, band); }
                    }

                    // 序号列居中
                    firstColumn = ws.Columns[1] as Excel.Range;
                    firstColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }

                // 4. 边框设置：极细浅灰色边框
                borders = fullRange.Borders;
                borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                borders.Weight = Excel.XlBorderWeight.xlThin;
                borders.Color = ColorTranslator.ToOle(Color.FromArgb(218, 223, 225));

                // 5. 自动列宽并留出一点边距
                fullColumns = fullRange.Columns;
                fullColumns.AutoFit();
                for (int i = 1; i <= totalCols; i++)
                {
                    Excel.Range column = null;
                    try
                    {
                        column = ws.Columns[i] as Excel.Range;
                        double currentWidth = (double)column.ColumnWidth;
                        column.ColumnWidth = currentWidth + 2;
                    }
                    finally { Release(column); }
                }
            }
            finally
            {
                Release(fullColumns, borders, firstColumn, headerInterior, headerFont, fullFont,
                        dataRange, headerRange, fullRange);
            }
        }

        /// <summary>
        /// 冻结首行、隐藏网格线。这两项是窗口级设置，必须在 ScreenUpdating 恢复之后、
        /// 目标表已激活的状态下做：屏幕刷新关着时改窗口的拆分/冻结会让 Excel 崩溃，
        /// 表没激活时改的还是别人那张表。
        /// </summary>
        private static void ApplySheetView(Excel.Worksheet ws)
        {
            if (ws == null) return;

            Excel.Window win = null;
            try
            {
                ws.Activate();

                // 走插件已经持有的 App，不用 ws.Application —— 后者每调一次就多一个
                // Application 引用，攥着不放 Excel 就退不掉
                Excel.Application app = WpsExcelAddIn.App;
                if (app == null) return;

                win = app.ActiveWindow;
                if (win == null) return;

                // 已冻结的窗口不接受新的 SplitRow，得先解冻
                win.FreezePanes = false;
                win.SplitColumn = 0;
                win.SplitRow = 1;
                win.FreezePanes = true;
                win.DisplayGridlines = false;
            }
            catch { }
            finally { Release(win); }
        }

        public static void CreateSheetIndex()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            Excel.Worksheet indexSheet = null;

            using (ExcelScope.Begin(app))
            try
            {
                indexSheet = AddSheetAtFront(workbook, "Index_" + DateTime.Now.ToString("HHmmss"));
                if (indexSheet == null) return;

                // 准备数据
                int rowNum = 1;
                SetCellValue(indexSheet, rowNum, 1, "No.");
                SetCellValue(indexSheet, rowNum, 2, "Worksheet Name");
                rowNum++;

                string indexName = indexSheet.Name;
                foreach (string sheetName in GetVisibleSheetNames(workbook, indexName))
                {
                    SetCellValue(indexSheet, rowNum, 1, rowNum - 1);
                    string formula = string.Format("=HYPERLINK(\"#'{0}'!A1\", \"{1}\")", sheetName.Replace("'", "''"), sheetName);
                    SetCellFormula(indexSheet, rowNum, 2, formula);
                    rowNum++;
                }

                // 应用美化样式
                ApplyModernStyle(indexSheet, rowNum - 1, 2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to generate the sheet index: " + ex.Message);
                return;
            }
            finally
            {
                Release(workbook);
            }

            ApplySheetView(indexSheet);
            Release(indexSheet);
        }

        /// <summary>
        /// 在最前面插一张表并命名。Worksheets 集合和当作锚点的第一张表都得还回去。
        /// </summary>
        private static Excel.Worksheet AddSheetAtFront(Excel.Workbook wb, string name)
        {
            Excel.Sheets sheets = null;
            Excel.Worksheet anchor = null;
            Excel.Worksheet created = null;
            try
            {
                sheets = wb.Worksheets;
                anchor = sheets[1] as Excel.Worksheet;
                created = sheets.Add(Before: anchor) as Excel.Worksheet;
                if (created != null) created.Name = name;
                return created;
            }
            finally { Release(anchor, sheets); }
        }

        private static System.Collections.Generic.List<string> GetVisibleSheetNames(Excel.Workbook wb, string excludeName)
        {
            return ComUtil.GetSheetNames(wb, true, excludeName);
        }

        private static void SetCellValue(Excel.Worksheet ws, int row, int col, object value)
        {
            ComUtil.SetCellValue(ws, row, col, value);
        }

        private static void SetCellFormula(Excel.Worksheet ws, int row, int col, string formula)
        {
            ComUtil.SetCellFormula(ws, row, col, formula);
        }

        /// <summary>
        /// 调宿主自己的 GetOpenFilename 多选文件，取消时返回 null。
        ///
        /// 不用 WinForms 的 OpenFileDialog：它会把第三方 shell 扩展（网盘之类）拉进宿主
        /// 进程，扩展赖着不放，Excel 关窗后进程退不掉，下次启动就进安全模式。实测宿主
        /// 自带的「打开」对话框没这个毛病，改用同一套 API 即可。AutoUpgradeEnabled=false
        /// 退回旧式对话框也救不了，试过。
        ///
        /// initialDir 靠临时切进程当前目录实现，GetOpenFilename 没有对应参数。
        /// </summary>
        private static string[] PickFiles(Excel.Application app, string filter, string title, string initialDir)
        {
            string prevDir = null;
            if (!string.IsNullOrEmpty(initialDir))
            {
                try
                {
                    prevDir = System.IO.Directory.GetCurrentDirectory();
                    System.IO.Directory.SetCurrentDirectory(initialDir);
                }
                catch { prevDir = null; }
            }

            object picked;
            try
            {
                picked = app.GetOpenFilename(filter, Type.Missing, title, Type.Missing, true);
            }
            finally
            {
                if (prevDir != null) { try { System.IO.Directory.SetCurrentDirectory(prevDir); } catch { } }
            }

            // 取消时返回 false，选中时是 1 基的字符串数组
            Array picks = picked as Array;
            if (picks == null) return null;

            string[] files = new string[picks.Length];
            int i = 0;
            foreach (object item in picks)
            {
                files[i++] = Convert.ToString(item);
            }
            return files;
        }

        public static void CreateFileIndex()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            string initialDir = "";
            try
            {
                if (!string.IsNullOrEmpty(workbook.Path))
                {
                    initialDir = workbook.Path;
                }
            }
            catch { }

            string[] selectedFiles = PickFiles(app, "All Files (*.*),*.*", "Select files to index", initialDir);
            if (selectedFiles == null || selectedFiles.Length == 0) return;

            Excel.Worksheet indexSheet = null;

            using (ExcelScope.Begin(app))
            try
            {
                // 每次都新建一张表，重名就往后排号，不覆盖之前生成的目录
                indexSheet = AddSheetAtFront(workbook, GetUniqueSheetName(workbook, "File Index"));
                if (indexSheet == null) return;

                // 准备数据
                int rowNum = 1;
                SetCellValue(indexSheet, rowNum, 1, "No.");
                SetCellValue(indexSheet, rowNum, 2, "File Name");
                SetCellValue(indexSheet, rowNum, 3, "Full Path");
                rowNum++;

                foreach (string filePath in selectedFiles)
                {
                    string fileName = System.IO.Path.GetFileName(filePath);
                    SetCellValue(indexSheet, rowNum, 1, rowNum - 1);

                    string formula = string.Format("=HYPERLINK(\"{0}\", \"{1}\")", filePath, fileName);
                    SetCellFormula(indexSheet, rowNum, 2, formula);
                    SetCellValue(indexSheet, rowNum, 3, filePath);
                    rowNum++;
                }

                // 应用美化样式
                ApplyModernStyle(indexSheet, rowNum - 1, 3);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to generate the file index: " + ex.Message);
                return;
            }
            finally
            {
                Release(workbook);
            }

            ApplySheetView(indexSheet);
            Release(indexSheet);
        }

        
        public static void MergeWorkbooks()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook targetWb = app.ActiveWorkbook;
            if (targetWb == null) return;

            string[] selectedFiles = PickFiles(app,
                "Excel Files (*.xls;*.xlsx;*.xlsm;*.xlsb),*.xls;*.xlsx;*.xlsm;*.xlsb,All Files (*.*),*.*",
                "Select Excel files to merge", null);
            if (selectedFiles == null || selectedFiles.Length == 0) return;

            bool originalAlerts = app.DisplayAlerts;
            Excel.Worksheet targetSheet = null;
            string targetSheetName = null;
            bool merged = false;

            using (ExcelScope.Begin(app))
            try
            {
                app.DisplayAlerts = false;

                // Create MergeSheet at the first position
                targetSheetName = GetUniqueMergeSheetName(targetWb);
                targetSheet = AddSheetAtFront(targetWb, targetSheetName);
                if (targetSheet == null) return;

                int currentTargetRow = 1;

                foreach (string filePath in selectedFiles)
                {
                    Excel.Workbook sourceWb = null;
                    Excel.Workbooks books = null;
                    try
                    {
                        books = app.Workbooks;
                        sourceWb = books.Open(filePath, UpdateLinks: false, ReadOnly: true);

                        foreach (string sourceName in ComUtil.GetSheetNames(sourceWb, true))
                        {
                            // Skip sheets that are named MergeSheet or starts with MergeSheet(
                            if (sourceName.Equals("MergeSheet", StringComparison.OrdinalIgnoreCase) ||
                                sourceName.StartsWith("MergeSheet(", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            Excel.Worksheet sourceWs = null;
                            Excel.Range usedRange = null;
                            Excel.Range rows = null;
                            Excel.Range cols = null;
                            Excel.Range anchor = null;
                            try
                            {
                                sourceWs = ComUtil.GetSheet(sourceWb, sourceName);
                                if (sourceWs == null) continue;

                                usedRange = sourceWs.UsedRange;
                                if (usedRange == null) continue;

                                rows = usedRange.Rows;
                                cols = usedRange.Columns;
                                if (rows.Count > 0 && cols.Count > 0)
                                {
                                    int rowCount = rows.Count;
                                    anchor = targetSheet.Cells[currentTargetRow, 1] as Excel.Range;
                                    usedRange.Copy(anchor);
                                    currentTargetRow += rowCount;
                                }
                            }
                            finally { Release(anchor, cols, rows, usedRange, sourceWs); }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(string.Format("Error opening or processing file {0}: {1}", System.IO.Path.GetFileName(filePath), ex.Message));
                    }
                    finally
                    {
                        if (sourceWb != null)
                        {
                            try { sourceWb.Close(SaveChanges: false); } catch { }
                        }
                        Release(sourceWb, books);
                    }
                }

                AutoFitColumns(targetSheet);
                merged = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Merge failed: " + ex.Message);
            }
            finally
            {
                try { app.CutCopyMode = (Excel.XlCutCopyMode)0; } catch { }
                try { Clipboard.Clear(); } catch { }

                app.DisplayAlerts = originalAlerts;
                Release(targetSheet, targetWb);
            }

            if (merged)
            {
                MessageBox.Show(string.Format("Workbook merge completed. Created sheet: {0}", targetSheetName));
            }
        }

        private static void AutoFitColumns(Excel.Worksheet ws)
        {
            Excel.Range columns = null;
            try
            {
                columns = ws.Columns;
                columns.AutoFit();
            }
            catch { }
            finally { Release(columns); }
        }

        private static bool SheetExists(Excel.Workbook wb, string name)
        {
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
                        if (sheet != null && sheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                    finally { Release(sheet); }
                }
            }
            finally { Release(sheets); }
            return false;
        }

        private static string GetUniqueSheetName(Excel.Workbook wb, string baseName)
        {
            if (!SheetExists(wb, baseName)) return baseName;

            int i = 2;
            while (SheetExists(wb, string.Format("{0}({1})", baseName, i)))
            {
                i++;
            }
            return string.Format("{0}({1})", baseName, i);
        }

        private static string GetUniqueMergeSheetName(Excel.Workbook wb)
        {
            return GetUniqueSheetName(wb, "MergeSheet");
        }

        public static void MergeSheets()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            // 弹出选择窗体
            System.Collections.Generic.List<string> selectedNames;
            using (frmSelectSheets selectForm = new frmSelectSheets(workbook))
            {
                if (selectForm.ShowDialog() != DialogResult.OK) return;
                selectedNames = new System.Collections.Generic.List<string>(selectForm.SelectedSheetNames);
            }

            const string targetName = "MergeSheet";
            Excel.Worksheet targetSheet = null;
            bool merged = false;

            using (ExcelScope.Begin(app))
            try
            {
                // 替换逻辑：如果已存在 MergeSheet，直接删除
                Excel.Worksheet existing = null;
                try
                {
                    existing = ComUtil.GetSheet(workbook, targetName);
                    if (existing != null)
                    {
                        app.DisplayAlerts = false;
                        existing.Delete();
                        app.DisplayAlerts = true;
                    }
                }
                finally { Release(existing); }

                // 在第1位创建新的 MergeSheet
                targetSheet = AddSheetAtFront(workbook, targetName);
                if (targetSheet == null) return;

                int currentTargetRow = 1;

                foreach (string sheetName in selectedNames)
                {
                    if (sheetName == targetName) continue;

                    Excel.Worksheet sheet = null;
                    Excel.Range usedRange = null;
                    Excel.Range rows = null;
                    Excel.Range cols = null;
                    Excel.Range anchor = null;
                    try
                    {
                        sheet = ComUtil.GetSheet(workbook, sheetName);
                        if (sheet == null) continue;

                        usedRange = sheet.UsedRange;
                        if (usedRange == null) continue;

                        rows = usedRange.Rows;
                        cols = usedRange.Columns;
                        int rowCount = rows.Count;
                        int colCount = cols.Count;

                        if (rowCount > 0 && colCount > 0)
                        {
                            anchor = targetSheet.Cells[currentTargetRow, 1] as Excel.Range;
                            usedRange.Copy(anchor);
                            currentTargetRow += rowCount;
                        }
                    }
                    finally { Release(anchor, cols, rows, usedRange, sheet); }
                }

                AutoFitColumns(targetSheet);
                targetSheet.Activate();
                merged = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Worksheet merge failed: " + ex.Message);
            }
            finally
            {
                try { app.CutCopyMode = (Excel.XlCutCopyMode)0; } catch { }
                try { Clipboard.Clear(); } catch { }

                Release(targetSheet, workbook);
            }

            if (merged)
            {
                MessageBox.Show(string.Format("Worksheet merge completed. Created/replaced sheet: {0}", targetName));
            }
        }

        
        public static void DeleteEmptyRows()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range selection = app.Selection as Excel.Range;
            if (selection == null) return;

            Excel.Range selRows = null;
            Excel.Range selCols = null;

            using (ExcelScope.Begin(app))
            try
            {
                selRows = selection.Rows;
                selCols = selection.Columns;
                int rowsCount = selRows.Count;
                int colCount = selCols.Count;

                for (int i = rowsCount; i >= 1; i--)
                {
                    Excel.Range row = null;
                    try
                    {
                        row = selRows[i] as Excel.Range;
                        if (row == null) continue;

                        bool isEmpty = true;
                        for (int j = 1; j <= colCount; j++)
                        {
                            Excel.Range cell = null;
                            try
                            {
                                cell = row.Cells[1, j] as Excel.Range;
                                if (cell != null && cell.Value != null)
                                {
                                    isEmpty = false;
                                    break;
                                }
                            }
                            finally { Release(cell); }
                        }
                        if (isEmpty)
                        {
                            row.Delete();
                        }
                    }
                    finally { Release(row); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to delete empty rows: " + ex.Message);
            }
            finally
            {
                Release(selCols, selRows, selection);
            }
        }

        
        public static void UnhideAllSheets()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            Excel.Sheets sheets = null;

            using (ExcelScope.Begin(app))
            try
            {
                sheets = workbook.Worksheets;
                int count = sheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets[i] as Excel.Worksheet;
                        if (sheet != null) sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    }
                    finally { Release(sheet); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to unhide sheets: " + ex.Message);
            }
            finally
            {
                Release(sheets, workbook);
            }
        }

        public static void FileDirectory()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook wb = app.ActiveWorkbook;
            if (wb == null) return;

            try
            {
                string path = wb.FullName;
                if (System.IO.File.Exists(path))
                {
                    System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + path + "\"");
                }
                else
                {
                    MessageBox.Show("The file has not been saved yet, so the folder cannot be opened.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open folder: " + ex.Message);
            }
            finally
            {
                Release(wb);
            }
        }
    }
}

