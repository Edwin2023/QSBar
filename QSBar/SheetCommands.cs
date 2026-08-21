using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public static class SheetCommands
    {
        
        private static void ApplyModernStyle(Excel.Worksheet ws, int totalRows, int totalCols)
        {
            if (totalRows < 1 || totalCols < 1) return;

            Excel.Range fullRange = ws.Range[ws.Cells[1, 1], ws.Cells[totalRows, totalCols]];
            Excel.Range headerRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, totalCols]];
            Excel.Range dataRange = null;
            if (totalRows > 1)
            {
                dataRange = ws.Range[ws.Cells[2, 1], ws.Cells[totalRows, totalCols]];
            }

            // 1. 基础字体设置
            fullRange.Font.Name = "微软雅黑";
            fullRange.Font.Size = 10;
            fullRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            // 2. 表头样式：深蓝色背景 + 白色粗体
            headerRange.RowHeight = 30;
            headerRange.Font.Size = 11;
            headerRange.Font.Bold = true;
            headerRange.Font.Color = ColorTranslator.ToOle(Color.White);
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(44, 62, 80)); // 优雅的深灰蓝
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // 3. 数据行样式
            if (dataRange != null)
            {
                dataRange.RowHeight = 22;
                
                // 隔行变色 (Banding) - 浅灰蓝色
                for (int i = 2; i <= totalRows; i += 2)
                {
                    ws.Rows[i].Range[ws.Cells[1, 1], ws.Cells[1, totalCols]].Interior.Color = 
                        ColorTranslator.ToOle(Color.FromArgb(242, 244, 247));
                }

                // 序号列居中
                ws.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }

            // 4. 边框设置：极细浅灰色边框
            fullRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            fullRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            fullRange.Borders.Color = ColorTranslator.ToOle(Color.FromArgb(218, 223, 225));

            // 5. 冻结首行
            ws.Application.ActiveWindow.SplitColumn = 0;
            ws.Application.ActiveWindow.SplitRow = 1;
            ws.Application.ActiveWindow.FreezePanes = true;

            // 6. 自动列宽并留出一点边距
            fullRange.Columns.AutoFit();
            for (int i = 1; i <= totalCols; i++)
            {
                double currentWidth = (double)ws.Columns[i].ColumnWidth;
                ws.Columns[i].ColumnWidth = currentWidth + 2;
            }

            // 7. 隐藏网格线，让界面更清爽
            ws.Application.ActiveWindow.DisplayGridlines = false;
        }

        public static void CreateSheetIndex()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            try
            {
                app.ScreenUpdating = false;

                Excel.Worksheet indexSheet = workbook.Worksheets.Add(Before: workbook.Worksheets[1]) as Excel.Worksheet;
                if (indexSheet == null) return;
                indexSheet.Name = "Index_" + DateTime.Now.ToString("HHmmss");

                // 准备数据
                int rowNum = 1;
                indexSheet.Cells[rowNum, 1].Value = "No.";
                indexSheet.Cells[rowNum, 2].Value = "Worksheet Name";
                rowNum++;

                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == indexSheet.Name) continue;
                    if (sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible) 
                    {
                        indexSheet.Cells[rowNum, 1].Value = rowNum - 1;
                        string sheetName = sheet.Name;
                        string formula = string.Format("=HYPERLINK(\"#'{0}'!A1\", \"{1}\")", sheetName.Replace("'", "''"), sheetName);
                        indexSheet.Cells[rowNum, 2].Formula = formula;
                        rowNum++;
                    }
                }

                // 应用美化样式
                ApplyModernStyle(indexSheet, rowNum - 1, 2);
                indexSheet.Activate();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to generate the sheet index: " + ex.Message);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
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

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select files to index";
            openFileDialog.Filter = "All Files|*.*";
            openFileDialog.Multiselect = true;
            if (!string.IsNullOrEmpty(initialDir))
            {
                openFileDialog.InitialDirectory = initialDir;
            }

            if (openFileDialog.ShowDialog() != DialogResult.OK) return;

            try
            {
                app.ScreenUpdating = false;

                string sheetName = "File Index";
                Excel.Worksheet indexSheet = null;
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        indexSheet = ws;
                        indexSheet.Cells.Clear();
                        break;
                    }
                }

                if (indexSheet == null)
                {
                    indexSheet = workbook.Worksheets.Add(Before: workbook.Worksheets[1]) as Excel.Worksheet;
                    indexSheet.Name = sheetName;
                }

                // 准备数据
                int rowNum = 1;
                indexSheet.Cells[rowNum, 1].Value = "No.";
                indexSheet.Cells[rowNum, 2].Value = "File Name";
                indexSheet.Cells[rowNum, 3].Value = "Full Path";
                rowNum++;

                foreach (string filePath in openFileDialog.FileNames)
                {
                    string fileName = System.IO.Path.GetFileName(filePath);
                    indexSheet.Cells[rowNum, 1].Value = rowNum - 1;
                    
                    string formula = string.Format("=HYPERLINK(\"{0}\", \"{1}\")", filePath, fileName);
                    indexSheet.Cells[rowNum, 2].Formula = formula;
                    indexSheet.Cells[rowNum, 3].Value = filePath;
                    rowNum++;
                }

                // 应用美化样式
                ApplyModernStyle(indexSheet, rowNum - 1, 3);
                indexSheet.Activate();
                
                MessageBox.Show(string.Format("File index created successfully. Total files: {0}", openFileDialog.FileNames.Length));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to generate the file index: " + ex.Message);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        
        public static void MergeWorkbooks()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook targetWb = app.ActiveWorkbook;
            if (targetWb == null) return;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select Excel files to merge";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb|All Files|*.*";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() != DialogResult.OK) return;

            Excel.XlCalculation originalCalc = app.Calculation;
            bool originalUpdating = app.ScreenUpdating;
            bool originalEvents = app.EnableEvents;
            bool originalAlerts = app.DisplayAlerts;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                // Create MergeSheet at the first position
                string targetSheetName = GetUniqueMergeSheetName(targetWb);
                Excel.Worksheet targetSheet = targetWb.Worksheets.Add(Before: targetWb.Worksheets[1]) as Excel.Worksheet;
                if (targetSheet == null) return;
                targetSheet.Name = targetSheetName;

                int currentTargetRow = 1;

                foreach (string filePath in openFileDialog.FileNames)
                {
                    Excel.Workbook sourceWb = null;
                    try
                    {
                        sourceWb = app.Workbooks.Open(filePath, UpdateLinks: false, ReadOnly: true);
                        foreach (Excel.Worksheet sourceWs in sourceWb.Worksheets)
                        {
                            // Skip hidden sheets
                            if (sourceWs.Visible != Excel.XlSheetVisibility.xlSheetVisible) continue;

                            // Skip sheets that are named MergeSheet or starts with MergeSheet(
                            if (sourceWs.Name.Equals("MergeSheet", StringComparison.OrdinalIgnoreCase) || 
                                sourceWs.Name.StartsWith("MergeSheet(", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            Excel.Range usedRange = sourceWs.UsedRange;
                            if (usedRange != null && usedRange.Rows.Count > 0 && usedRange.Columns.Count > 0)
                            {
                                int rowCount = usedRange.Rows.Count;
                                usedRange.Copy(targetSheet.Cells[currentTargetRow, 1]);
                                currentTargetRow += rowCount;
                            }
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
                            sourceWb.Close(SaveChanges: false);
                            Marshal.ReleaseComObject(sourceWb);
                        }
                    }
                }

                targetSheet.Columns.AutoFit();
                MessageBox.Show(string.Format("Workbook merge completed. Created sheet: {0}", targetSheetName));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Merge failed: " + ex.Message);
            }
            finally
            {
                try { app.CutCopyMode = (Excel.XlCutCopyMode)0; } catch { }
                try { Clipboard.Clear(); } catch { }

                app.ScreenUpdating = originalUpdating;
                app.Calculation = originalCalc;
                app.EnableEvents = originalEvents;
                app.DisplayAlerts = originalAlerts;
            }
        }

        private static bool SheetExists(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static string GetUniqueMergeSheetName(Excel.Workbook wb)
        {
            string baseName = "MergeSheet";
            if (!SheetExists(wb, baseName)) return baseName;

            int i = 2;
            while (SheetExists(wb, string.Format("{0}({1})", baseName, i)))
            {
                i++;
            }
            return string.Format("{0}({1})", baseName, i);
        }

        public static void MergeSheets()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            // 弹出选择窗体
            frmSelectSheets selectForm = new frmSelectSheets(workbook);
            if (selectForm.ShowDialog() != DialogResult.OK) return;
            var selectedNames = selectForm.SelectedSheetNames;

            try
            {
                app.ScreenUpdating = false;

                // 替换逻辑：如果已存在 MergeSheet，直接删除
                string targetName = "MergeSheet";
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                    {
                        app.DisplayAlerts = false;
                        ws.Delete();
                        app.DisplayAlerts = true;
                        break;
                    }
                }

                // 在第1位创建新的 MergeSheet
                Excel.Worksheet targetSheet = workbook.Worksheets.Add(Before: workbook.Worksheets[1]) as Excel.Worksheet;
                if (targetSheet == null) return;
                targetSheet.Name = targetName;

                int currentTargetRow = 1;

                foreach (string sheetName in selectedNames)
                {
                    Excel.Worksheet sheet = util.getSheetByEqual(workbook, sheetName);
                    if (sheet == null || sheet.Name == targetSheet.Name) continue;

                    Excel.Range usedRange = sheet.UsedRange;
                    if (usedRange == null) continue;

                    int rowCount = usedRange.Rows.Count;
                    int colCount = usedRange.Columns.Count;

                    if (rowCount > 0 && colCount > 0)
                    {
                        usedRange.Copy(targetSheet.Cells[currentTargetRow, 1]);
                        currentTargetRow += rowCount;
                    }
                }

                targetSheet.Columns.AutoFit();
                targetSheet.Activate();
                MessageBox.Show(string.Format("Worksheet merge completed. Created/replaced sheet: {0}", targetName));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Worksheet merge failed: " + ex.Message);
            }
            finally
            {
                try { app.CutCopyMode = (Excel.XlCutCopyMode)0; } catch { }
                try { Clipboard.Clear(); } catch { }
                app.ScreenUpdating = true;
            }
        }

        
        public static void DeleteEmptyRows()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Range selection = app.Selection as Excel.Range;
            if (selection == null) return;

            try
            {
                app.ScreenUpdating = false;
                int rowsCount = selection.Rows.Count;
                int colCount = selection.Columns.Count;

                for (int i = rowsCount; i >= 1; i--)
                {
                    Excel.Range row = selection.Rows[i] as Excel.Range;
                    if (row == null) continue;
                    
                    bool isEmpty = true;
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (((Excel.Range)row.Cells[1, j]).Value != null)
                        {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty)
                    {
                        row.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to delete empty rows: " + ex.Message);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        
        public static void UnhideAllSheets()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            try
            {
                app.ScreenUpdating = false;
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to unhide sheets: " + ex.Message);
            }
            finally
            {
                app.ScreenUpdating = true;
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
        }
    }
}

