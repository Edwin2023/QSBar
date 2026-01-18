using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public static class SheetCommands
    {
        
        public static void CreateSheetIndex()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            try
            {
                app.ScreenUpdating = false;

                // Add new sheet at the beginning
                Excel.Worksheet indexSheet = workbook.Sheets.Add(workbook.Worksheets[1]) as Excel.Worksheet;
                if (indexSheet == null) return;
                indexSheet.Name = "INDEX";

                // Setup header
                Excel.Range headerRange = indexSheet.Range["A1:B1"];
                headerRange.Value = new object[,] { { "序号", "工作表名称" } };
                headerRange.Font.Bold = true;
                headerRange.Font.Name = "微软雅黑";
                headerRange.Font.Size = 11;
                // xlThemeColorAccent1 = 5
                headerRange.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
                headerRange.Interior.TintAndShade = 0.599993896298105;

                int rowNum = 2;
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    // Skip the index sheet itself if iterated (though we just added it)
                    if (sheet.Name == "INDEX") continue;

                    // xlSheetVisible = -1 (True in VBA, but enum value is -1)
                    // Check if visible
                    if (sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible) 
                    {
                        indexSheet.Cells[rowNum, 1].Value = rowNum - 1;
                        
                        string sheetName = sheet.Name;
                        string subAddress = $"'{sheetName}'!A1";
                        string formula = $"=HYPERLINK(\"#{subAddress}\", \"{sheetName}\")";
                        
                        indexSheet.Cells[rowNum, 2].Value = formula;
                        rowNum++;
                    }
                }

                // Formatting
                indexSheet.Columns["A:A"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                indexSheet.Columns["B:B"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                indexSheet.Columns["A:C"].AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成索引失败: " + ex.Message);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        
        public static void DeleteHyperlinks()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            app.ScreenUpdating = false;
            try
            {
                Excel.Range selection = app.Selection as Excel.Range;
                if (selection != null)
                {
                    selection.Hyperlinks.Delete();
                }
                else
                {
                    foreach (Excel.Worksheet sheet in app.ActiveWorkbook.Worksheets)
                    {
                        sheet.UsedRange.Hyperlinks.Delete();
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        
        public static void MergeSheets()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            try
            {
                app.ScreenUpdating = false;
                Excel.Worksheet targetSheet = workbook.Worksheets.Add() as Excel.Worksheet;
                if (targetSheet == null) return;
                targetSheet.Name = "Merged_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                int currentTargetRow = 1;

                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == targetSheet.Name) continue;
                    if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible) continue;

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
                MessageBox.Show("合并工作表完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("合并工作表失败: " + ex.Message);
            }
            finally
            {
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
                MessageBox.Show("删除空行失败: " + ex.Message);
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
                MessageBox.Show("取消隐藏失败: " + ex.Message);
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
                    MessageBox.Show("文件尚未保存，无法打开目录。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开目录失败: " + ex.Message);
            }
        }
    }
}

