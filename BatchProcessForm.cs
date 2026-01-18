namespace QSBar
{
    public static class SheetCommands
    {
        
        public static void CreateSheetIndex()
        {
            dynamic app = WpsExcelAddIn.App;
            dynamic workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            try
            {
                app.ScreenUpdating = false;

                // Add new sheet at the beginning
                dynamic indexSheet = workbook.Sheets.Add(Before: workbook.Worksheets[1]);
                indexSheet.Name = "INDEX";

                // Setup header
                dynamic headerRange = indexSheet.Range["A1:B1"];
                headerRange.Value = new object[,] { { "序号", "工作表名称" } };
                headerRange.Font.Bold = true;
                headerRange.Font.Name = "微软雅黑";
                headerRange.Font.Size = 11;
                // xlThemeColorAccent1 = 5
                headerRange.Interior.ThemeColor = 5;
                headerRange.Interior.TintAndShade = 0.599993896298105;

                int rowNum = 2;
                foreach (dynamic sheet in workbook.Worksheets)
                {
                    // Skip the index sheet itself if iterated (though we just added it)
                    if (sheet.Name == "INDEX") continue;

                    // xlSheetVisible = -1 (True in VBA, but enum value is -1)
                    // Check if visible
                    if (sheet.Visible == -1) 
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
                indexSheet.Columns["A:A"].HorizontalAlignment = -4108; // xlCenter
                indexSheet.Columns["B:B"].HorizontalAlignment = -4131; // xlLeft
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
            dynamic app = WpsExcelAddIn.App;
            app.ScreenUpdating = false;
            try
            {
                dynamic selection = app.Selection;
                if (selection != null)
                {
                    selection.Hyperlinks.Delete();
                }
                else
                {
                    foreach (dynamic sheet in app.ActiveWorkbook.Worksheets)
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
            dynamic app = WpsExcelAddIn.App;
            dynamic workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            try
            {
                app.ScreenUpdating = false;
                dynamic targetSheet = workbook.Worksheets.Add();
                targetSheet.Name = "Merged_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                int currentTargetRow = 1;

                foreach (dynamic sheet in workbook.Worksheets)
                {
                    if (sheet.Name == targetSheet.Name) continue;
                    if (sheet.Visible != -1) continue;

                    dynamic usedRange = sheet.UsedRange;
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
            dynamic app = WpsExcelAddIn.App;
            dynamic selection = app.Selection;
            if (selection == null) return;

            try
            {
                app.ScreenUpdating = false;
                int rowsCount = selection.Rows.Count;
                int startRow = selection.Row;
                int colCount = selection.Columns.Count;
                int startCol = selection.Column;

                for (int i = rowsCount; i >= 1; i--)
                {
                    dynamic row = selection.Rows[i];
                    bool isEmpty = true;
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (row.Cells[1, j].Value != null)
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
            dynamic app = WpsExcelAddIn.App;
            dynamic workbook = app.ActiveWorkbook;
            if (workbook == null) return;

            try
            {
                app.ScreenUpdating = false;
                foreach (dynamic sheet in workbook.Worksheets)
                {
                    sheet.Visible = -1; // xlSheetVisible
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
    }
}
