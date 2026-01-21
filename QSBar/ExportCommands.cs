using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public static class ExportCommands
    {
        
        public static void ShowLevel1() { SetRowLevel(1); }
        public static void ShowLevel2() { SetRowLevel(2); }
        public static void ShowLevel3() { SetRowLevel(3); }
        public static void ShowLevel4() { SetRowLevel(4); }

        private static void SetRowLevel(int level)
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Worksheet ws = app.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            try { ws.Outline.ShowLevels(RowLevels: level); } catch { }
        }

        public static void ExportCurrentSheet()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            
            Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
            if (activeSheet == null) return;

            // Preserve settings
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

                // File Name Processing
                string dateStr = DateTime.Now.ToString("yyyy-MM-dd");
                string sheetName = activeSheet.Name;
                string fileName = $"(OUT{dateStr}){sheetName}";
                
                // Sanitize filename
                fileName = fileName.Replace(":", "-").Replace("\\", "-").Replace("/", "-");
                if (fileName.Length > 50) fileName = fileName.Substring(0, 50);
                fileName += ".xlsx";

                // Path Processing (Desktop)
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (!savePath.EndsWith("\\")) savePath += "\\";
                string fullPath = Path.Combine(savePath, fileName);

                // Copy Sheet (creates new workbook)
                activeSheet.Copy();
                Excel.Workbook newWb = app.ActiveWorkbook;
                Excel.Worksheet targetSheet = newWb.Sheets[1] as Excel.Worksheet;

                // Paste Values & Formats
                Excel.Range usedRange = targetSheet.UsedRange;
                usedRange.Copy();
                usedRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                usedRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                app.CutCopyMode = (Excel.XlCutCopyMode)0; // xlCopy (0) or false

                // Delete existing file if any
                if (File.Exists(fullPath))
                {
                    try { File.Delete(fullPath); } catch { }
                }

                // Save
                newWb.SaveAs(fullPath, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出失败: " + ex.Message);
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

        
        public static void ConvertAllToValues()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            
            Excel.Workbook wb = app.ActiveWorkbook;
            if (wb == null) return;
            
            bool originalUpdating = app.ScreenUpdating;
            try
            {
                app.ScreenUpdating = false;
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    Excel.Range usedRange = ws.UsedRange;
                    if (usedRange != null)
                    {
                        try
                        {
                            usedRange.Copy();
                            usedRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("转换失败: " + ex.Message);
            }
            finally
            {
                try { app.CutCopyMode = (Excel.XlCutCopyMode)0; } catch { }
                try { Clipboard.Clear(); } catch { }
                app.ScreenUpdating = originalUpdating;
            }
        }



        

        public static void ExportStandardReport()

        {

            PerformExport(isInternal: false);

        }



        

        public static void ExportInternalReport()

        {

            PerformExport(isInternal: true);

        }



        private static void PerformExport(bool isInternal)

        {

            dynamic app = WpsExcelAddIn.App;

            dynamic sourceWb = app.ActiveWorkbook;

            

            XlCalculation originalCalc = (XlCalculation)app.Calculation;

            bool originalUpdating = app.ScreenUpdating;

            bool originalEvents = app.EnableEvents;

            bool originalAlerts = app.DisplayAlerts;



            try

            {

                app.ScreenUpdating = false;

                app.Calculation = XlCalculation.xlCalculationManual;

                app.EnableEvents = false;

                app.DisplayAlerts = false;



                // Prepare Filename

                string dateStr = DateTime.Now.ToString("yyyy-MM-dd"); // Replaced / with -

                string prefix = isInternal ? $"(OUT内部全数据{dateStr})" : $"(OUT{dateStr})";

                string sourcePath = sourceWb.Path;

                string sourceName = sourceWb.Name;

                

                // Strip extension

                int dotIndex = sourceName.LastIndexOf('.');

                if (dotIndex > 0) sourceName = sourceName.Substring(0, dotIndex);



                string newFileName = prefix + sourceName + ".xlsx";

                string fullPath = Path.Combine(sourcePath, newFileName);



                // Copy Workbook logic:

                // VBA: Sheets.Copy -> Copies all sheets to new workbook

                sourceWb.Sheets.Copy(); 

                dynamic newWb = app.ActiveWorkbook;



                // Process New Workbook

                ExpandAllRows(newWb);

                CancelAutoFilters(newWb);

                ClearErrorNames(newWb);

                

                // Convert to values

                ConvertAllToValuesInternal(newWb);



                if (!isInternal)

                {

                    // Additional steps for Standard Report

                    HideNonLevel1Columns(newWb);

                    DeleteHiddenColumns(newWb);

                }



                CleanOutsidePrintArea(newWb);



                // Save

                if (File.Exists(fullPath))

                {

                    try { File.Delete(fullPath); } catch { }

                }



                newWb.SaveAs(fullPath, 51); // xlOpenXMLWorkbook

            }

            catch (Exception ex)

            {

                MessageBox.Show("导出失败: " + ex.Message);

            }

            finally
            {
                try { app.CutCopyMode = false; } catch { }
                try { Clipboard.Clear(); } catch { }

                app.ScreenUpdating = originalUpdating;
                app.Calculation = originalCalc;
                app.EnableEvents = originalEvents;
                app.DisplayAlerts = originalAlerts;
            }

        }



        private static void ConvertAllToValuesInternal(dynamic wb)

        {

            dynamic app = WpsExcelAddIn.App;

            foreach (dynamic ws in wb.Worksheets)

            {

                app.CutCopyMode = false;

                dynamic usedRange = ws.UsedRange;

                if (usedRange != null)

                {

                    try

                    {

                        usedRange.Copy();

                        usedRange.PasteSpecial(-4163); // xlPasteValues

                    }

                    catch { }

                }

            }

            app.CutCopyMode = false;

        }



        private static void CancelAutoFilters(dynamic wb)

        {

            foreach (dynamic ws in wb.Worksheets)

            {

                if (ws.AutoFilterMode)

                {

                    ws.AutoFilterMode = false;

                }

            }

        }



        private static void ClearErrorNames(dynamic wb)

        {

            try

            {

                List<string> namesToDelete = new List<string>();

                foreach (dynamic name in wb.Names)

                {

                    string val = "";

                    try { val = name.Value; } catch { continue; }

                    

                    if (val.Contains("#REF!") || val.Contains(":\\") || val.Contains("\\\\") || val.Contains("#N/A"))

                    {

                        namesToDelete.Add(name.Name);

                    }

                }



                foreach (string n in namesToDelete)

                {

                    try { wb.Names.Item(n).Delete(); } catch { }

                }

            }

            catch { }

        }



        private static void HideNonLevel1Columns(dynamic wb)

        {

            foreach (dynamic ws in wb.Worksheets)

            {

                try { ws.Outline.ShowLevels(ColumnLevels: 1); } catch { }

            }

        }



        private static void ExpandAllRows(dynamic wb)

        {

            foreach (dynamic ws in wb.Worksheets)

            {

                try { ws.Outline.ShowLevels(RowLevels: 4); } catch { }

            }

        }



        private static void DeleteHiddenColumns(dynamic wb)

        {

            foreach (dynamic ws in wb.Worksheets)

            {

                try

                {

                    dynamic usedRange = ws.UsedRange;

                    int colCount = usedRange.Columns.Count;

                    // Loop backwards is safer for deletion, but VBA used a Do While loop with Counter adjustment

                    

                    for (int i = colCount; i >= 1; i--)

                    {

                        dynamic col = usedRange.Columns[i];

                        if (col.EntireColumn.Hidden)

                        {

                            col.EntireColumn.Delete();

                        }

                    }

                }

                catch { }

            }

        }



        private static void CleanOutsidePrintArea(dynamic wb)

        {

            dynamic app = WpsExcelAddIn.App;

            foreach (dynamic ws in wb.Worksheets)

            {

                try

                {

                    dynamic printRange = null;

                    // 1. Find Print_Area

                    foreach (dynamic name in ws.Names)

                    {

                        if (name.Name.EndsWith("!Print_Area"))

                        {

                            printRange = name.RefersToRange;

                            break;

                        }

                    }



                    // Fallback to UsedRange if no print area (creating a temporary print area name logic from VBA seems redundant if we just use UsedRange, but let's follow logic: "Use UsedRange as PrintArea if none")

                    if (printRange == null)

                    {

                        printRange = ws.UsedRange;

                    }



                    dynamic usedRange = ws.UsedRange;

                    

                    // 2.1 Delete Columns outside

                    dynamic deleteCols = null;

                    foreach (dynamic col in usedRange.Columns)

                    {

                        dynamic intersect = app.Intersect(col, printRange);

                        if (intersect == null)

                        {

                            if (deleteCols == null) deleteCols = col;

                            else deleteCols = app.Union(deleteCols, col);

                        }

                    }

                    if (deleteCols != null) deleteCols.EntireColumn.Delete();



                    // 2.2 Delete Rows outside

                    dynamic deleteRows = null;

                    foreach (dynamic row in usedRange.Rows)

                    {

                        dynamic intersect = app.Intersect(row, printRange);

                        if (intersect == null)

                        {

                            if (deleteRows == null) deleteRows = row;

                            else deleteRows = app.Union(deleteRows, row);

                        }

                    }

                    if (deleteRows != null) deleteRows.EntireRow.Delete();

                }

                catch { }

            }

        }



        private enum XlCalculation

        {

            xlCalculationAutomatic = -4105,

            xlCalculationManual = -4135,

            xlCalculationSemiautomatic = 2

        }

    }

}

