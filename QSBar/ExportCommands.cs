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
                string fileName = string.Format("(OUT{0}){1}", dateStr, sheetName);
                
                // Sanitize filename
                fileName = fileName.Replace(":", "-").Replace("\\", "-").Replace("/", "-");
                if (fileName.Length > 50) fileName = fileName.Substring(0, 50);
                fileName += ".xlsx";

                // Path Processing (Same as Workbook)
                string savePath = app.ActiveWorkbook.Path;
                if (string.IsNullOrEmpty(savePath))
                {
                    savePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }
                string fullPath = Path.Combine(savePath, fileName);

                // Cancel auto-filter on source before copy to avoid filter interference
                if (activeSheet.AutoFilterMode)
                {
                    activeSheet.AutoFilterMode = false;
                }

                // Copy Sheet (creates new workbook)
                activeSheet.Copy();
                Excel.Workbook newWb = app.ActiveWorkbook;
                Excel.Worksheet targetSheet = newWb.Sheets[1] as Excel.Worksheet;

                // Cancel auto-filter on target (defensive, in case copy preserved filter state)
                if (targetSheet.AutoFilterMode)
                {
                    targetSheet.AutoFilterMode = false;
                }

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
                MessageBox.Show("Export failed: " + ex.Message);
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
                    // Cancel auto-filter before copy to ensure all rows are included
                    if (ws.AutoFilterMode)
                    {
                        try { ws.AutoFilterMode = false; } catch { }
                    }

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
                MessageBox.Show("Conversion failed: " + ex.Message);
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

            if (app == null)

            {

                MessageBox.Show("QSBar add-in is not loaded. Please restart Excel and ensure the add-in is enabled.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return;

            }

            dynamic sourceWb = app.ActiveWorkbook;

            if (sourceWb == null)

            {

                MessageBox.Show("No workbook is open.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return;

            }



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

                string prefix = isInternal ? string.Format("(OUT-Internal-{0})", dateStr) : string.Format("(OUT{0})", dateStr);

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

                app.StatusBar = "导出报表: 展开所有行...";

                ExpandAllRows(newWb);



                app.StatusBar = "导出报表: 取消筛选...";

                CancelAutoFilters(newWb);



                if (!isInternal)

                {

                    // Standard Report: selective conversion — only convert formulas

                    // referencing hidden columns or problematic defined names

                    app.StatusBar = "导出报表: 识别非一级列...";

                    HideNonLevel1Columns(newWb);



                    ConvertReferencesToHiddenColsAndNames(app, newWb);



                    app.StatusBar = "导出报表: 清理错误名称...";

                    ClearErrorNames(newWb);



                    app.StatusBar = "导出报表: 删除隐藏列...";

                    DeleteHiddenColumns(newWb);

                }

                else

                {

                    // Internal Report: full paste-as-values

                    app.StatusBar = "导出报表: 清理错误名称...";

                    ClearErrorNames(newWb);



                    app.StatusBar = "导出报表: 全量数值化...";

                    ConvertAllToValuesInternal(newWb);

                }



                app.StatusBar = "导出报表: 清理打印区域外内容...";

                CleanOutsidePrintArea(newWb);



                // Save

                app.StatusBar = "导出报表: 保存文件...";

                if (File.Exists(fullPath))

                {

                    try { File.Delete(fullPath); } catch { }

                }



                newWb.SaveAs(fullPath, 51); // xlOpenXMLWorkbook

                app.StatusBar = "导出报表: 完成";

            }

            catch (Exception ex)

            {

                app.StatusBar = false;

                MessageBox.Show("Export failed: " + ex.Message);

            }

            finally
            {
                app.StatusBar = false;

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

                if (ws.AutoFilterMode)

                {

                    try { ws.AutoFilterMode = false; } catch { }

                }

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



        /// <summary>

        /// 两阶段选择性数值化：

        /// 阶段1：只粘死自身公式含外部工作簿引用 [xxx.xlsx] 或问题名称的单元格。

        /// 阶段2：粘死引用了隐藏列（将被删除）的公式单元格。

        /// 阶段1先执行，外部引用粘死后阶段2只检查隐藏列，表内公式（如=2*D2）不受影响。

        /// </summary>

        private static void ConvertReferencesToHiddenColsAndNames(dynamic app, dynamic wb)

        {

            // 收集即将被 ClearErrorNames 删除的问题名称

            var problemNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try

            {

                foreach (dynamic name in wb.Names)

                {

                    string val = "";

                    try { val = name.Value; } catch { continue; }

                    if (val.Contains("#REF!") || val.Contains(":\\") || val.Contains("\\\\") || val.Contains("#N/A"))

                    {

                        problemNames.Add(name.Name);

                    }

                }

            }

            catch { }



            var problemNameRegex = BuildProblemNameRegex(problemNames);



            // 外部工作簿引用模式: 匹配所有 [xxx] 格式（含 [141]Laldia清单 等无扩展名格式）

            var externalRefRegex = new System.Text.RegularExpressions.Regex(

                @"\[[^\]]*\]",

                System.Text.RegularExpressions.RegexOptions.Compiled |

                System.Text.RegularExpressions.RegexOptions.IgnoreCase);



            int totalSheets = wb.Worksheets.Count;

            int sheetIndex = 0;



            // ===== 阶段1：只粘死自身公式含外部链接的单元格 =====

            foreach (dynamic ws in wb.Worksheets)

            {

                sheetIndex++;

                app.StatusBar = string.Format("导出报表: 粘死外部引用 ({0}/{1}) {2}...",

                    sheetIndex, totalSheets, ws.Name);



                try

                {

                    dynamic usedRange = ws.UsedRange;

                    if (usedRange == null) continue;



                    if (problemNames.Count == 0 && !HasAnyExternalRef(wb))

                        continue;



                    dynamic formulaArray = usedRange.Formula;

                    int rows = formulaArray.GetLength(0);

                    int cols = formulaArray.GetLength(1);

                    int startRow = usedRange.Row;

                    int startCol = usedRange.Column;



                    for (int r = 1; r <= rows; r++)

                    {

                        for (int c = 1; c <= cols; c++)

                        {

                            object cellObj = formulaArray[r, c];

                            string formula = cellObj as string;

                            if (string.IsNullOrEmpty(formula))

                                continue;

                            bool isArray = formula.StartsWith("{=");

                            if (!formula.StartsWith("=") && !isArray)

                                continue;

                            if (isArray)

                                formula = formula.Substring(1, formula.Length - 2);



                            bool needsConversion = false;



                            // 名称匹配前先剥离引号字面量（工作表名 'xxx'、文本常量 "xxx"）。

                            // 否则表名里的字母会撞上单字母坏名称，把正常公式一起粘死。

                            string scrubbed = QuotedLiteralRegex.Replace(formula, "''");



                            // 只检查问题名称。必须是完整标识符，不能用裸子串匹配

                            if (problemNameRegex != null && problemNameRegex.IsMatch(scrubbed))

                            {

                                needsConversion = true;

                            }



                            // 只检查外部工作簿引用（公式文本里带 [xxx.xlsx]）

                            // 注意用原始 formula：路径式外部引用 'C:\x\[a.xlsx]Sheet'!A1 的方括号在引号内

                            if (!needsConversion && externalRefRegex.IsMatch(formula))

                            {

                                needsConversion = true;

                            }



                            if (needsConversion)

                            {

                                try

                                {

                                    int rowIdx = startRow + r - 1;

                                    int colIdx = startCol + c - 1;

                                    dynamic cell = ws.Cells[rowIdx, colIdx];

                                    if (isArray)

                                    {

                                        dynamic arr = cell.CurrentArray;

                                        arr.Value = arr.Value;

                                    }

                                    else

                                    {

                                        cell.Value = cell.Value;

                                    }

                                }

                                catch { }

                            }

                        }

                    }

                }

                catch { }

            }



            // ===== 阶段2：粘死引用了隐藏列的公式（隐藏列将在步骤7被删除） =====

            sheetIndex = 0;

            foreach (dynamic ws in wb.Worksheets)

            {

                sheetIndex++;

                app.StatusBar = string.Format("导出报表: 粘死隐藏列引用 ({0}/{1}) {2}...",

                    sheetIndex, totalSheets, ws.Name);



                try

                {

                    dynamic usedRange = ws.UsedRange;

                    if (usedRange == null) continue;



                    // 获取当前工作表被隐藏的列字母集合

                    var hiddenColLetters = new HashSet<string>();

                    int firstCol = usedRange.Column;

                    int lastCol = firstCol + usedRange.Columns.Count - 1;

                    for (int c = firstCol; c <= lastCol; c++)

                    {

                        try

                        {

                            if (ws.Columns[c].Hidden)

                                hiddenColLetters.Add(ColumnNumberToLetter(c));

                        }

                        catch { }

                    }



                    if (hiddenColLetters.Count == 0)

                        continue;



                    dynamic formulaArray = usedRange.Formula;

                    int rows = formulaArray.GetLength(0);

                    int cols = formulaArray.GetLength(1);

                    int startRow = usedRange.Row;

                    int startCol = usedRange.Column;



                    var cellRefPattern = new System.Text.RegularExpressions.Regex(

                        @"\$?([A-Z]{1,3})\$?\d+",

                        System.Text.RegularExpressions.RegexOptions.Compiled);



                    for (int r = 1; r <= rows; r++)

                    {

                        for (int c = 1; c <= cols; c++)

                        {

                            object cellObj = formulaArray[r, c];

                            string formula = cellObj as string;

                            if (string.IsNullOrEmpty(formula))

                                continue;

                            bool isArray = formula.StartsWith("{=");

                            if (!formula.StartsWith("=") && !isArray)

                                continue;

                            if (isArray)

                                formula = formula.Substring(1, formula.Length - 2);



                            // 只检查是否引用了隐藏列（跳过跨表引用如 ='Sheet'!A1）

                            var matches = cellRefPattern.Matches(formula);

                            bool needsConversion = false;

                            foreach (System.Text.RegularExpressions.Match match in matches)

                            {

                                if (match.Index > 0 && formula[match.Index - 1] == '!')

                                    continue;

                                string colLetter = match.Groups[1].Value;

                                if (hiddenColLetters.Contains(colLetter))

                                {

                                    needsConversion = true;

                                    break;

                                }

                            }



                            if (needsConversion)

                            {

                                try

                                {

                                    int rowIdx = startRow + r - 1;

                                    int colIdx = startCol + c - 1;

                                    dynamic cell = ws.Cells[rowIdx, colIdx];

                                    if (isArray)

                                    {

                                        dynamic arr = cell.CurrentArray;

                                        arr.Value = arr.Value;

                                    }

                                    else

                                    {

                                        cell.Value = cell.Value;

                                    }

                                }

                                catch { }

                            }

                        }

                    }

                }

                catch { }

            }

        }




        private static bool HasAnyExternalRef(dynamic wb)

        {

            try

            {

                foreach (dynamic link in wb.LinkSources(1)) // xlExcelLinks

                {

                    if (link != null) return true;

                }

            }

            catch { }

            return false;

        }



        // Excel 公式里的引号字面量：工作表名 'xxx'（内部单引号转义为 ''）与文本常量 "xxx"

        private static readonly System.Text.RegularExpressions.Regex QuotedLiteralRegex =

            new System.Text.RegularExpressions.Regex(

                @"'(?:[^']|'')*'|""(?:[^""]|"""")*""",

                System.Text.RegularExpressions.RegexOptions.Compiled);



        /// <summary>

        /// 把全部问题名称合并成一个带标识符边界的正则。

        /// 名称合法字符为字母数字下划线点反斜杠（\w 已覆盖中日韩），前后紧邻这些字符时不算名称引用。

        /// 合并成单个正则是为了避免逐名称匹配在大表上退化成 名称数 x 公式数 次扫描。

        /// </summary>

        private static System.Text.RegularExpressions.Regex BuildProblemNameRegex(HashSet<string> names)

        {

            if (names == null || names.Count == 0) return null;



            List<string> ordered = new List<string>(names);

            // 长的排前面，避免交替分支里短名称抢先匹配

            ordered.Sort(delegate(string a, string b) { return b.Length.CompareTo(a.Length); });



            System.Text.StringBuilder sb = new System.Text.StringBuilder();

            sb.Append(@"(?<![\w.\\])(?:");

            for (int i = 0; i < ordered.Count; i++)

            {

                if (i > 0) sb.Append('|');

                sb.Append(System.Text.RegularExpressions.Regex.Escape(ordered[i]));

            }

            sb.Append(@")(?![\w.\\])");



            try

            {

                return new System.Text.RegularExpressions.Regex(

                    sb.ToString(),

                    System.Text.RegularExpressions.RegexOptions.Compiled |

                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            }

            catch

            {

                return null;

            }

        }



        private static string ColumnNumberToLetter(int col)

        {

            string result = "";

            while (col > 0)

            {

                col--;

                result = (char)('A' + (col % 26)) + result;

                col /= 26;

            }

            return result;

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

