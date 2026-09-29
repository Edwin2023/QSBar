using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public static class ExportCommands
    {
        /// <summary>
        /// 先把表名取成字符串，循环里再按名字拿表、用完就还。
        /// 直接 foreach 工作表集合的话，每张表都留下一个不还的引用，Excel 退不掉进程。
        /// </summary>
        private static List<string> SheetNames(dynamic wb)
        {
            var names = new List<string>();
            dynamic sheets = wb.Worksheets;
            try
            {
                int count = sheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic ws = sheets[i];
                    try { names.Add((string)ws.Name); }
                    catch { }
                    finally { ComUtil.Release((object)ws); }
                }
            }
            finally { ComUtil.Release((object)sheets); }
            return names;
        }

        // 调用方负责 Release
        private static dynamic GetSheet(dynamic wb, string name)
        {
            dynamic sheets = null;
            try
            {
                sheets = wb.Worksheets;
                return sheets[name];
            }
            catch { return null; }
            finally { ComUtil.Release((object)sheets); }
        }

        public static void ShowLevel1() { SetRowLevel(1); }
        public static void ShowLevel2() { SetRowLevel(2); }
        public static void ShowLevel3() { SetRowLevel(3); }
        public static void ShowLevel4() { SetRowLevel(4); }

        private static void SetRowLevel(int level)
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;

            Excel.Worksheet ws = null;
            Excel.Outline outline = null;
            try
            {
                ws = app.ActiveSheet as Excel.Worksheet;
                if (ws == null) return;
                outline = ws.Outline;
                outline.ShowLevels(RowLevels: level);
            }
            catch { }
            finally { ComUtil.Release(outline, ws); }
        }

        public static void ExportCurrentSheet()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            
            Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
            if (activeSheet == null) return;

            bool originalAlerts = app.DisplayAlerts;

            Excel.Workbook sourceWb = null;
            Excel.Workbook newWb = null;
            Excel.Sheets newSheets = null;
            Excel.Worksheet targetSheet = null;
            Excel.Range usedRange = null;

            using (ExcelScope.Begin(app))
            try
            {
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
                sourceWb = app.ActiveWorkbook;
                string savePath = sourceWb == null ? "" : sourceWb.Path;
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
                newWb = app.ActiveWorkbook;
                newSheets = newWb.Sheets;
                targetSheet = newSheets[1] as Excel.Worksheet;

                // Cancel auto-filter on target (defensive, in case copy preserved filter state)
                if (targetSheet.AutoFilterMode)
                {
                    targetSheet.AutoFilterMode = false;
                }

                // Paste Values & Formats
                usedRange = targetSheet.UsedRange;
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

                app.DisplayAlerts = originalAlerts;
                ComUtil.Release(usedRange, targetSheet, newSheets, newWb, sourceWb, activeSheet);
            }
        }

        
        public static void ConvertAllToValues()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            
            Excel.Workbook wb = app.ActiveWorkbook;
            if (wb == null) return;
            
            // 逐表整片 PasteSpecial 成数值，自动计算开着的话每贴一张表就重算一遍
            using (ExcelScope.Begin(app))
            try
            {
                foreach (string sheetName in ComUtil.GetSheetNames(wb))
                {
                    Excel.Worksheet ws = null;
                    Excel.Range usedRange = null;
                    try
                    {
                        ws = ComUtil.GetSheet(wb, sheetName);
                        if (ws == null) continue;

                        // Cancel auto-filter before copy to ensure all rows are included
                        if (ws.AutoFilterMode)
                        {
                            try { ws.AutoFilterMode = false; } catch { }
                        }

                        usedRange = ws.UsedRange;
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
                    finally { ComUtil.Release(usedRange, ws); }
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

                ComUtil.Release(wb);
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



            bool originalAlerts = app.DisplayAlerts;

            dynamic newWb = null;

            using (ExcelScope.Begin(app))

            try

            {

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

                dynamic sourceSheets = sourceWb.Sheets;
                try { sourceSheets.Copy(); }
                finally { ComUtil.Release((object)sourceSheets); }

                newWb = app.ActiveWorkbook;



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

                app.DisplayAlerts = originalAlerts;

                ComUtil.Release((object)newWb, (object)sourceWb);
            }

        }



        private static void ConvertAllToValuesInternal(dynamic wb)

        {

            dynamic app = WpsExcelAddIn.App;

            foreach (string sheetName in SheetNames(wb))
            {
                dynamic ws = GetSheet(wb, sheetName);
                if (ws == null) continue;

                dynamic usedRange = null;
                try
                {
                    if (ws.AutoFilterMode)
                    {
                        try { ws.AutoFilterMode = false; } catch { }
                    }

                    app.CutCopyMode = false;

                    usedRange = ws.UsedRange;
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
                finally { ComUtil.Release((object)usedRange, (object)ws); }
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

            dynamic wbNames = null;
            try
            {
                wbNames = wb.Names;
                foreach (dynamic name in wbNames)
                {
                    try
                    {
                        string val = "";
                        try { val = name.Value; } catch { continue; }

                        if (val.Contains("#REF!") || val.Contains(":\\") || val.Contains("\\\\") || val.Contains("#N/A"))
                        {
                            problemNames.Add((string)name.Name);
                        }
                    }
                    finally { ComUtil.Release((object)name); }
                }
            }
            catch { }
            finally { ComUtil.Release((object)wbNames); }



            // 外部工作簿引用模式: 匹配所有 [xxx] 格式（含 [141]Laldia清单 等无扩展名格式）

            var externalRefRegex = new System.Text.RegularExpressions.Regex(

                @"\[[^\]]*\]",

                System.Text.RegularExpressions.RegexOptions.Compiled |

                System.Text.RegularExpressions.RegexOptions.IgnoreCase);



            List<string> allSheetNames = SheetNames(wb);
            int totalSheets = allSheetNames.Count;

            int sheetIndex = 0;



            // ===== 阶段1：只粘死自身公式含外部链接的单元格 =====

            foreach (string phase1Name in allSheetNames)

            {

                sheetIndex++;

                app.StatusBar = string.Format("导出报表: 粘死外部引用 ({0}/{1}) {2}...",

                    sheetIndex, totalSheets, phase1Name);

                dynamic ws = GetSheet(wb, phase1Name);
                if (ws == null) continue;

                dynamic usedRange = null;

                try

                {

                    usedRange = ws.UsedRange;

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

                            if (problemNames.Count > 0)

                            {

                                foreach (System.Text.RegularExpressions.Match token in IdentifierRegex.Matches(scrubbed))

                                {

                                    if (problemNames.Contains(token.Value))

                                    {

                                        needsConversion = true;

                                        break;

                                    }

                                }

                            }



                            // 只检查外部工作簿引用（公式文本里带 [xxx.xlsx]）

                            // 注意用原始 formula：路径式外部引用 'C:\x\[a.xlsx]Sheet'!A1 的方括号在引号内

                            if (!needsConversion && externalRefRegex.IsMatch(formula))

                            {

                                needsConversion = true;

                            }



                            if (needsConversion)

                            {

                                FlattenCell(ws, startRow + r - 1, startCol + c - 1, isArray);

                            }

                        }

                    }

                }

                catch { }

                finally { ComUtil.Release((object)usedRange, (object)ws); }

            }



            // ===== 阶段2：粘死引用了隐藏列的公式（隐藏列将在步骤7被删除） =====

            sheetIndex = 0;

            foreach (string phase2Name in allSheetNames)

            {

                sheetIndex++;

                app.StatusBar = string.Format("导出报表: 粘死隐藏列引用 ({0}/{1}) {2}...",

                    sheetIndex, totalSheets, phase2Name);

                dynamic ws = GetSheet(wb, phase2Name);
                if (ws == null) continue;

                dynamic usedRange = null;

                try

                {

                    usedRange = ws.UsedRange;

                    if (usedRange == null) continue;



                    // 获取当前工作表被隐藏的列字母集合

                    var hiddenColLetters = new HashSet<string>();

                    int firstCol = usedRange.Column;

                    dynamic usedCols = usedRange.Columns;
                    int lastCol;
                    try { lastCol = firstCol + (int)usedCols.Count - 1; }
                    finally { ComUtil.Release((object)usedCols); }

                    for (int c = firstCol; c <= lastCol; c++)

                    {

                        dynamic column = null;

                        try

                        {

                            column = ws.Columns[c];

                            if (column.Hidden)

                                hiddenColLetters.Add(ColumnNumberToLetter(c));

                        }

                        catch { }

                        finally { ComUtil.Release((object)column); }

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

                                FlattenCell(ws, startRow + r - 1, startCol + c - 1, isArray);

                            }

                        }

                    }

                }

                catch { }

                finally { ComUtil.Release((object)usedRange, (object)ws); }

            }

        }

        /// <summary>
        /// 把一格公式粘成值。数组公式要整块处理，写单格会报「不能更改数组的某一部分」。
        /// </summary>
        private static void FlattenCell(dynamic ws, int rowIdx, int colIdx, bool isArray)
        {
            dynamic cell = null;
            dynamic arr = null;
            try
            {
                cell = ws.Cells[rowIdx, colIdx];
                if (isArray)
                {
                    arr = cell.CurrentArray;
                    arr.Value = arr.Value;
                }
                else
                {
                    cell.Value = cell.Value;
                }
            }
            catch { }
            finally { ComUtil.Release((object)arr, (object)cell); }
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



        // 从公式里切出标识符再查表。名称合法字符为字母数字下划线点反斜杠（\w 已覆盖中日韩）。

        // 不能反过来拿几千个名称合并成交替正则去匹配公式，那是 O(公式长度 x 名称数)，

        // 工作簿有上万个定义名称时会直接卡死。

        private static readonly System.Text.RegularExpressions.Regex IdentifierRegex =

            new System.Text.RegularExpressions.Regex(

                @"[\w.\\]+",

                System.Text.RegularExpressions.RegexOptions.Compiled);



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
            foreach (string sheetName in SheetNames(wb))
            {
                dynamic ws = GetSheet(wb, sheetName);
                if (ws == null) continue;
                try
                {
                    if (ws.AutoFilterMode) ws.AutoFilterMode = false;
                }
                catch { }
                finally { ComUtil.Release((object)ws); }
            }
        }

        private static void ClearErrorNames(dynamic wb)
        {
            dynamic wbNames = null;
            try
            {
                List<string> namesToDelete = new List<string>();

                wbNames = wb.Names;
                foreach (dynamic name in wbNames)
                {
                    try
                    {
                        string val = "";
                        try { val = name.Value; } catch { continue; }

                        if (val.Contains("#REF!") || val.Contains(":\\") || val.Contains("\\\\") || val.Contains("#N/A"))
                        {
                            namesToDelete.Add((string)name.Name);
                        }
                    }
                    finally { ComUtil.Release((object)name); }
                }

                foreach (string n in namesToDelete)
                {
                    dynamic item = null;
                    try
                    {
                        item = wbNames.Item(n);
                        item.Delete();
                    }
                    catch { }
                    finally { ComUtil.Release((object)item); }
                }
            }
            catch { }
            finally { ComUtil.Release((object)wbNames); }
        }

        private static void ShowOutlineLevels(dynamic wb, bool rows)
        {
            foreach (string sheetName in SheetNames(wb))
            {
                dynamic ws = GetSheet(wb, sheetName);
                if (ws == null) continue;

                dynamic outline = null;
                try
                {
                    outline = ws.Outline;
                    if (rows) outline.ShowLevels(RowLevels: 4);
                    else outline.ShowLevels(ColumnLevels: 1);
                }
                catch { }
                finally { ComUtil.Release((object)outline, (object)ws); }
            }
        }

        private static void HideNonLevel1Columns(dynamic wb)
        {
            ShowOutlineLevels(wb, false);
        }

        private static void ExpandAllRows(dynamic wb)
        {
            ShowOutlineLevels(wb, true);
        }

        private static void DeleteHiddenColumns(dynamic wb)
        {
            foreach (string sheetName in SheetNames(wb))
            {
                dynamic ws = GetSheet(wb, sheetName);
                if (ws == null) continue;

                dynamic usedRange = null;
                dynamic usedCols = null;
                try
                {
                    usedRange = ws.UsedRange;
                    usedCols = usedRange.Columns;
                    int colCount = (int)usedCols.Count;
                    // Loop backwards is safer for deletion, but VBA used a Do While loop with Counter adjustment

                    for (int i = colCount; i >= 1; i--)
                    {
                        dynamic col = null;
                        dynamic entire = null;
                        try
                        {
                            col = usedCols[i];
                            entire = col.EntireColumn;
                            if (entire.Hidden) entire.Delete();
                        }
                        catch { }
                        finally { ComUtil.Release((object)entire, (object)col); }
                    }
                }
                catch { }
                finally { ComUtil.Release((object)usedCols, (object)usedRange, (object)ws); }
            }
        }



        private static void CleanOutsidePrintArea(dynamic wb)
        {
            dynamic app = WpsExcelAddIn.App;

            foreach (string sheetName in SheetNames(wb))
            {
                dynamic ws = GetSheet(wb, sheetName);
                if (ws == null) continue;

                dynamic printRange = null;
                dynamic usedRange = null;
                dynamic deleteCols = null;
                dynamic deleteRows = null;
                try
                {
                    // 1. Find Print_Area
                    dynamic wsNames = null;
                    try
                    {
                        wsNames = ws.Names;
                        foreach (dynamic name in wsNames)
                        {
                            bool matched = false;
                            try
                            {
                                if (((string)name.Name).EndsWith("!Print_Area"))
                                {
                                    printRange = name.RefersToRange;
                                    matched = true;
                                }
                            }
                            catch { }
                            finally { ComUtil.Release((object)name); }
                            if (matched) break;
                        }
                    }
                    catch { }
                    finally { ComUtil.Release((object)wsNames); }

                    // Fallback to UsedRange if no print area (creating a temporary print area name logic from VBA seems redundant if we just use UsedRange, but let's follow logic: "Use UsedRange as PrintArea if none")
                    if (printRange == null)
                    {
                        printRange = ws.UsedRange;
                    }

                    usedRange = ws.UsedRange;

                    // 2.1 Delete Columns outside
                    dynamic cols = null;
                    try
                    {
                        cols = usedRange.Columns;
                        foreach (dynamic col in cols)
                        {
                            bool keepCol = false;
                            dynamic intersect = null;
                            try
                            {
                                intersect = app.Intersect(col, printRange);
                                if (intersect == null)
                                {
                                    if (deleteCols == null) { deleteCols = col; keepCol = true; }
                                    else
                                    {
                                        dynamic merged = app.Union(deleteCols, col);
                                        ComUtil.Release((object)deleteCols);
                                        deleteCols = merged;
                                    }
                                }
                            }
                            catch { }
                            finally
                            {
                                ComUtil.Release((object)intersect);
                                if (!keepCol) ComUtil.Release((object)col);
                            }
                        }
                    }
                    finally { ComUtil.Release((object)cols); }

                    if (deleteCols != null)
                    {
                        dynamic entire = null;
                        try { entire = deleteCols.EntireColumn; entire.Delete(); }
                        catch { }
                        finally { ComUtil.Release((object)entire); }
                    }

                    // 2.2 Delete Rows outside
                    dynamic rows = null;
                    try
                    {
                        rows = usedRange.Rows;
                        foreach (dynamic row in rows)
                        {
                            bool keepRow = false;
                            dynamic intersect = null;
                            try
                            {
                                intersect = app.Intersect(row, printRange);
                                if (intersect == null)
                                {
                                    if (deleteRows == null) { deleteRows = row; keepRow = true; }
                                    else
                                    {
                                        dynamic merged = app.Union(deleteRows, row);
                                        ComUtil.Release((object)deleteRows);
                                        deleteRows = merged;
                                    }
                                }
                            }
                            catch { }
                            finally
                            {
                                ComUtil.Release((object)intersect);
                                if (!keepRow) ComUtil.Release((object)row);
                            }
                        }
                    }
                    finally { ComUtil.Release((object)rows); }

                    if (deleteRows != null)
                    {
                        dynamic entire = null;
                        try { entire = deleteRows.EntireRow; entire.Delete(); }
                        catch { }
                        finally { ComUtil.Release((object)entire); }
                    }
                }
                catch { }
                finally
                {
                    ComUtil.Release((object)deleteRows, (object)deleteCols,
                                    (object)usedRange, (object)printRange, (object)ws);
                }
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

