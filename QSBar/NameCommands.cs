using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace QSBar
{
    public static class NameCommands
    {
        public class NameEntry
        {
            public Excel.Name ComName;
            public string FullName;
            public string ShortName;
            public string RefersTo;
            public bool IsExternal;
            public bool IsBroken;
            public bool IsHidden;
        }

        public class PurgeResult
        {
            public List<NameEntry> Broken = new List<NameEntry>();
            public List<NameEntry> Unused = new List<NameEntry>();
            public List<NameEntry> UnusedExternal = new List<NameEntry>();
            public List<NameEntry> UsedExternal = new List<NameEntry>();
            public List<NameEntry> Kept = new List<NameEntry>();
            public int BuiltInSkipped;
        }

        // RefersTo 本身是错误值的名称一律删除，不看在不在用：引用它们的公式
        // 本来就已经在返回 #REF! / #NAME?，留着没有意义
        private static readonly string[] ErrorMarkers =
        {
            "#REF!", "#N/A", "#NAME?", "#VALUE!", "#DIV/0!", "#NULL!", "#NUM!"
        };

        // 工作表名与文本常量不是名称引用，匹配前必须剥离，否则表名里的字母会撞上短名称
        private static readonly Regex QuotedLiteralRegex =
            new Regex(@"'(?:[^']|'')*'|""(?:[^""]|"""")*""", RegexOptions.Compiled);

        // 从公式里切出标识符再查表，而不是拿几千个名称合并成交替正则去匹配公式。
        // 后者是 O(文本长度 x 名称数)，工作簿有上万个名称时会直接卡死。
        private static readonly Regex IdentifierRegex =
            new Regex(@"[\w.\\]+", RegexOptions.Compiled);

        private static readonly HashSet<string> BuiltInShortNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Print_Area", "Print_Titles", "_FilterDatabase", "FilterDatabase",
                "Criteria", "Extract", "Database", "Consolidate_Area", "Sheet_Title"
            };

        public static void PurgeNames()
        {
            Excel.Application app = WpsExcelAddIn.App;
            if (app == null) return;
            Excel.Workbook wb = app.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show(Cn() ? "没有打开的工作簿。" : "No workbook is open.",
                    Title(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool oldUpdating = app.ScreenUpdating;
            Excel.XlCalculation oldCalc = app.Calculation;
            bool oldEvents = app.EnableEvents;

            PurgeResult result;
            var sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                app.StatusBar = Cn() ? "净化名称: 收集定义名称..." : "Purge Names: collecting names...";
                result = Analyze(app, wb);
                sw.Stop();
            }
            catch (Exception ex)
            {
                app.StatusBar = false;
                MessageBox.Show((Cn() ? "分析名称时出错: " : "Error while analyzing names: ") + ex.Message,
                    Title(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                app.StatusBar = false;
                app.EnableEvents = oldEvents;
                app.Calculation = oldCalc;
                app.ScreenUpdating = oldUpdating;
            }

            int deletable = result.Broken.Count + result.Unused.Count + result.UnusedExternal.Count;
            if (deletable == 0 && result.UsedExternal.Count == 0)
            {
                MessageBox.Show(
                    Cn() ? "没有可清理的名称，当前工作簿的自定义名称都在使用中。"
                         : "Nothing to purge. All custom names in this workbook are in use.",
                    Title(), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            List<NameEntry> confirmed;
            using (var form = new PurgeNamesForm(result))
            {
                if (form.ShowDialog() != DialogResult.OK) return;
                confirmed = form.NamesToDelete;
            }

            if (confirmed.Count == 0) return;

            int deleted = 0;
            int failed = 0;
            var failedNames = new List<string>();
            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;

                for (int i = 0; i < confirmed.Count; i++)
                {
                    if (i % 50 == 0)
                    {
                        app.StatusBar = string.Format(
                            Cn() ? "净化名称: 删除 {0}/{1} ..." : "Purge Names: deleting {0}/{1} ...",
                            i + 1, confirmed.Count);
                        System.Windows.Forms.Application.DoEvents();
                    }

                    if (TryDeleteName(wb, confirmed[i])) deleted++;
                    else
                    {
                        failed++;
                        if (failedNames.Count < 8) failedNames.Add(confirmed[i].FullName);
                    }
                }
            }
            finally
            {
                app.StatusBar = false;
                app.EnableEvents = oldEvents;
                app.ScreenUpdating = oldUpdating;
            }

            string message = string.Format(
                Cn() ? "已删除 {0} 个名称。\n保留在用名称 {1} 个，内置名称 {2} 个。\n删除失败 {3} 个。\n扫描耗时 {4} 秒。"
                     : "Deleted {0} names.\nKept {1} names in use and {2} built-in names.\nFailed: {3}.\nScan took {4}s.",
                deleted, result.Kept.Count, result.BuiltInSkipped, failed,
                sw.Elapsed.TotalSeconds.ToString("0.00"));

            if (failedNames.Count > 0)
            {
                message += (Cn() ? "\n\n删除失败的名称：\n" : "\n\nFailed names:\n")
                    + string.Join("\n", failedNames.ToArray());
            }

            MessageBox.Show(message, Title(), MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Print_Area / Print_Titles 由工作表的 PageSetup 拥有。只 Delete 名称的话，
        /// PageSetup 里仍指着原区域，Excel 会立刻把名称重建回来，看上去就是「删不掉」。
        /// </summary>
        internal static bool TryDeleteName(Excel.Workbook wb, NameEntry entry)
        {
            bool isPrintArea = string.Equals(entry.ShortName, "Print_Area", StringComparison.OrdinalIgnoreCase);
            bool isPrintTitles = string.Equals(entry.ShortName, "Print_Titles", StringComparison.OrdinalIgnoreCase);

            if (isPrintArea || isPrintTitles)
            {
                Excel.Worksheet owner = null;
                try { owner = entry.ComName.Parent as Excel.Worksheet; }
                catch { }

                if (owner != null)
                {
                    ClearPrintSetup(owner, isPrintArea);
                }
                else
                {
                    // 工作簿级的 Print_Titles 拿不到所属工作表，只能逐表清
                    foreach (Excel.Worksheet ws in wb.Worksheets) ClearPrintSetup(ws, isPrintArea);
                }
            }

            try { entry.ComName.Delete(); }
            catch { }
            if (!StillAlive(entry)) return true;

            // RefersTo 已经是 #REF! 时 Delete 会静默失效，先把它指向一个合法区域再删
            try
            {
                Excel.Worksheet first = wb.Worksheets[1] as Excel.Worksheet;
                if (first != null)
                {
                    entry.ComName.RefersTo = "='" + first.Name.Replace("'", "''") + "'!$A$1";
                    entry.ComName.Delete();
                }
            }
            catch { }

            return !StillAlive(entry);
        }

        private static void ClearPrintSetup(Excel.Worksheet ws, bool printArea)
        {
            try
            {
                if (printArea)
                {
                    ws.PageSetup.PrintArea = "";
                }
                else
                {
                    ws.PageSetup.PrintTitleRows = "";
                    ws.PageSetup.PrintTitleColumns = "";
                }
            }
            catch { }
        }

        // 已删除的 Name 对象再访问属性会抛 COM 异常，据此判断是否真的删掉了
        private static bool StillAlive(NameEntry entry)
        {
            try { return !string.IsNullOrEmpty(entry.ComName.Name); }
            catch { return false; }
        }

        private static PurgeResult Analyze(Excel.Application app, Excel.Workbook wb)
        {
            var result = new PurgeResult();
            var candidates = new List<NameEntry>();

            foreach (Excel.Name nm in wb.Names)
            {
                string full;
                try { full = nm.Name; }
                catch { continue; }

                // RefersTo 在名称已损坏时可能直接抛 COM 异常。不能因此跳过整个名称，
                // 否则最该删的那些反而会留下来。读不出来就当作坏名称处理
                string refersTo;
                try { refersTo = nm.RefersTo as string; }
                catch { refersTo = null; }

                string shortName = StripSheetPrefix(full);
                if (IsNeverDelete(shortName) || IsNeverDelete(full))
                {
                    result.BuiltInSkipped++;
                    continue;
                }

                var entry = new NameEntry
                {
                    ComName = nm,
                    FullName = full,
                    ShortName = shortName,
                    RefersTo = refersTo ?? "",
                    IsExternal = IsExternalRefersTo(refersTo),
                    IsBroken = IsBrokenRefersTo(refersTo),
                    IsHidden = IsHiddenName(nm)
                };

                if (IsSheetBuiltIn(shortName) || IsSheetBuiltIn(full))
                {
                    // 内置名称不参与「在用」判定，但坏掉或指向外部文件时同样要清掉：
                    // Print_Titles = #REF! 的打印标题早已失效，留着只会继续污染
                    if (entry.IsBroken || entry.IsExternal) result.Broken.Add(entry);
                    else result.BuiltInSkipped++;
                    continue;
                }

                candidates.Add(entry);
            }

            if (candidates.Count == 0) return result;

            var lookup = new Dictionary<string, List<NameEntry>>(StringComparer.OrdinalIgnoreCase);
            foreach (NameEntry e in candidates)
            {
                List<NameEntry> bucket;
                if (!lookup.TryGetValue(e.ShortName, out bucket))
                {
                    bucket = new List<NameEntry>();
                    lookup[e.ShortName] = bucket;
                }
                bucket.Add(e);
            }

            var live = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectCellSideReferences(app, wb, lookup, live);

            // 名称链要传播到收敛：删掉一批后，只被这批引用的名称也变成垃圾。
            // 必须从「存活集」出发正向传播 —— 被垃圾名称引用不算在用。
            // 用队列只处理新入集的名称，避免反复重扫全部候选。
            var pending = new Queue<string>(live);
            while (pending.Count > 0)
            {
                string current = pending.Dequeue();
                List<NameEntry> owners;
                if (!lookup.TryGetValue(current, out owners)) continue;

                foreach (NameEntry owner in owners)
                {
                    // 坏名称自己就要被删，不能让它把引用到的名称也拖成「在用」。
                    // 那串 'file:///...[人工表(直表)1]MAT0195'!#REF! 里就带着可被误认的标识符
                    if (owner.IsBroken) continue;

                    foreach (string hit in MatchNames(owner.RefersTo, lookup))
                    {
                        if (live.Add(hit)) pending.Enqueue(hit);
                    }
                }
            }

            foreach (NameEntry e in candidates)
            {
                if (e.IsBroken)
                {
                    result.Broken.Add(e);
                    continue;
                }

                bool inUse = live.Contains(e.ShortName);
                if (inUse && e.IsExternal) result.UsedExternal.Add(e);
                else if (inUse) result.Kept.Add(e);
                else if (e.IsExternal) result.UnusedExternal.Add(e);
                else result.Unused.Add(e);
            }

            return result;
        }

        private static void CollectCellSideReferences(Excel.Application app, Excel.Workbook wb,
            Dictionary<string, List<NameEntry>> lookup, HashSet<string> live)
        {
            int total = wb.Worksheets.Count;
            int index = 0;

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                index++;
                app.StatusBar = string.Format(
                    Cn() ? "净化名称: 扫描引用 ({0}/{1}) {2}..." : "Purge Names: scanning ({0}/{1}) {2}...",
                    index, total, ws.Name);
                System.Windows.Forms.Application.DoEvents();

                Excel.Range used = null;
                try { used = ws.UsedRange; }
                catch { }
                if (used == null) continue;

                try
                {
                    object formula = used.Formula;
                    object[,] grid = formula as object[,];
                    if (grid != null)
                    {
                        int rLo = grid.GetLowerBound(0), rHi = grid.GetUpperBound(0);
                        int cLo = grid.GetLowerBound(1), cHi = grid.GetUpperBound(1);
                        for (int r = rLo; r <= rHi; r++)
                        {
                            // 合并正则有几百个交替分支，对普通文本逐格跑会慢到不可用。
                            // 名称只可能出现在公式里，非公式单元格必须先滤掉。
                            for (int c = cLo; c <= cHi; c++)
                            {
                                string cellText = grid[r, c] as string;
                                if (cellText == null || cellText.Length == 0) continue;
                                if (cellText[0] != '=' && !cellText.StartsWith("{=")) continue;
                                AddHits(cellText, lookup, live);
                            }

                            if ((r - rLo) % 200 == 0) System.Windows.Forms.Application.DoEvents();
                        }
                    }
                    else
                    {
                        string single = formula as string;
                        if (single != null && single.Length > 0 &&
                            (single[0] == '=' || single.StartsWith("{=")))
                        {
                            AddHits(single, lookup, live);
                        }
                    }
                }
                catch { }

                CollectValidationReferences(used, lookup, live);
                CollectFormatConditionReferences(used, lookup, live);
            }
        }

        private static void CollectValidationReferences(Excel.Range used,
            Dictionary<string, List<NameEntry>> lookup, HashSet<string> live)
        {
            Excel.Range validated = null;
            // 必须限定在 UsedRange 内。对 ws.Cells 调用等于扫整表 104 万行
            try { validated = used.SpecialCells(Excel.XlCellType.xlCellTypeAllValidation); }
            catch { return; }
            if (validated == null) return;

            foreach (Excel.Range area in validated.Areas)
            {
                // 整个 Area 验证规则一致时可一次取到；不一致会抛异常，退回取首格
                try { AddHits(area.Validation.Formula1, lookup, live); }
                catch
                {
                    try { AddHits(((Excel.Range)area.Cells[1, 1]).Validation.Formula1, lookup, live); }
                    catch { }
                }
                try { AddHits(area.Validation.Formula2, lookup, live); }
                catch { }
            }
        }

        private static void CollectFormatConditionReferences(Excel.Range used,
            Dictionary<string, List<NameEntry>> lookup, HashSet<string> live)
        {
            try
            {
                Excel.FormatConditions conditions = used.FormatConditions as Excel.FormatConditions;
                if (conditions == null) return;

                for (int i = 1; i <= conditions.Count; i++)
                {
                    // FormatCondition 有多种类型，部分没有 Formula1/Formula2
                    dynamic fc = conditions[i];
                    try { AddHits(fc.Formula1 as string, lookup, live); }
                    catch { }
                    try { AddHits(fc.Formula2 as string, lookup, live); }
                    catch { }
                }
            }
            catch { }
        }

        private static void AddHits(string text,
            Dictionary<string, List<NameEntry>> lookup, HashSet<string> live)
        {
            foreach (string hit in MatchNames(text, lookup))
            {
                live.Add(hit);
            }
        }

        private static IEnumerable<string> MatchNames(string text,
            Dictionary<string, List<NameEntry>> lookup)
        {
            foreach (string ident in Identifiers(text))
            {
                List<NameEntry> bucket;
                if (lookup.TryGetValue(ident, out bucket) && bucket.Count > 0)
                {
                    yield return bucket[0].ShortName;
                }
            }
        }

        /// <summary>
        /// 从公式里切出可能是名称引用的标识符。工作表名与文本常量已剥离。
        /// </summary>
        internal static IEnumerable<string> Identifiers(string text)
        {
            if (string.IsNullOrEmpty(text)) yield break;

            // 绝大多数公式里没有引号，先看一眼能省掉一次全文正则替换
            string scrubbed = (text.IndexOf('\'') >= 0 || text.IndexOf('"') >= 0)
                ? QuotedLiteralRegex.Replace(text, "''")
                : text;

            foreach (Match m in IdentifierRegex.Matches(scrubbed))
            {
                // 紧跟左括号的是函数调用而非名称引用。工作簿里存在名为 IF 的
                // 垃圾名称时，不排除会被每个 =IF(...) 判成在用而永远删不掉
                int end = m.Index + m.Length;
                if (end < scrubbed.Length && scrubbed[end] == '(') continue;

                yield return m.Value;
            }
        }

        internal static string StripSheetPrefix(string fullName)
        {
            if (string.IsNullOrEmpty(fullName)) return fullName;
            int bang = fullName.LastIndexOf('!');
            return bang >= 0 ? fullName.Substring(bang + 1) : fullName;
        }

        // _xlfn 是当前版本不支持的新函数占位、_xlpm 是 LAMBDA 参数，
        // 它们的 RefersTo 正常情况下就是 #NAME?，删掉会直接破坏公式
        internal static bool IsNeverDelete(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            return name.StartsWith("_xlfn", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("_xlpm", StringComparison.OrdinalIgnoreCase)
                // _xlcn 由数据连接拥有，删不掉，列进去只会每次都报失败
                || name.StartsWith("_xlcn", StringComparison.OrdinalIgnoreCase);
        }

        // 打印区域、打印标题、高级筛选条件区等。完好时保留，坏掉或指向外部文件时照删
        private static bool IsSheetBuiltIn(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            if (name.StartsWith("_xlnm", StringComparison.OrdinalIgnoreCase)) return true;
            return BuiltInShortNames.Contains(name);
        }

        private static bool IsExternalRefersTo(string refersTo)
        {
            if (string.IsNullOrEmpty(refersTo)) return false;
            return refersTo.IndexOf('[') >= 0
                || refersTo.IndexOf(":\\", StringComparison.Ordinal) >= 0
                || refersTo.IndexOf("\\\\", StringComparison.Ordinal) >= 0
                || refersTo.IndexOf("://", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // Visible=false 的名称不会出现在 Excel 的名称管理器里，用户找不到它们，
        // 清单里必须标出来，否则看着像凭空冒出来的
        private static bool IsHiddenName(Excel.Name nm)
        {
            try { return !nm.Visible; }
            catch { return false; }
        }

        private static bool IsBrokenRefersTo(string refersTo)
        {
            if (string.IsNullOrEmpty(refersTo)) return true;

            for (int i = 0; i < ErrorMarkers.Length; i++)
            {
                if (refersTo.IndexOf(ErrorMarkers[i], StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        private static bool Cn()
        {
            return WpsExcelAddIn.UseChineseRibbon;
        }

        private static string Title()
        {
            return Cn() ? "净化名称" : "Purge Names";
        }
    }
}
