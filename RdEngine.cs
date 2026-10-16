using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace ReviewMerge {
    // Columns of the RD check sheet («Проверка комплектов») found by their headers: a reviewer's copy may have them in other places.
    public sealed class RdLayout {
        public int Header, First;
        public readonly Dictionary<string, int> Col = new Dictionary<string, int>(StringComparer.Ordinal);
        // П1…П6, Э1…Э6, ПР in the order of the sheet.
        public readonly List<string> Flags = new List<string>();
        public int this[string key] { get { int c; return Col.TryGetValue(key, out c) ? c : 0; } }
        public string L(string key) { return MergeEngine.ExcelColumn(this[key]); }
        // Columns the program writes formulas into, and the columns of the reviewer.
        public static readonly string[] AutoKeys = { "ve", "te", "tedate", "teletter", "cnt", "status", "latest", "order", "elkey" };
        public static readonly string[] Identity = { "el", "elname", "pd", "rd", "name", "rev", "section" };
        public static readonly string[] Review = { "who", "date", "text", "note", "smr", "fixed" };
        static readonly string[][] Heads = {
            new[] { "el", "Номер элемента" }, new[] { "elname", "Наименование элемента" }, new[] { "pd", "Шифр ПД" }, new[] { "rd", "Шифр РД" },
            new[] { "name", "Наименование комплекта" }, new[] { "name", "Наименование документации" }, new[] { "rev", "Изм." }, new[] { "rev", "Изм" },
            new[] { "section", "Раздел" }, new[] { "ve", "Сверка с ВЭ РД" }, new[] { "te", "Согласование ТЭ" }, new[] { "tedate", "Дата ТЭ" },
            new[] { "teletter", "№ письма ТЭ" }, new[] { "who", "Проверил" }, new[] { "date", "Дата проверки" }, new[] { "text", "Замечания" },
            new[] { "note", "Примечание" }, new[] { "smr", "Факт СМР" }, new[] { "cnt", "Видов замечаний" }, new[] { "status", "Статус комплекта" },
            new[] { "latest", "Действующая редакция" }, new[] { "order", "Порядковый № элемента" }, new[] { "elkey", "Элемент строки" } };
        public static readonly Dictionary<string, string> Names = Heads.GroupBy(h => h[0]).ToDictionary(g => g.Key, g => g.First()[1]);
        static string KeyOf(object header) {
            string h = XlsxReader.Normal(header);
            if (h == "") return null;
            var flag = Regex.Match(h, @"^(П[1-6]|Э[1-6]|ПР)(?:\.|\s|$)");
            if (flag.Success) return "flag:" + flag.Groups[1].Value;
            if (Rules.IsResolvedHeader(header)) return "fixed";
            if (h.StartsWith(XlsxReader.Normal("Замечания (текст"), StringComparison.Ordinal)) return "text";
            foreach (var head in Heads) if (h == XlsxReader.Normal(head[1])) return head[0];
            return null;
        }
        // The header is the first of the top 30 rows with «Шифр РД» and «Проверил»; the row of column numbers under it is skipped.
        public static RdLayout Detect(RawSheet sheet) {
            foreach (var row in sheet.Rows.Where(r => r.Key <= 30)) {
                var keys = row.Value.Select(KeyOf).ToList();
                if (!keys.Contains("rd") || !keys.Contains("who")) continue;
                var layout = new RdLayout { Header = row.Key, First = row.Key + 1 };
                for (int c = 0; c < keys.Count; c++) {
                    string k = keys[c]; if (k == null) continue;
                    if (k.StartsWith("flag:", StringComparison.Ordinal)) { k = k.Substring(5); if (!layout.Col.ContainsKey(k)) layout.Flags.Add(k); }
                    if (!layout.Col.ContainsKey(k)) layout.Col[k] = c + 1;
                }
                object[] next;
                if (sheet.Rows.TryGetValue(layout.First, out next) && XlsxReader.Text(next[layout["rd"] - 1]) == layout["rd"].ToString(CultureInfo.InvariantCulture)) layout.First++;
                return layout;
            }
            return null;
        }
    }
    // A row of a check sheet with its element carried down from the rows above (as «Элемент строки»).
    public sealed class RdRow {
        public string File, ElKey, ElName;
        public int Row;
        public object[] Values;
        public RdLayout Layout;
        public object this[string key] { get { int c = Layout[key]; return c > 0 && c <= Values.Length ? Values[c - 1] : null; } }
    }
    // A set (Шифр РД and Изм.) of the consolidated book after the merge; Values holds the reviewer's columns and the flags by code.
    public sealed class RdSet {
        public string Key;
        public int SheetRow;
        public bool IsNew, WriteElement;
        public RdRow TargetRow;
        public readonly List<RdRow> Reviews = new List<RdRow>();
        public readonly Dictionary<string, object> Values = new Dictionary<string, object>(StringComparer.Ordinal);
        public object this[string key] { get { object v; return Values.TryGetValue(key, out v) ? v : null; } set { Values[key] = value; } }
        // Formula results of the main sheet, calculated the way Excel does for the cached values and the PDF.
        public string ElKey, Ve, Status, TeLetter;
        public object Te, TeDate;
        public int Cnt, Order;
        public bool Latest;
        public double? Date;
        public string Who { get { return Rules.CleanName(this["who"]); } }
        public bool Expert { get { return RdBook.IsExpert(this["who"]); } }
    }
    // An element of the «Элементы» sheet, in the order the check sheet first names it.
    public sealed class RdElement {
        public string Key, Status, Act, Who;
        public object Name, Dgp, ActVe, Vpr;
        public int InVe, Presented, Missing, NotInVe, Checked, Remarks, NotAgreed, Ready, Expert;
        public double? Date;
        public readonly List<RdSet> Live = new List<RdSet>();
    }
    // The stage list on «Ведомость» (the pasted «База»): the first row of every trimmed code and the element of every row carried down.
    public sealed class RdVed {
        public RawSheet Sheet;
        // Data rows, and the last row of the helper formulas AF:AH the build writes.
        public int First, Last, HelperLast;
        public readonly Dictionary<string, int> CodeRow = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<int, string> Element = new Dictionary<int, string>();
        public readonly Dictionary<string, int> ElementRow = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, int> ElementSets = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public int Sets, Elements;
        public object At(int row, int col) { return Sheet.At(row, col); }
    }
    // «Проверка актов»: the result of the first act of every element and the checks by reviewer and date.
    public sealed class RdActs {
        public string Name = "Проверка актов";
        public bool Exists;
        public int First = 7, Last = 100, El, Who, Date, Status, Conform, Remarks;
        public readonly Dictionary<string, string> ElementStatus = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public readonly List<Tuple<string, double>> Checks = new List<Tuple<string, double>>();
    }
    // Everything the RD build writes: the sets in sheet order, the stage list, the acts, the elements and the calendar.
    public sealed class RdBook {
        public const string Expert = "ТЭ";
        public static readonly string[] Agreed = { "Да", "Согласовано", "Согласована" };
        public string MainName;
        public RdLayout Layout;
        public RawSheet Target;
        // The last row of the main sheet that gets formulas, and the last row of the formula ranges before the build.
        public int Last, OldLast, ElementsLast;
        public readonly List<RdSet> Sets = new List<RdSet>();
        public RdVed Ved;
        public RdActs Acts;
        public readonly List<RdElement> Elements = new List<RdElement>();
        public readonly List<string> People = new List<string>();
        public readonly List<double> Days = new List<double>();
        // «Элемент строки» and «Порядковый № элемента» of every row of the main sheet.
        public readonly Dictionary<int, string> RowElKey = new Dictionary<int, string>();
        public readonly Dictionary<int, int> RowOrder = new Dictionary<int, int>();
        public readonly Dictionary<string, SortedSet<double>> Activity = new Dictionary<string, SortedSet<double>>(StringComparer.OrdinalIgnoreCase);
        public static bool IsExpert(object who) { return XlsxReader.Normal(who) == Expert; }
        public static bool Same(object a, object b) { return string.Equals(XlsxReader.Text(a), XlsxReader.Text(b), StringComparison.OrdinalIgnoreCase); }
    }

    public static class RdEngine {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const int Width = 80;
        static string Txt(object v) { return XlsxReader.Text(v); }
        static void Check(CancellationToken token) { token.ThrowIfCancellationRequested(); }
        static void AddProblem(MergeResult r, RdRow row, string detail) { r.Problems.Add(new Problem { File = row.File, Row = row.Row, Document = Txt(row["rd"]), Detail = detail }); }
        static IEnumerable<string> ReviewKeys(RdLayout layout) { return RdLayout.Review.Concat(layout.Flags); }
        static bool HasReview(RdRow r) { return ReviewKeys(r.Layout).Any(k => Txt(r[k]) != ""); }
        static object Flag(object value) {
            string s = Txt(value); if (s == "") return null;
            double n; if (double.TryParse(s.Replace(',', '.'), NumberStyles.Float, Inv, out n) && (n == 0 || n == 1)) return n;
            return s;
        }
        // Изм.: empty is the first issue (0); a number written as text becomes a number, as «Действующая редакция» compares numbers.
        static object Revision(object value) {
            string s = Txt(value); if (s == "") return null;
            double n; if (double.TryParse(s.Replace(',', '.'), NumberStyles.Float, Inv, out n)) return n;
            return s;
        }
        static string Key(RdRow row) {
            object rev = Revision(row["rev"]);
            return XlsxReader.Normal(row["rd"]) + "\u001f" + (rev == null ? "0" : rev is double ? ((double)rev).ToString(Inv) : XlsxReader.Normal(rev));
        }
        static RdRow Latest(IEnumerable<RdRow> rows) {
            return rows.Select((r, i) => new { Row = r, Order = i, Date = XlsxReader.DateSerial(r["date"]) ?? -1 }).OrderBy(x => x.Date).ThenBy(x => x.Order).Select(x => x.Row).LastOrDefault();
        }
        static string Pretty(string key, object value) {
            double? d = key == "date" || key == "fixed" ? XlsxReader.DateSerial(value) : null;
            return d.HasValue ? DateTime.FromOADate(d.Value).ToString("dd.MM.yyyy", Inv) : Txt(value) == "" ? "(пусто)" : Txt(value);
        }
        static RawSheet Sheet(string path, string name, bool required) {
            var sheet = XlsxReader.ReadSheet(path, name, Width);
            if (sheet == null && required) throw new InvalidDataException(Path.GetFileName(path) + ": нет листа «" + name + "».");
            return sheet;
        }
        static RdLayout Layout(RawSheet sheet, string name, bool target) {
            var layout = RdLayout.Detect(sheet);
            string file = Path.GetFileName(sheet.Path);
            if (layout == null) throw new InvalidDataException(file + ": на листе «" + name + "» нет шапки формы проверки РД — столбцов «Шифр РД» и «Проверил».");
            var required = target ? RdLayout.Identity.Where(k => k != "pd" && k != "name" && k != "section").Concat(RdLayout.Review.Where(k => k != "smr")).Concat(RdLayout.AutoKeys) : new[] { "rd", "who", "date" };
            var missing = required.Where(k => layout[k] == 0).Select(k => "«" + (k == "fixed" ? Rules.ResolvedHeader : RdLayout.Names[k]) + "»").ToList();
            if (layout.Flags.Count == 0) missing.Add("видов замечаний П1…ПР");
            if (missing.Count > 0) throw new InvalidDataException(file + ": на листе «" + name + "» нет столбцов " + string.Join(", ", missing) + ". Нужна форма проверки РД (реестр комплектов).");
            if (target) {
                var cols = layout.Flags.Select(f => layout[f]).ToList();
                if (cols.Max() - cols.Min() != cols.Count - 1) throw new InvalidDataException(file + ": столбцы видов замечаний П1…ПР должны идти подряд.");
            }
            return layout;
        }
        // The rows of a check sheet, with «Элемент строки» carried down over every row as the formula does.
        static List<RdRow> Rows(RawSheet sheet, RdLayout layout, List<RdRow> orphans, List<int> numbers) {
            var rows = new List<RdRow>(); string elkey = "", elname = "";
            int last = sheet.Rows.Keys.DefaultIfEmpty(0).Max();
            for (int r = layout.First; r <= last; r++) {
                object[] values; if (!sheet.Rows.TryGetValue(r, out values)) values = new object[Width];
                var row = new RdRow { File = sheet.Path, Row = r, Values = values, Layout = layout };
                if (sheet.Date1904) foreach (string k in new[] { "date", "fixed" }) { int c = layout[k]; if (c > 0 && values[c - 1] is double && (double)values[c - 1] > 1000) values[c - 1] = (double)values[c - 1] + 1462; }
                string el = Txt(row["el"]), rd = Txt(row["rd"]);
                // The row of column numbers moved among the data (e.g. by a sort in Excel) is not a set.
                if (new[] { "el", "rd", "who" }.All(k => Txt(row[k]) == layout[k].ToString(Inv))) { numbers.Add(r); continue; }
                elkey = el != "" ? el : rd == "" ? "" : elkey; row.ElKey = elkey;
                elname = el != "" ? Txt(row["elname"]) : rd == "" ? "" : elname; row.ElName = elname;
                if (rd != "") rows.Add(row); else if (HasReview(row) && orphans != null) orphans.Add(row);
            }
            return rows;
        }
        static RdVed ReadVed(string path) {
            var sheet = Sheet(path, "Ведомость", true);
            var ved = new RdVed { Sheet = sheet };
            var header = sheet.Rows.Where(r => r.Key <= 15).FirstOrDefault(r => XlsxReader.Normal(r.Value[3]) == "ШИФР РД" && XlsxReader.Normal(r.Value[0]) == "НОМЕР ЭЛЕМЕНТА");
            if (header.Value == null) throw new InvalidDataException(Path.GetFileName(path) + ": на листе «Ведомость» нет шапки ведомости элементов РД («Номер элемента» в A, «Шифр РД» в D).");
            ved.First = header.Key + 1;
            if (Txt(sheet.At(ved.First, 4)) == "4" && Txt(sheet.At(ved.First, 1)) == "1") ved.First++;
            ved.Last = sheet.Rows.Where(r => r.Key >= ved.First && r.Value.Take(30).Any(v => Txt(v) != "")).Select(r => r.Key).DefaultIfEmpty(ved.First - 1).Max();
            string above = "";
            for (int r = ved.First; r <= ved.Last; r++) {
                string el = Txt(sheet.At(r, 1)), code = Txt(sheet.At(r, 4));
                above = el != "" ? el : code == "" ? "" : above; ved.Element[r] = above;
                if (code != "") { ved.Sets++; if (!ved.CodeRow.ContainsKey(code)) ved.CodeRow[code] = r; }
                if (above != "") {
                    if (!ved.ElementRow.ContainsKey(above)) { ved.ElementRow[above] = r; ved.ElementSets[above] = 0; }
                    ved.ElementSets[above]++;
                    if (r == ved.First || !RdBook.Same(ved.Element[r - 1], above)) ved.Elements++;
                }
            }
            return ved;
        }
        static RdActs ReadActs(string path) {
            var acts = new RdActs(); var sheet = Sheet(path, acts.Name, false);
            if (sheet == null) return acts;
            foreach (var row in sheet.Rows.Where(r => r.Key <= 30)) {
                var heads = row.Value.Select(XlsxReader.Normal).ToList();
                int el = heads.IndexOf("НОМЕР ЭЛЕМЕНТА") + 1, who = heads.IndexOf("ПРОВЕРИЛ") + 1, date = heads.IndexOf("ДАТА ПРОВЕРКИ") + 1;
                if (el == 0 || who == 0 || date == 0) continue;
                var flags = heads.Select((h, i) => new { h, i }).Where(x => Regex.IsMatch(x.h, @"^А\d(?:\.|\s|$)")).Select(x => x.i + 1).ToList();
                int fixedCol = row.Value.Select((h, i) => new { h, i }).Where(x => Rules.IsResolvedHeader(x.h)).Select(x => x.i + 1).FirstOrDefault();
                acts.Exists = true; acts.El = el; acts.Who = who; acts.Date = date;
                acts.Status = heads.IndexOf(XlsxReader.Normal("Результат проверки акта")) + 1;
                acts.First = row.Key + 1;
                if (Txt(sheet.At(acts.First, el)) == el.ToString(Inv)) acts.First++;
                acts.Last = Math.Max(100, sheet.Rows.Keys.DefaultIfEmpty(0).Max());
                for (int r = acts.First; r <= acts.Last; r++) {
                    object[] v; if (!sheet.Rows.TryGetValue(r, out v)) continue;
                    string element = Txt(v[el - 1]); if (element == "") continue;
                    string whoText = Rules.CleanName(v[who - 1]); double? day = XlsxReader.DateSerial(v[date - 1]);
                    // The result as the formula of the sheet gives it: not checked without a reviewer or a date, remarks while a mark is not resolved.
                    string status = whoText == "" || Txt(v[date - 1]) == "" ? "не проверен" : flags.Any(c => Flag(v[c - 1]) is double && (double)Flag(v[c - 1]) == 1) && !(fixedCol > 0 && Rules.IsResolved(v[fixedCol - 1])) ? "замечания" : "соответствует";
                    if (status == "соответствует") acts.Conform++; else if (status == "замечания") acts.Remarks++;
                    string key = XlsxReader.Text(v[el - 1]);
                    if (!acts.ElementStatus.ContainsKey(key)) acts.ElementStatus[key] = status;
                    if (whoText != "" && day.HasValue) acts.Checks.Add(Tuple.Create(whoText, day.Value));
                }
                return acts;
            }
            return acts;
        }
        public static RdBook Collect(IList<string> sources, string sheetName, DateTime asOf, Action<int, string> progress, CancellationToken token, MergeOptions options, MergeResult result, DateTime importDate) {
            var book = new RdBook { MainName = sheetName };
            var byKey = new Dictionary<string, RdSet>(StringComparer.Ordinal);
            var marks = new List<KeyValuePair<RdRow, string>>();
            var spelling = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int fi = 0; fi < sources.Count; fi++) {
                Check(token);
                progress((int)(40.0 * fi / sources.Count), "Чтение " + Path.GetFileName(sources[fi]));
                var sheet = Sheet(sources[fi], sheetName, true);
                var layout = Layout(sheet, sheetName, fi == 0);
                var orphans = new List<RdRow>();
                var numbers = new List<int>();
                var rows = Rows(sheet, layout, orphans, numbers);
                // In the consolidated book its formulas would be broken by such a row: nothing is written until it is back under the header.
                if (fi == 0 && numbers.Count > 0) throw new InvalidDataException(Path.GetFileName(sources[0]) + ": на листе «" + sheetName + "» строка с номерами столбцов (1, 2, 3 …) стоит среди данных, в строке " + numbers[0] + " — вероятно, её переместила сортировка в Excel. Верните её под шапку (строка " + (layout.Header + 1) + ") или удалите и повторите сборку. Сводный документ не изменён.");
                foreach (int n in numbers) result.Problems.Add(new Problem { File = sources[fi], Row = n, Document = "", Detail = "Строка с номерами столбцов (1, 2, 3 …) стоит среди данных — вероятно, её переместила сортировка в Excel. Она не считается комплектом." });
                if (fi == 0) { book.Target = sheet; book.Layout = layout; }
                foreach (var orphan in orphans) AddProblem(result, orphan, fi == 0 ? "Есть данные проверки, но нет шифра РД. Строка сохраняется на месте и в статистику не входит." : "Есть данные проверки, но нет шифра РД: строка не перенесена в сводный документ.");
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var row in rows) {
                    Check(token);
                    string key = Key(row);
                    if (!seen.Add(key)) AddProblem(result, row, "Повтор комплекта с тем же шифром РД и «Изм.»: данные объединены с первой строкой.");
                    // One spelling of a reviewer: the first met (the consolidated book first) for every way of writing the same name.
                    string who = Rules.CleanName(row["who"]);
                    if (who != "" && !RdBook.IsExpert(who)) {
                        string known, nameKey = Rules.NameKey(who);
                        if (!spelling.TryGetValue(nameKey, out known)) spelling[nameKey] = who;
                        else if (known != who) { row.Values[layout["who"] - 1] = known; string merge = "«" + who + "» → «" + known + "»"; if (!result.PeopleMerges.Contains(merge)) result.PeopleMerges.Add(merge); }
                    }
                    else if (RdBook.IsExpert(who)) row.Values[layout["who"] - 1] = RdBook.Expert;
                    RdSet set;
                    if (!byKey.TryGetValue(key, out set)) {
                        set = new RdSet { Key = key, IsNew = fi > 0, TargetRow = fi == 0 ? row : null, SheetRow = fi == 0 ? row.Row : 0, ElKey = row.ElKey };
                        foreach (string k in RdLayout.Identity) set[k] = row[k];
                        if (Txt(set["el"]) == "") set["el"] = row.ElKey;
                        if (Txt(set["elname"]) == "") set["elname"] = row.ElName;
                        byKey[key] = set; book.Sets.Add(set);
                    }
                    if (fi == 0 || HasReview(row)) set.Reviews.Add(row);
                    if (!HasReview(row)) continue;
                    if (Txt(row["date"]) != "" && !XlsxReader.DateSerial(row["date"]).HasValue) AddProblem(result, row, "Дата проверки не распознана.");
                    if (Txt(row["who"]) != "" && !RdBook.IsExpert(row["who"]) && Txt(row["date"]) == "" && !options.CurrentDateForNewReviews) AddProblem(result, row, "Указан проверяющий, но нет даты проверки.");
                    foreach (string f in layout.Flags) if (Flag(row[f]) is string) marks.Add(new KeyValuePair<RdRow, string>(row, f));
                }
                progress((int)(40.0 * (fi + 1) / sources.Count), "Прочитано комплектов: " + rows.Count + (orphans.Count > 0 ? ", строк без шифра: " + orphans.Count : ""));
            }
            // «да», «+», «х» in the flags become 1 only when the operator agrees, as in the PD register.
            foreach (var g in marks.GroupBy(m => XlsxReader.Normal(m.Key[m.Value]))) {
                var cells = g.ToList(); string shown = Txt(cells[0].Key[cells[0].Value]);
                bool toOne = options.ConvertMark != null && options.ConvertMark(shown, cells.Count);
                foreach (var m in cells) {
                    if (toOne) m.Key.Values[m.Key.Layout[m.Value] - 1] = 1.0;
                    else AddProblem(result, m.Key, "Недопустимая отметка в «" + m.Value + "»: «" + shown + "». Допустимы 1, 0 или пусто.");
                }
                if (toOne) result.ConvertedMarks.Add("«" + shown + "» → 1 (ячеек: " + cells.Count + ")");
            }
            foreach (var set in book.Sets) { Check(token); MergeSet(set, book.Layout, options, result, importDate); }
            Place(book);
            book.Ved = ReadVed(sources[0]);
            book.Acts = ReadActs(sources[0]);
            Evaluate(book, asOf);
            result.Rd = true; result.Documents = book.Sets.Count; result.Elements = book.Elements.Count;
            result.Reviewed = book.Sets.Count(s => s.Who != "");
            result.WithIssues = book.Sets.Count(s => s.Cnt > 0 || Txt(s["text"]) != "");
            result.Boxes = book.Elements.Count(e => e.Date.HasValue && e.Date.Value > 0 && e.Date.Value <= asOf.Date.ToOADate());
            return book;
        }
        // The rules of the PD register for one set: every «1» is kept, different texts are joined, the reviewer and the date of the consolidated book stay.
        static void MergeSet(RdSet set, RdLayout layout, MergeOptions options, MergeResult result, DateTime importDate) {
            var reviews = set.Reviews.Where(HasReview).ToList();
            var kept = set.TargetRow != null && HasReview(set.TargetRow) ? set.TargetRow : null;
            string document = Txt(set["rd"]) + (Txt(Revision(set["rev"])) != "" && Txt(Revision(set["rev"])) != "0" ? ", изм. " + Txt(set["rev"]) : "");
            if (reviews.Count == 0) { foreach (string k in ReviewKeys(layout)) set[k] = set.TargetRow == null ? null : set.TargetRow[k]; return; }
            var named = reviews.Where(r => Txt(r["who"]) != "").ToList();
            var winner = Latest(named.Count > 0 ? named : reviews);
            set["who"] = Txt(winner["who"]) == "" ? null : Rules.CleanName(winner["who"]);
            double? date = XlsxReader.DateSerial(winner["date"]); set["date"] = date.HasValue ? (object)date.Value : winner["date"];
            if (kept != null && Txt(kept["who"]) != "") {
                set["who"] = Rules.CleanName(kept["who"]);
                if (Txt(kept["date"]) != "") { double? d = XlsxReader.DateSerial(kept["date"]); set["date"] = d.HasValue ? (object)d.Value : kept["date"]; }
                else { var same = Latest(named.Where(r => r != kept && Rules.NameKey(Txt(r["who"])) == Rules.NameKey(Txt(kept["who"])))); double? d = same == null ? null : XlsxReader.DateSerial(same["date"]); set["date"] = same == null ? null : d.HasValue ? (object)d.Value : same["date"]; }
            }
            if (options.CurrentDateForNewReviews && Txt(set["who"]) != "") {
                double? saved = kept != null && Txt(kept["who"]) != "" ? XlsxReader.DateSerial(kept["date"]) : null;
                if (saved.HasValue) set["date"] = saved.Value; else { set["date"] = importDate.Date.ToOADate(); result.DatesAssignedToday++; }
            }
            foreach (string f in layout.Flags) {
                var flags = reviews.Select(r => Flag(r[f])).ToList();
                set[f] = flags.Any(v => v is double && (double)v == 1) ? (object)1.0 : flags.FirstOrDefault(v => v != null);
            }
            foreach (string k in new[] { "text", "note" }) {
                string joined = MergeEngine.JoinTexts(reviews.Select(r => Txt(r[k]).Replace("\r\n", "\n")));
                if (joined != null && joined.Length > 32767) throw new InvalidDataException("Объединённый текст превышает предел ячейки Excel (32767 символов): " + document + ". Сократите тексты перед повторной сборкой.");
                set[k] = joined;
            }
            // «Факт СМР»: the value of the consolidated book, otherwise of the latest dated check; «Замечания устранены»: the last filled value in file order.
            var smr = kept != null && Txt(kept["smr"]) != "" ? kept : Latest(reviews.Where(r => Txt(r["smr"]) != ""));
            set["smr"] = smr == null ? null : smr["smr"];
            var fixedRow = reviews.LastOrDefault(r => Txt(r["fixed"]) != "");
            set["fixed"] = fixedRow == null ? null : fixedRow["fixed"];
            // A difference already in the consolidated book is not reported again.
            foreach (string k in new[] { "who", "date", "smr" }.Concat(layout.Flags).Concat(new[] { "text", "note" })) {
                if (k == "date" && options.CurrentDateForNewReviews) continue;
                var active = reviews.Where(r => Txt(r["who"]) != "" || Txt(r[k]) != "").ToList();
                var values = active.Select(r => k == "date" ? Pretty(k, r[k]) : XlsxReader.Normal(layout.Flags.Contains(k) ? Flag(r[k]) : r[k])).Distinct().ToList();
                if (values.Count < 2) continue;
                if (kept != null && active.Contains(kept) && active.Where(r => r != kept).All(r => Contained(k, r[k], kept[k], layout))) continue;
                string decision = (k == "who" || k == "date") && kept != null && Txt(kept["who"]) != "" ? "Проверяющий и дата сводного документа сохранены: из присланных файлов обновляются отметки и замечания." :
                    k == "who" || k == "date" ? "Проверяющий и дата из последней датированной проверки; при равной дате — последний файл в списке." :
                    k == "smr" ? "Значение сводного документа или последней датированной проверки; требуется сверка." :
                    layout.Flags.Contains(k) ? "Все отметки 1 сохранены; различия показаны в окне конфликтов." : "Все разные исходные тексты объединены без добавления фамилий и дат.";
                string field = k == "fixed" ? Rules.ResolvedHeader : layout.Flags.Contains(k) ? k : RdLayout.Names[k];
                result.Conflicts.Add(new Conflict { Document = document, Field = layout.L(k) + " — " + field, Details = string.Join("\n", active.Select(r => Path.GetFileName(r.File) + ", строка " + r.Row + ", " + (Txt(r["who"]) == "" ? "без проверяющего" : Txt(r["who"])) + ": " + Pretty(k, r[k]))), Decision = decision });
            }
        }
        static bool Contained(string key, object value, object target, RdLayout layout) {
            if (Txt(value) == "") return true;
            if (key == "date") return Pretty(key, value) == Pretty(key, target);
            if (layout.Flags.Contains(key)) { object f = Flag(value); return f is double && (double)f == 0 || XlsxReader.Normal(f) == XlsxReader.Normal(Flag(target)); }
            if (key == "text" || key == "note") return MergeEngine.JoinTexts(new[] { Txt(target), Txt(value) }) == MergeEngine.JoinTexts(new[] { Txt(target) });
            return XlsxReader.Normal(value) == XlsxReader.Normal(target);
        }
        // New sets go after the last filled row of the consolidated book, in the order of the files; the first set of an element
        // gets its number and name when the row above belongs to another element. The formulas reach 50 rows past the data.
        static void Place(RdBook book) {
            var layout = book.Layout; var target = book.Target;
            var inputs = RdLayout.Identity.Concat(RdLayout.Review).Concat(layout.Flags).Select(k => layout[k]).Where(c => c > 0).ToList();
            int lastUsed = target.Rows.Where(r => r.Key >= layout.First && inputs.Any(c => c <= r.Value.Length && Txt(r.Value[c - 1]) != "")).Select(r => r.Key).DefaultIfEmpty(layout.First - 1).Max();
            string above = ElementAbove(book, lastUsed + 1);
            int next = lastUsed + 1;
            foreach (var set in book.Sets.Where(s => s.IsNew)) {
                set.SheetRow = next++;
                set.WriteElement = !RdBook.Same(set.ElKey, above);
                if (!set.WriteElement) { set["el"] = null; set["elname"] = null; }
                above = set.ElKey;
            }
            // The writer widens Last to the formula rows the sheet already has.
            book.Last = Math.Max(next - 1, lastUsed) + 50;
            if (book.Last > 1048576) throw new InvalidDataException("На листе «" + book.MainName + "» недостаточно строк Excel.");
        }
        // «Элемент строки» of the row above a row of the consolidated book, carried down as the formula does.
        static string ElementAbove(RdBook book, int row) {
            string elkey = ""; var layout = book.Layout;
            for (int r = layout.First; r < row; r++) {
                string el = Txt(book.Target.At(r, layout["el"])), rd = Txt(book.Target.At(r, layout["rd"]));
                elkey = el != "" ? el : rd == "" ? "" : elkey;
            }
            return elkey;
        }
        // The formulas of the main sheet and of «Элементы» calculated for the cached values, the report numbers and the PDF.
        static void Evaluate(RdBook book, DateTime asOf) {
            var layout = book.Layout; var ved = book.Ved;
            var bySheetRow = book.Sets.ToDictionary(s => s.SheetRow);
            string elkey = ""; int order = 0; var orders = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var rowKeys = new List<KeyValuePair<int, string>>();
            int bottom = Math.Max(book.Last, book.Target.Rows.Keys.DefaultIfEmpty(0).Max());
            for (int r = layout.First; r <= bottom; r++) {
                RdSet set; bySheetRow.TryGetValue(r, out set);
                string el = Txt(set != null ? (set.IsNew ? set["el"] : book.Target.At(r, layout["el"])) : book.Target.At(r, layout["el"]));
                string rd = set != null ? Txt(set["rd"]) : Txt(book.Target.At(r, layout["rd"]));
                elkey = el != "" ? el : rd == "" ? "" : elkey;
                if (elkey != "" && !orders.ContainsKey(elkey)) { orders[elkey] = ++order; book.RowOrder[r] = order; }
                book.RowElKey[r] = elkey;
                if (set != null) { set.ElKey = elkey; set.Order = elkey == "" ? 0 : orders[elkey]; }
                rowKeys.Add(new KeyValuePair<int, string>(r, elkey));
            }
            foreach (var set in book.Sets) {
                string code = Txt(set["rd"]); int vr;
                if (!ved.CodeRow.TryGetValue(code, out vr)) { set.Ve = "нет в ВЭ РД"; set.Te = "—"; set.TeDate = ""; set.TeLetter = ""; }
                else {
                    string element = ved.Element[vr];
                    set.Ve = string.Equals(element, set.ElKey, StringComparison.OrdinalIgnoreCase) ? "да" : "элемент " + element;
                    set.Te = ved.At(vr, 9) ?? (object)0.0; set.TeDate = Txt(ved.At(vr, 11)) == "" ? (object)"" : ved.At(vr, 11); set.TeLetter = Txt(ved.At(vr, 12));
                }
                set.Cnt = layout.Flags.Count(f => { object v = Flag(set[f]); return v is double && (double)v == 1; }) + (Txt(set["note"]) != "" ? 1 : 0);
                set.Date = XlsxReader.DateSerial(set["date"]);
                set.Status = set.Ve != "да" ? "не соответствует ВЭ РД" : set.Who == "" ? "не проверен" : set.Cnt > 0 && !Rules.IsResolved(set["fixed"]) ? "замечания" :
                    !RdBook.Agreed.Any(a => string.Equals(Txt(set.Te), a, StringComparison.OrdinalIgnoreCase)) ? "не согласован ТЭ" : "готово к ВПР";
            }
            // The current edition: no set with the same code has a greater «Изм.» (numbers only, as COUNTIFS with ">"&N()).
            foreach (var g in book.Sets.GroupBy(s => Txt(s["rd"]), StringComparer.OrdinalIgnoreCase)) {
                var list = g.ToList();
                foreach (var s in list) { object rev = Revision(s["rev"]); double n = rev is double ? (double)rev : 0; s.Latest = !list.Any(o => { object x = Revision(o["rev"]); return x is double && (double)x > n; }); }
            }
            foreach (var key in rowKeys.Where(k => k.Value != "").Select(k => k.Value).Distinct(StringComparer.OrdinalIgnoreCase)) {
                var e = new RdElement { Key = key };
                e.Live.AddRange(book.Sets.Where(s => s.Latest && string.Equals(s.ElKey, key, StringComparison.OrdinalIgnoreCase)));
                int vr;
                if (ved.ElementRow.TryGetValue(key, out vr)) { e.Name = ved.At(vr, 2) ?? (object)0.0; e.Dgp = ved.At(vr, 7) ?? (object)0.0; e.ActVe = ved.At(vr, 14) ?? (object)0.0; e.Vpr = ved.At(vr, 18) ?? (object)0.0; e.InVe = ved.ElementSets[key]; }
                else {
                    var first = book.Sets.FirstOrDefault(s => string.Equals(s.ElKey, key, StringComparison.OrdinalIgnoreCase));
                    e.Name = first == null ? "" : (object)(first.IsNew ? first["elname"] : book.Target.At(first.SheetRow, layout["elname"])) ?? 0.0; e.Dgp = ""; e.ActVe = ""; e.Vpr = "";
                }
                e.Presented = e.Live.Count(s => s.Ve == "да"); e.Missing = Math.Max(0, e.InVe - e.Presented); e.NotInVe = e.Live.Count(s => s.Ve != "да");
                e.Checked = e.Live.Count(s => s.Ve == "да" && s.Status != "не проверен" && !s.Expert);
                e.Remarks = e.Live.Count(s => s.Status == "замечания"); e.NotAgreed = e.Live.Count(s => s.Status == "не согласован ТЭ"); e.Ready = e.Live.Count(s => s.Status == "готово к ВПР");
                e.Status = e.Missing == 0 && e.NotInVe == 0 && e.Ready == e.InVe ? "готов к ВПР" : e.Missing > 0 ? "неполный комплект" : e.NotInVe > 0 ? "расхождение с ВЭ РД" : e.Remarks > 0 ? "замечания" : e.NotAgreed > 0 ? "не согласован ТЭ" : "в проверке";
                string act; e.Act = book.Acts.ElementStatus.TryGetValue(key, out act) ? act : "не поступал";
                e.Expert = e.Live.Count(s => s.Expert);
                // An element is checked on the day of its last current set when every current set has a date and none is a remark of ТЭ.
                if (e.Live.All(s => s.Date.HasValue) && e.Expert == 0) {
                    e.Date = e.Live.Count == 0 ? 0 : e.Live.Max(s => s.Date.Value);
                    var by = e.Live.LastOrDefault(s => s.Date == e.Date); e.Who = by == null ? "" : by.Who;
                } else e.Who = "";
                book.Elements.Add(e);
            }
            // Reviewers of the reports: everyone with a dated check of a set or an act, except ТЭ; the days: working days from the first check, check days and the report date.
            var team = book.Sets.Where(s => s.Who != "" && !s.Expert && s.Date.HasValue).ToList();
            foreach (var s in team) Activity(book, s.Who).Add(s.Date.Value);
            foreach (var c in book.Acts.Checks.Where(c => !RdBook.IsExpert(c.Item1))) Activity(book, c.Item1).Add(c.Item2);
            book.People.AddRange(book.Activity.Keys.OrderBy(v => v, StringComparer.CurrentCulture));
            double report = asOf.Date.ToOADate();
            var dates = team.Select(s => s.Date.Value).Where(d => d <= report).ToList();
            double start = dates.Count == 0 ? report : dates.Min();
            if (report - start > 730) throw new InvalidDataException("Период проверки больше двух лет. Проверьте даты в столбце «Дата проверки».");
            var days = new SortedSet<double>(dates) { report };
            for (double d = start; d <= report; d++) { var day = DateTime.FromOADate(d).DayOfWeek; if (day != DayOfWeek.Saturday && day != DayOfWeek.Sunday) days.Add(d); }
            book.Days.AddRange(days);
        }
        static SortedSet<double> Activity(RdBook book, string who) {
            SortedSet<double> days; if (!book.Activity.TryGetValue(who, out days)) book.Activity[who] = days = new SortedSet<double>();
            return days;
        }
        public static MergeResult Run(IList<string> files, string output, string mainName, DateTime asOf, Action<int, string> progress, CancellationToken token, MergeOptions options) {
            output = Path.GetFullPath(output);
            if (Path.GetExtension(output).ToLowerInvariant() != ".xlsx") throw new InvalidOperationException("Результат необходимо сохранить с расширением .xlsx.");
            if (!File.Exists(output)) throw new InvalidOperationException("Выберите существующий сводный документ РД (форма проверки РД с листами «" + mainName + "» и «Ведомость»).");
            string before;
            try { using (var check = new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { } before = MergeEngine.Digest(output); }
            catch (IOException) { throw new IOException("Закройте сводный документ в Excel и повторите сборку: файл занят другой программой."); }
            var sources = new List<string> { output };
            foreach (string f in files.Select(Path.GetFullPath)) if (!sources.Any(s => string.Equals(s, f, StringComparison.OrdinalIgnoreCase))) sources.Add(f);
            var result = new MergeResult { Files = sources.Count - 1 };
            var book = Collect(sources, mainName, asOf, progress, token, options, result, DateTime.Today);
            if (options.CurrentDateForNewReviews) progress(42, "Текущая дата назначена новым проверкам: " + result.DatesAssignedToday + ". Даты уже внесённых проверок сохранены.");
            string temp = Path.Combine(Path.GetDirectoryName(output), "~merge-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try {
                Check(token); progress(45, "Запись комплектов и формул");
                new XlsxWriter(options).SaveRd(output, temp, book, asOf, progress, token);
                Check(token);
                if (!File.Exists(output) || MergeEngine.Digest(output) != before) throw new IOException("Сводный документ изменился во время сборки. Запись отменена; повторите сборку с актуальным файлом.");
                string backup = output + ".backup-" + DateTime.Now.ToString("yyyyMMdd-HHmmssfff", Inv);
                File.Replace(temp, output, backup); progress(97, "Предыдущий результат сохранён: " + Path.GetFileName(backup));
                result.Output = output;
                try {
                    var pdf = options.Pdf ?? new PdfSettings(); pdf.WorkSheet = RdWork;
                    foreach (var file in PdfReport.Write(output, asOf, pdf)) {
                        result.Pdfs.Add(file.Path);
                        progress(99, file.Path == file.Requested ? "PDF: " + Path.GetFileName(file.Path) : "PDF «" + Path.GetFileName(file.Requested) + "» открыт в другой программе, создан «" + Path.GetFileName(file.Path) + "»");
                    }
                    result.Pdf = result.Pdfs.FirstOrDefault();
                } catch (Exception ex) { result.PdfError = ex.Message; progress(99, "PDF не создан: " + ex.Message); }
                progress(100, "Готово: " + result.Documents + " комплектов, " + result.Elements + " элементов, " + result.Conflicts.Count + " конфликтов, " + result.Problems.Count + " проблем данных");
                return result;
            } finally { if (File.Exists(temp)) { try { File.Delete(temp); } catch { } } }
        }
        public const string RdWork = "Итоги проверки";
    }
}
