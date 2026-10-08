using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Text.RegularExpressions;
using System.Security.Cryptography;

namespace ReviewMerge {
    public sealed class Conflict {
        public string Document { get; set; }
        public string Field { get; set; }
        public string Details { get; set; }
        public string Decision { get; set; }
    }
    public sealed class Problem {
        public string File, Document, Detail;
        public int Row;
    }
    public sealed class MergedRow {
        public object[] Values;
        public SourceRow TargetRow;
        public readonly List<SourceRow> Reviews = new List<SourceRow>();
        public readonly List<SourceRow> TargetDuplicates = new List<SourceRow>();
    }
    // Rows of one box (or of one volume without a box) where only a part is marked as checked.
    public sealed class IncompleteBox {
        public string Box, Volume, Reviewer;
        public double Date;
        public int Checked, Total;
        public readonly List<string> Missing = new List<string>();
        internal readonly Dictionary<MergedRow, MergedRow> Rows = new Dictionary<MergedRow, MergedRow>();
        public override string ToString() {
            return (Box != "" ? "Короб " + Box : "Том без номера короба " + Volume) + ": отмечено " + Checked + " из " + Total + " строк; не отмечены: " + string.Join(", ", Missing.Take(3)) + (Missing.Count > 3 ? " и ещё " + (Missing.Count - 3) : "") +
                " (" + Reviewer + ", " + DateTime.FromOADate(Date).ToString("dd.MM.yyyy", CultureInfo.InvariantCulture) + ")";
        }
    }
    public sealed class MergeOptions {
        public bool CurrentDateForNewReviews, MatchSurnames;
        // A reviewer name from the files and a known person with the same surname; true treats them as one person.
        public Func<string, string, bool> SamePerson;
        // A saved answer for such a pair, if any: initials that agree merge people without a question unless the operator said "no" before.
        public Func<string, string, bool?> SavedSamePerson;
        // A mark other than 1 or 0 in M:R ("да", "+", "х") and the number of cells with it; true turns it into 1.
        public Func<string, int, bool> ConvertMark;
        // Receives the partly checked boxes and returns those whose remaining rows are filled without remarks.
        public Func<IList<IncompleteBox>, IList<IncompleteBox>> FillIncomplete;
    }
    public sealed class MergeResult {
        public int Files, Documents, Reviewed, WithIssues, Boxes, DatesAssignedToday, FilledRows;
        public string Output;
        public readonly List<Conflict> Conflicts = new List<Conflict>();
        public readonly List<Problem> Problems = new List<Problem>();
        public readonly List<SourceRow> History = new List<SourceRow>();
        public readonly List<MergedRow> Rows = new List<MergedRow>();
        public readonly List<DateTime> Dates = new List<DateTime>();
        public readonly List<IncompleteBox> IncompleteBoxes = new List<IncompleteBox>();
        public readonly List<IncompleteBox> FilledBoxes = new List<IncompleteBox>();
        public readonly List<string> PeopleMerges = new List<string>();
        public readonly List<string> ConvertedMarks = new List<string>();
        public int StartRow, HeaderRow;
    }
    public static class MergeEngine {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static readonly string[] Fields = { "Проверил", "Дата проверки", "Вид документации", "Короб", "Название тома", "Шифр тома", "Количество страниц", "Подписи и печати на титуле", "Контрольная сумма", "Подписи в ИУЛ", "Ошибки в описи", "Примечание" };
        static void Check(CancellationToken token) { token.ThrowIfCancellationRequested(); }
        static string Txt(object v) { return XlsxReader.Text(v); }
        static void AddProblem(MergeResult r, SourceRow row, string detail) {
            r.Problems.Add(new Problem { File = row.File, Row = row.Row, Document = Txt(row.Values[2]), Detail = detail });
        }
        static bool HasReview(SourceRow r) { return r.Values.Skip(8).Any(v => Txt(v) != ""); }
        static object Flag(object value) {
            string s = Txt(value);
            if (s == "") return null;
            double n;
            if (double.TryParse(s.Replace(',', '.'), NumberStyles.Float, Inv, out n) && (n == 0 || n == 1)) return n;
            return s;
        }
        static string Pretty(int col, object value) {
            double? d = col == 9 ? XlsxReader.DateSerial(value) : null;
            return d.HasValue ? DateTime.FromOADate(d.Value).ToString("dd.MM.yyyy", Inv) : (Txt(value) == "" ? "(пусто)" : Txt(value));
        }
        static string SourceDescription(SourceRow r, int col) {
            return Path.GetFileName(r.File) + ", строка " + r.Row + ", " + (Txt(r.Values[8]) == "" ? "без проверяющего" : Txt(r.Values[8])) + ": " + Pretty(col, r.Values[col]);
        }
        static SourceRow Latest(IEnumerable<SourceRow> rows) {
            return rows.Select((r, i) => new { Row = r, Order = i, Date = XlsxReader.DateSerial(r.Values[9]) ?? -1 })
                .OrderBy(x => x.Date).ThenBy(x => x.Order).Select(x => x.Row).LastOrDefault();
        }
        static readonly Regex Attribution = new Regex(@"^\[[^\]\n]+, (?:\d{2}\.\d{2}\.\d{4}|\(пусто\))\] ");
        static IEnumerable<string> SourceTexts(SourceRow r, int col) {
            string value = Txt(r.Values[col]).Replace("\r\n", "\n");
            if (value == "") yield break;
            if (!r.IsConsolidated || !Attribution.IsMatch(value)) { yield return value; yield break; }
            // Attribution added by an older version of the program is removed from recognised consolidated books only.
            foreach (string part in Regex.Split(value, @"\n\n(?=\[[^\]\n]+, (?:\d{2}\.\d{2}\.\d{4}|\(пусто\))\] )"))
                yield return Attribution.IsMatch(part) ? Regex.Replace(part, @"^\[[^\]\n]+\] ", "") : part;
        }
        static List<string> Paragraphs(string text) {
            return Regex.Split(text, @"\n[ \t]*\n\s*").Select(p => p.Trim('\n')).Where(p => p.Trim() != "").ToList();
        }
        // The lines of inner form a contiguous run of lines in outer, e.g. a list before a new item was appended.
        static bool Covers(string outer, string inner) { return ("\n" + outer + "\n").Contains("\n" + inner + "\n"); }
        static bool NoteContained(object value, object target) {
            var known = Paragraphs(Txt(target).Replace("\r\n", "\n"));
            return Paragraphs(Txt(value).Replace("\r\n", "\n")).All(p => known.Any(n => Covers(n, p)));
        }
        static string JoinNotes(IEnumerable<SourceRow> rows, int col) {
            var texts = rows.SelectMany(r => SourceTexts(r, col)).Where(t => t.Trim() != "").ToList();
            if (texts.Count == 0) return null;
            if (texts.Distinct(StringComparer.Ordinal).Count() == 1) return texts[0];
            // Union by paragraph: a text that extends an earlier one replaces it instead of repeating it.
            var notes = new List<string>();
            foreach (string p in texts.SelectMany(Paragraphs)) {
                if (notes.Any(n => Covers(n, p))) continue;
                int at = notes.FindIndex(n => Covers(p, n));
                notes.RemoveAll(n => Covers(p, n));
                notes.Insert(at < 0 ? notes.Count : at, p);
            }
            // Only source text belongs in a remark: no reviewer names or dates are added.
            return string.Join("\n\n", notes);
        }
        static bool Contained(int col, object value, object target) {
            if (Txt(value) == "") return true;
            if (col == 9) return Pretty(9, value) == Pretty(9, target);
            if (col >= 12 && col < 18) { object f = Flag(value); return f is double && (double)f == 0 || XlsxReader.Normal(f) == XlsxReader.Normal(Flag(target)); }
            if (col >= 18) return NoteContained(value, target);
            return XlsxReader.Normal(value) == XlsxReader.Normal(target);
        }
        static bool Checked(MergedRow r) { return Txt(r.Values[8]) != "" && XlsxReader.DateSerial(r.Values[9]).HasValue; }
        static MergedRow LatestChecked(IEnumerable<MergedRow> rows) { return rows.Where(Checked).OrderByDescending(r => XlsxReader.DateSerial(r.Values[9]).Value).FirstOrDefault(); }
        // A box is normally checked as a whole: rows of a box (by their own box number or by the volume they belong to, such as ИУЛ rows) that stay unmarked while others are marked.
        static List<IncompleteBox> FindIncompleteBoxes(MergeResult result) {
            var volumeOf = new Dictionary<MergedRow, List<MergedRow>>();
            foreach (var g in result.Rows.GroupBy(r => Rules.VolumeKey(r.Values[2], r.Values[3]), StringComparer.Ordinal)) { var rows = g.ToList(); foreach (var r in rows) volumeOf[r] = rows; }
            Func<MergedRow, List<string>> own = r => Rules.BoxKeys(r.Values[10], r.Values[11]);
            Func<MergedRow, bool> partOfVolume = r => volumeOf[r].Any(v => !Rules.IsUl(v.Values[2], v.Values[3]));
            var groups = new Dictionary<string, List<MergedRow>>(StringComparer.Ordinal);var order = new List<string>();
            foreach (var r in result.Rows) {
                var boxes = own(r);
                if (boxes.Count == 0 && partOfVolume(r)) boxes = volumeOf[r].Select(own).FirstOrDefault(b => b.Count > 0) ?? boxes;
                string key = boxes.Count > 0 ? Rules.BoxList(boxes) : partOfVolume(r) ? "\u001f" + Rules.VolumeKey(r.Values[2], r.Values[3]) : null;
                if (key == null) continue;
                List<MergedRow> list; if (!groups.TryGetValue(key, out list)) { groups[key] = list = new List<MergedRow>(); order.Add(key); }
                list.Add(r);
            }
            var result2 = new List<IncompleteBox>();
            foreach (string key in order) {
                var rows = groups[key]; var source = LatestChecked(rows);
                if (source == null || rows.All(Checked)) continue;
                var main = rows.FirstOrDefault(r => !Rules.IsUl(r.Values[2], r.Values[3])) ?? rows[0];
                var box = new IncompleteBox { Box = key[0] == '\u001f' ? "" : Rules.BoxListDisplay(key), Volume = Txt(main.Values[2]), Reviewer = Txt(source.Values[8]), Date = XlsxReader.DateSerial(source.Values[9]).Value, Checked = rows.Count(Checked), Total = rows.Count };
                // A missing part takes the reviewer and date of its own volume when that volume has a checked part, otherwise of the box.
                foreach (var r in rows.Where(r => !Checked(r))) { box.Rows[r] = LatestChecked(volumeOf[r]) ?? source; box.Missing.Add(Txt(r.Values[2])); }
                result2.Add(box);
            }
            return result2;
        }
        // Marks the remaining rows as checked without remarks: reviewer, date, kind and box of the checked part.
        static void Fill(MergeResult result, IncompleteBox box) {
            foreach (var pair in box.Rows) {
                MergedRow r = pair.Key, source = pair.Value; if (Checked(r)) continue;
                double date = XlsxReader.DateSerial(source.Values[9]).Value;
                if (Txt(r.Values[8]) == "") r.Values[8] = source.Values[8];
                if (!XlsxReader.DateSerial(r.Values[9]).HasValue) r.Values[9] = date;
                if (Txt(r.Values[10]) == "") r.Values[10] = source.Values[10];
                if (Txt(r.Values[11]) == "") r.Values[11] = source.Values[11];
                result.FilledRows++;
            }
        }
        public static MergeResult Collect(IList<string> files, string sheetName, DateTime asOf, Action<int,string> progress, CancellationToken token, MergeOptions options=null, bool existingTarget=false, DateTime? importDate=null) {
            if (files.Count == 0) throw new InvalidOperationException("Добавьте хотя бы один файл .xlsx.");
            options = options ?? new MergeOptions();
            bool currentDateForNewReviews = options.CurrentDateForNewReviews;
            var result = new MergeResult { Files = files.Count };
            var byKey = new Dictionary<string, MergedRow>(StringComparer.Ordinal);
            var marks = new List<KeyValuePair<SourceRow, int>>();
            for (int fi = 0; fi < files.Count; fi++) {
                Check(token);
                progress((int)(40.0 * fi / files.Count), "Чтение " + Path.GetFileName(files[fi]));
                SourceBook source = XlsxReader.Read(files[fi], sheetName);
                if (fi == 0) { result.StartRow = source.DataStart; result.HeaderRow = source.HeaderRow; }
                foreach (var orphan in source.Orphans) {
                    result.History.Add(orphan);
                    AddProblem(result, orphan, "Есть данные проверки, но нет названия документа. В сводном документе такая строка сохраняется на месте; строка из другого источника не переносится без идентификации документа. В статистику документов не включена.");
                }
                var seenInFile = new HashSet<string>();
                foreach (var row in source.Rows) {
                    Check(token);
                    string key = XlsxReader.Key(row);
                    if (!seenInFile.Add(key)) AddProblem(result, row, fi == 0 ? "Повтор записи в сводном документе с одинаковыми именем, форматом, CRC32 и временем выгрузки. Объединён с первой записью; обе строки получают одинаковые данные проверки." : "Повтор записи с одинаковыми именем, форматом, CRC32 и временем выгрузки. Объединён с первой записью.");
                    if (Txt(row.Values[7]) == "" && HasReview(row)) AddProblem(result, row, "Не указана контрольная сумма: идентификация выполнена по остальным реквизитам.");
                    MergedRow merged;
                    if (!byKey.TryGetValue(key, out merged)) {
                        merged = new MergedRow { Values = (object[])row.Values.Clone(), TargetRow = fi == 0 ? row : null };
                        byKey.Add(key, merged); result.Rows.Add(merged);
                    } else if (fi == 0 && merged.TargetRow != null) merged.TargetDuplicates.Add(row);
                    if (!HasReview(row)) continue;
                    result.History.Add(row); merged.Reviews.Add(row);
                    bool suppliedDate=currentDateForNewReviews&&Txt(row.Values[8])!=""&&!(existingTarget&&merged.TargetRow!=null&&Txt(merged.TargetRow.Values[8])!=""&&XlsxReader.DateSerial(merged.TargetRow.Values[9]).HasValue);
                    if (!suppliedDate&&Txt(row.Values[9]) != "" && !XlsxReader.DateSerial(row.Values[9]).HasValue)
                        AddProblem(result, row, "Дата проверки в источнике не распознана.");
                    if (!suppliedDate&&Txt(row.Values[8]) != "" && Txt(row.Values[9]) == "")
                        AddProblem(result, row, "Указан проверяющий, но отсутствует дата проверки.");
                    for (int col = 12; col < 18; col++) {
                        if (Flag(row.Values[col]) is string) marks.Add(new KeyValuePair<SourceRow, int>(row, col));
                    }
                    string kind = XlsxReader.Normal(row.Values[10]);
                    if (kind != "" && kind != "ИИ" && kind != "ПД" && kind != "ДПТ") AddProblem(result, row, "Неизвестный вид документации: «" + Txt(row.Values[10]) + "».");
                    bool validBox; Rules.ParseBoxes(row.Values[11], out validBox);
                    if (!validBox) AddProblem(result, row, "Номер короба не распознан: «" + Txt(row.Values[11]) + "». Укажите целое число; несколько коробов — через запятую или «и», короба подряд — через дефис (45-48).");
                }
                progress((int)(40.0 * (fi + 1) / files.Count), "Прочитано: " + source.Rows.Count + " документов, " + source.Orphans.Count + " строк без документа");
            }
            // Marks such as "да", "+" or "х" become 1 only when the operator agrees; otherwise they stay as written and are reported.
            foreach (var g in marks.GroupBy(m => XlsxReader.Normal(m.Key.Values[m.Value]))) {
                var cells = g.ToList(); string shown = Txt(cells[0].Key.Values[cells[0].Value]);
                bool toOne = options.ConvertMark != null && options.ConvertMark(shown, cells.Count);
                foreach (var m in cells) {
                    if (!toOne) { AddProblem(result, m.Key, "Недопустимая отметка в " + ExcelColumn(m.Value + 1) + ": «" + Txt(m.Key.Values[m.Value]) + "». Допустимы 1, 0 или пусто."); continue; }
                    if (m.Key.Written == null) m.Key.Written = (object[])m.Key.Values.Clone();
                    m.Key.Values[m.Value] = 1.0;
                }
                if (toOne) result.ConvertedMarks.Add("«" + shown + "» → 1 (ячеек: " + cells.Count + ")");
            }
            foreach (var merged in result.Rows) {
                Check(token);
                var reviews = merged.Reviews;
                if (reviews.Count == 0) { for (int k=8;k<20;k++) merged.Values[k]=null; continue; }
                var named = reviews.Where(r => Txt(r.Values[8]) != "").ToList();
                SourceRow winner = Latest(named.Count > 0 ? named : reviews);
                merged.Values[8] = Txt(winner.Values[8]) == "" ? null : Txt(winner.Values[8]);
                double? date = XlsxReader.DateSerial(winner.Values[9]);
                merged.Values[9] = date.HasValue ? (object)date.Value : winner.Values[9];
                for (int col = 10; col <= 11; col++) {
                    var chosen = reviews.FirstOrDefault(r => Txt(r.Values[col]) != "");
                    object value = chosen == null ? null : chosen.Values[col];
                    if (col == 10) value = Txt(value) == "" ? null : XlsxReader.Normal(value);
                    if (col == 11 && value != null) {
                        double n; if (double.TryParse(Txt(value), NumberStyles.Float, Inv, out n) && n > 0 && n == Math.Floor(n)) value = n;
                    }
                    merged.Values[col] = value;
                }
                for (int col = 12; col < 18; col++) {
                    var flags = reviews.Select(r => Flag(r.Values[col])).ToList();
                    merged.Values[col] = flags.Any(v => v is double && (double)v == 1) ? (object)1.0 : flags.FirstOrDefault(v => v != null);
                }
                // Fields already filled in the consolidated book stay as they are: re-sent files update only the fields right of the box (M:T).
                var kept = existingTarget ? merged.TargetRow : null;
                if (kept != null && Txt(kept.Values[8]) != "") {
                    merged.Values[8] = Txt(kept.Values[8]);
                    double? keptDate = XlsxReader.DateSerial(kept.Values[9]);
                    if (Txt(kept.Values[9]) != "") merged.Values[9] = keptDate.HasValue ? (object)keptDate.Value : kept.Values[9];
                    else {
                        var same = Latest(named.Where(r => r != kept && Rules.NameKey(Txt(r.Values[8])) == Rules.NameKey(Txt(kept.Values[8]))));
                        double? sameDate = same == null ? null : XlsxReader.DateSerial(same.Values[9]);
                        merged.Values[9] = same == null ? null : sameDate.HasValue ? (object)sameDate.Value : same.Values[9];
                    }
                }
                merged.Values[18] = JoinNotes(reviews,18); merged.Values[19] = JoinNotes(reviews,19);
                if (Txt(merged.Values[18]).Length > 32767 || Txt(merged.Values[19]).Length > 32767)
                    throw new InvalidDataException("Объединённый текст замечаний превышает предел ячейки Excel (32767 символов): " + Txt(merged.Values[2]) + ". Источники сохранены; сократите тексты перед повторной сборкой.");
                if(currentDateForNewReviews&&Txt(merged.Values[8])!=""){
                    double? saved=existingTarget&&merged.TargetRow!=null&&Txt(merged.TargetRow.Values[8])!=""?XlsxReader.DateSerial(merged.TargetRow.Values[9]):null;
                    if(saved.HasValue)merged.Values[9]=saved.Value;
                    else {merged.Values[9]=(importDate??DateTime.Today).Date.ToOADate();result.DatesAssignedToday++;}
                }
                // A difference already taken into the consolidated book is not reported again on the next build.
                var target = existingTarget ? merged.TargetRow : null;
                for (int col = 8; col < 20; col++) {
                    if (col == 9 && currentDateForNewReviews) continue;
                    var active = reviews.Where(r => Txt(r.Values[8]) != "" || Txt(r.Values[col]) != "").ToList();
                    var distinct = active.Select(r => col == 9 ? Pretty(col,r.Values[col]) : XlsxReader.Normal(col >=12 && col<18 ? Flag(r.Values[col]) : r.Values[col])).Distinct().ToList();
                    if (distinct.Count < 2) continue;
                    if (target != null && active.Contains(target) && active.Where(r => r != target).All(r => Contained(col, r.Values[col], target.Values[col]))) continue;
                    string decision = col <= 11 && target != null && Txt(target.Values[col <= 9 ? 8 : col]) != "" ? "Значение сводного документа сохранено: из присланных файлов обновляются только отметки и замечания (M:T)." :
                        col == 8 || col == 9 ? "Проверяющий и дата из последней датированной проверки; при равной дате — последний файл в списке." :
                        col == 10 || col == 11 ? "Первое непустое значение в порядке файлов; требуется сверка реквизитов." :
                        col < 18 ? "Все отметки 1 сохранены; различия показаны в окне конфликтов." : "Все разные исходные тексты объединены без добавления фамилий и дат.";
                    result.Conflicts.Add(new Conflict { Document = Txt(merged.Values[2]), Field = ExcelColumn(col+1) + " — " + Fields[col-8],
                        Details = string.Join("\n",active.Select(r => SourceDescription(r,col))), Decision = decision });
                }
            }
            var incomplete = FindIncompleteBoxes(result);
            var selected = incomplete.Count > 0 && options.FillIncomplete != null ? options.FillIncomplete(incomplete) : null;
            foreach (var box in incomplete) {
                if (selected != null && selected.Contains(box)) { Fill(result, box); result.FilledBoxes.Add(box); }
                else result.IncompleteBoxes.Add(box);
            }
            result.Documents = result.Rows.Count;
            DateTime min = asOf.Date;
            foreach (var row in result.Rows) {
                if (Txt(row.Values[8]) != "") {
                    result.Reviewed++;
                    double? d = XlsxReader.DateSerial(row.Values[9]);
                    if (d.HasValue && DateTime.FromOADate(d.Value) < min) min = DateTime.FromOADate(d.Value);
                }
                if (row.Values.Skip(12).Take(6).Any(v => Txt(v) != "" && Txt(v) != "0") || Txt(row.Values[18]) != "" || Txt(row.Values[19]) != "") result.WithIssues++;
            }
            if ((asOf.Date-min.Date).TotalDays > 730) throw new InvalidDataException("Период проверки больше двух лет. Проверьте даты в столбце J.");
            for (var d=min.Date;d<=asOf.Date;d=d.AddDays(1)) result.Dates.Add(d);
            return result;
        }
        public static string ExcelColumn(int column) {
            string result=""; while(column>0) { column--; result=(char)('A'+column%26)+result; column/=26; } return result;
        }
        static string Digest(string path) {
            using(var stream=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.Read))using(var hash=SHA256.Create())
                return Convert.ToBase64String(hash.ComputeHash(stream));
        }
        public static MergeResult Run(IList<string> files,string output,string mainName,DateTime asOf,Action<int,string> progress,CancellationToken token,MergeOptions options=null) {
            options=options??new MergeOptions();
            output=Path.GetFullPath(output);
            if(Path.GetExtension(output).ToLowerInvariant()!=".xlsx") throw new InvalidOperationException("Результат необходимо сохранить с расширением .xlsx.");
            bool exists=File.Exists(output);string before=null;
            if(exists) {
                try {using(var check=new FileStream(output,FileMode.Open,FileAccess.ReadWrite,FileShare.None)){} before=Digest(output);}
                catch(IOException) {throw new IOException("Закройте сводный документ в Excel и повторите сборку: файл занят другой программой.");}
            }
            var sources=new List<string>();if(exists)sources.Add(output);
            foreach(string f in files.Select(Path.GetFullPath))if(!sources.Any(s=>string.Equals(s,f,StringComparison.OrdinalIgnoreCase)))sources.Add(f);
            var result=Collect(sources,mainName,asOf,progress,token,options,exists,DateTime.Today);result.Files=files.Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            if(options.CurrentDateForNewReviews)progress(42,"Текущая дата назначена новым проверенным записям: "+result.DatesAssignedToday+". Даты уже внесённых проверок сохранены.");
            if(result.FilledBoxes.Count>0)progress(43,"Заполнены без замечаний остальные строки: "+result.FilledRows+" в "+result.FilledBoxes.Count+" коробах.");
            if(result.IncompleteBoxes.Count>0)progress(43,"Коробов, проверенных не полностью: "+result.IncompleteBoxes.Count+". Их неотмеченные тома не входят в статистику.");
            string parent=Path.GetDirectoryName(output);Directory.CreateDirectory(parent);
            string temp=Path.Combine(parent,"~merge-"+Guid.NewGuid().ToString("N")+".xlsx");
            try {
                Check(token);progress(45,"Создание XLSX с формулами");
                new XlsxWriter(options).Save(sources[0],temp,mainName,asOf,result,progress,token);
                Check(token);
                if(exists) {
                    if(!File.Exists(output)||Digest(output)!=before)throw new IOException("Сводный документ изменился во время сборки. Запись отменена; повторите сборку с актуальным файлом.");
                    string backup=output+".backup-"+DateTime.Now.ToString("yyyyMMdd-HHmmssfff",Inv);
                    File.Replace(temp,output,backup);progress(97,"Предыдущий результат сохранён: "+Path.GetFileName(backup));
                } else {if(File.Exists(output))throw new IOException("Файл результата появился во время сборки. Повторите сборку.");File.Move(temp,output);}
                result.Output=output;progress(100,"Готово: "+result.Documents+" документов, "+result.Conflicts.Count+" конфликтов, "+result.Problems.Count+" проблем данных");return result;
            } finally { if(File.Exists(temp)) { try {File.Delete(temp);}catch{} } }
        }
    }
}
