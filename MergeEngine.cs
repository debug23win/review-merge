using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
        public double? FirstReviewDate;
        public readonly List<SourceRow> Reviews = new List<SourceRow>();
    }
    public sealed class MergeResult {
        public int Files, Documents, Reviewed, WithIssues, Boxes;
        public string Output;
        public readonly List<Conflict> Conflicts = new List<Conflict>();
        public readonly List<Problem> Problems = new List<Problem>();
        public readonly List<SourceRow> History = new List<SourceRow>();
        public readonly List<MergedRow> Rows = new List<MergedRow>();
        public readonly List<DateTime> Dates = new List<DateTime>();
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
        static string JoinNotes(IEnumerable<SourceRow> rows, int col) {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var notes = new List<string>();
            var tagged = new List<string>();
            foreach (var r in rows) {
                string value = Txt(r.Values[col]).Replace("\r\n", "\n");
                if (value == "") continue;
                string[] parts = r.IsConsolidated && Regex.IsMatch(value,@"^\[[^\]\n]+, (?:\d{2}\.\d{2}\.\d{4}|\(пусто\))\] ") ? Regex.Split(value,@"\n\n(?=\[[^\]\n]+, (?:\d{2}\.\d{2}\.\d{4}|\(пусто\))\] )") : new [] { value };
                foreach (string part in parts) {
                    bool annotated = r.IsConsolidated && Regex.IsMatch(part,@"^\[[^\]\n]+, (?:\d{2}\.\d{2}\.\d{4}|\(пусто\))\] ");
                    string text = annotated ? Regex.Replace(part,@"^\[[^\]\n]+\] ","") : part;
                    if (!seen.Add(text)) continue;
                    notes.Add(text);
                    tagged.Add(annotated ? part : "[" + (Txt(r.Values[8]) == "" ? "без проверяющего" : Txt(r.Values[8])) + ", " + Pretty(9, r.Values[9]) + "] " + part);
                }
            }
            if (notes.Count == 0) return null;
            if (notes.Count == 1) return notes[0];
            return string.Join("\n\n", tagged);
        }
        public static MergeResult Collect(IList<string> files, string sheetName, DateTime asOf, Action<int,string> progress, CancellationToken token) {
            if (files.Count == 0) throw new InvalidOperationException("Добавьте хотя бы один файл .xlsx.");
            var result = new MergeResult { Files = files.Count };
            var byKey = new Dictionary<string, MergedRow>(StringComparer.Ordinal);
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
                    if (!seenInFile.Add(key)) AddProblem(result, row, "Повтор записи с одинаковыми именем, форматом, CRC32 и временем выгрузки. Объединён с первой записью.");
                    if (Txt(row.Values[7]) == "" && HasReview(row)) AddProblem(result, row, "Не указана контрольная сумма: идентификация выполнена по остальным реквизитам.");
                    MergedRow merged;
                    if (!byKey.TryGetValue(key, out merged)) {
                        merged = new MergedRow { Values = (object[])row.Values.Clone(), TargetRow = fi == 0 ? row : null };
                        byKey.Add(key, merged); result.Rows.Add(merged);
                    }
                    double? actualDate = Txt(row.Values[8]) != "" ? XlsxReader.DateSerial(row.Values[9]) : null;
                    double? firstDate = row.PriorFirstDate.HasValue && actualDate.HasValue ? Math.Min(row.PriorFirstDate.Value,actualDate.Value) : actualDate;
                    if(firstDate.HasValue && (!merged.FirstReviewDate.HasValue || firstDate.Value < merged.FirstReviewDate.Value)) merged.FirstReviewDate=firstDate;
                    if (!HasReview(row)) continue;
                    result.History.Add(row); merged.Reviews.Add(row);
                    if (Txt(row.Values[9]) != "" && !XlsxReader.DateSerial(row.Values[9]).HasValue)
                        AddProblem(result, row, "Дата проверки не распознана; запись не включена в статистику по дням.");
                    if (Txt(row.Values[8]) != "" && Txt(row.Values[9]) == "")
                        AddProblem(result, row, "Указан проверяющий, но отсутствует дата проверки.");
                    for (int col = 12; col < 18; col++) {
                        object f = Flag(row.Values[col]);
                        if (f is string) AddProblem(result, row, "Недопустимая отметка в " + ExcelColumn(col + 1) + ": «" + Txt(f) + "». Допустимы 1, 0 или пусто.");
                    }
                    string kind = XlsxReader.Normal(row.Values[10]);
                    if (kind != "" && kind != "ИИ" && kind != "ПД" && kind != "ДПТ") AddProblem(result, row, "Неизвестный вид документации: «" + Txt(row.Values[10]) + "».");
                    double box;
                    if (Txt(row.Values[11]) != "" && (!double.TryParse(Txt(row.Values[11]), NumberStyles.Float, Inv, out box) || box <= 0 || box != Math.Floor(box)))
                        AddProblem(result, row, "Номер короба должен быть положительным целым числом.");
                }
                progress((int)(40.0 * (fi + 1) / files.Count), "Прочитано: " + source.Rows.Count + " документов, " + source.Orphans.Count + " строк без документа");
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
                merged.Values[18] = JoinNotes(reviews,18); merged.Values[19] = JoinNotes(reviews,19);
                if (Txt(merged.Values[18]).Length > 32767 || Txt(merged.Values[19]).Length > 32767)
                    throw new InvalidDataException("Объединённый текст замечаний превышает предел ячейки Excel (32767 символов): " + Txt(merged.Values[2]) + ". Источники сохранены; сократите тексты перед повторной сборкой.");
                for (int col = 8; col < 20; col++) {
                    var active = reviews.Where(r => Txt(r.Values[8]) != "" || Txt(r.Values[col]) != "").ToList();
                    var distinct = active.Select(r => col == 9 ? Pretty(col,r.Values[col]) : XlsxReader.Normal(col >=12 && col<18 ? Flag(r.Values[col]) : r.Values[col])).Distinct().ToList();
                    if (distinct.Count < 2) continue;
                    string decision = col == 8 || col == 9 ? "Проверяющий и дата из последней датированной проверки; при равной дате — последний файл в списке." :
                        col == 10 || col == 11 ? "Первое непустое значение в порядке файлов; требуется сверка реквизитов." :
                        col < 18 ? "Все отметки 1 сохранены; различия показаны в окне конфликтов." : "Все разные тексты объединены с указанием проверяющего и даты.";
                    result.Conflicts.Add(new Conflict { Document = Txt(merged.Values[2]), Field = ExcelColumn(col+1) + " — " + Fields[col-8],
                        Details = string.Join("\n",active.Select(r => SourceDescription(r,col))), Decision = decision });
                }
            }
            result.Documents = result.Rows.Count;
            DateTime min = asOf.Date;
            foreach (var row in result.Rows) {
                if (Txt(row.Values[8]) != "") {
                    result.Reviewed++;
                    double? d = row.FirstReviewDate ?? XlsxReader.DateSerial(row.Values[9]);
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
        public static MergeResult Run(IList<string> files,string output,string mainName,DateTime asOf,Action<int,string> progress,CancellationToken token) {
            output=Path.GetFullPath(output);
            if(Path.GetExtension(output).ToLowerInvariant()!=".xlsx") throw new InvalidOperationException("Результат необходимо сохранить с расширением .xlsx.");
            bool exists=File.Exists(output);string before=null;
            if(exists) {
                try {using(var check=new FileStream(output,FileMode.Open,FileAccess.ReadWrite,FileShare.None)){} before=Digest(output);}
                catch(IOException) {throw new IOException("Закройте сводный документ в Excel и повторите сборку: файл занят другой программой.");}
            }
            var sources=new List<string>();if(exists)sources.Add(output);
            foreach(string f in files.Select(Path.GetFullPath))if(!sources.Any(s=>string.Equals(s,f,StringComparison.OrdinalIgnoreCase)))sources.Add(f);
            var result=Collect(sources,mainName,asOf,progress,token);result.Files=files.Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            string parent=Path.GetDirectoryName(output);Directory.CreateDirectory(parent);
            string temp=Path.Combine(parent,"~merge-"+Guid.NewGuid().ToString("N")+".xlsx");
            try {
                Check(token);progress(45,"Создание XLSX с формулами");
                new XlsxWriter().Save(sources[0],temp,mainName,asOf,result,progress,token);
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
