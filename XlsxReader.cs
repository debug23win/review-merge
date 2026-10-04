using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace ReviewMerge {
    public sealed class SourceRow {
        public int Row;
        public string File;
        public object[] Values = new object[20];
        public object FirstDateSource;
        public double? PriorFirstDate;
        public bool IsConsolidated;
    }
    public sealed class SourceBook {
        public string Path;
        public int HeaderRow, DataStart;
        public readonly List<SourceRow> Rows = new List<SourceRow>();
        public readonly List<SourceRow> Orphans = new List<SourceRow>();
    }
    public static class XlsxReader {
        static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly XNamespace Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        public static string Text(object value) {
            return value == null ? "" : Convert.ToString(value, Inv).Trim();
        }
        public static string Normal(object value) {
            return Regex.Replace(Text(value).Normalize(NormalizationForm.FormKC), @"\s+", " ").ToUpperInvariant();
        }
        public static double? DateSerial(object value) {
            if (value == null || Text(value) == "") return null;
            double number;
            if (value is double || value is int || value is long) {
                number = Convert.ToDouble(value, Inv);
                if (number >= 36526 && number < 109575) return Math.Floor(number);
                return null;
            }
            DateTime dt;
            var s = Text(value);
            var formats = new [] { "dd.MM.yyyy", "d.M.yyyy", "dd.MM.yy", "yyyy-MM-dd", "dd/MM/yyyy", "dd.MM.yyyy HH:mm", "yyyy-MM-ddTHH:mm:ss" };
            if (DateTime.TryParseExact(s, formats, new CultureInfo("ru-RU"), DateTimeStyles.AllowWhiteSpaces, out dt))
                return dt.Year >= 2000 && dt.Year < 2200 ? (double?)dt.Date.ToOADate() : null;
            if (Regex.IsMatch(s,@"^\d{4}-\d{2}-\d{2}") && DateTime.TryParse(s,Inv,DateTimeStyles.RoundtripKind,out dt))
                return dt.Year >= 2000 && dt.Year < 2200 ? (double?)dt.Date.ToOADate() : null;
            if (double.TryParse(s.Replace(',', '.'), NumberStyles.Float, Inv, out number) && number >= 36526 && number < 109575)
                return Math.Floor(number);
            return null;
        }
        static string TimeKey(object value) {
            double n;
            if (value is double && (n = (double)value) >= 0 && n < 1)
                return DateTime.FromOADate(n).ToString("HH:mm:ss", Inv);
            DateTime dt;
            if (DateTime.TryParseExact(Text(value), new [] { "H:mm", "HH:mm", "H:mm:ss", "HH:mm:ss" }, Inv, DateTimeStyles.None, out dt))
                return dt.ToString("HH:mm:ss", Inv);
            return Normal(value);
        }
        public static string Key(SourceRow row) {
            string filename = Normal(row.Values[2]);
            string extension = Normal(row.Values[3]);
            if (extension != "" && filename.EndsWith("." + extension, StringComparison.Ordinal))
                filename = filename.Substring(0, filename.Length - extension.Length - 1);
            double? date = DateSerial(row.Values[5]);
            return string.Join("\u001f", new [] { filename, extension, Normal(row.Values[7]),
                date.HasValue ? date.Value.ToString(Inv) : Normal(row.Values[5]), TimeKey(row.Values[6]) });
        }
        static XDocument Load(ZipArchive z, string entryName) {
            var entry = z.GetEntry(entryName);
            if (entry == null) throw new InvalidDataException("В Excel отсутствует раздел " + entryName);
            if (entry.Length > 150000000) throw new InvalidDataException("Раздел Excel слишком большой: " + entryName);
            using (var stream = entry.Open())
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 150000000 }))
                return XDocument.Load(reader);
        }
        static int Column(string address) {
            int result = 0;
            foreach (char ch in address) { if (ch < 'A' || ch > 'Z') break; result = result * 26 + ch - 'A' + 1; }
            return result;
        }
        public static SourceBook Read(string path, string sheetName) {
            if (!string.Equals(System.IO.Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Поддерживаются файлы .xlsx: " + System.IO.Path.GetFileName(path));
            var result = new SourceBook { Path = path };
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read)) {
                var workbook = Load(zip, "xl/workbook.xml");
                var sheet = workbook.Descendants(Ns + "sheet").FirstOrDefault(s =>
                    string.Equals(((string)s.Attribute("name") ?? "").Trim(), sheetName.Trim(), StringComparison.OrdinalIgnoreCase));
                if (sheet == null) throw new InvalidDataException(System.IO.Path.GetFileName(path) + ": нет листа «" + sheetName + "».");
                var rels = Load(zip, "xl/_rels/workbook.xml.rels");
                string id = (string)sheet.Attribute(Rel + "id");
                var relationship = rels.Root.Elements().FirstOrDefault(e => (string)e.Attribute("Id") == id);
                if (relationship == null) throw new InvalidDataException("Не удалось найти основной лист Excel.");
                string target = ((string)relationship.Attribute("Target")).Replace('\\', '/');
                target = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
                if (target.Contains("../")) throw new InvalidDataException("Неподдерживаемая ссылка на лист Excel.");
                var strings = new List<string>();
                if (zip.GetEntry("xl/sharedStrings.xml") != null)
                    strings.AddRange(Load(zip, "xl/sharedStrings.xml").Root.Elements(Ns + "si")
                        .Select(e => string.Concat(e.Descendants(Ns + "t").Select(t => t.Value))));
                bool date1904 = (string)workbook.Root.Element(Ns + "workbookPr").NullAttribute("date1904") == "1";
                var data = Load(zip, target);
                var all = new List<SourceRow>();
                foreach (var xmlRow in data.Descendants(Ns + "sheetData").Elements(Ns + "row")) {
                    int rn;
                    if (!int.TryParse((string)xmlRow.Attribute("r"), out rn)) continue;
                    var row = new SourceRow { Row = rn, File = path };
                    foreach (var cell in xmlRow.Elements(Ns + "c")) {
                        int col = Column((string)cell.Attribute("r") ?? "");
                        if (col < 1 || (col > 20 && col != 27)) continue;
                        string type = (string)cell.Attribute("t");
                        string raw = (string)cell.Element(Ns + "v");
                        object value = null;
                        if (type == "s") {
                            int index; if (int.TryParse(raw, out index) && index >= 0 && index < strings.Count) value = strings[index];
                        } else if (type == "inlineStr") value = string.Concat(cell.Descendants(Ns + "t").Select(t => t.Value));
                        else if (raw != null) {
                            double n;
                            value = type != "str" && type != "e" && double.TryParse(raw, NumberStyles.Float, Inv, out n) ? (object)n : raw;
                        }
                        if (date1904 && (col == 6 || col == 10 || col == 27) && value is double) value = (double)value + 1462;
                        if (col == 27) row.FirstDateSource = value;
                        else row.Values[col - 1] = value;
                    }
                    all.Add(row);
                }
                var header = all.FirstOrDefault(r => r.Row <= 30 && Normal(r.Values[8]) == "ПРОВЕРИЛ" && Normal(r.Values[9]).Contains("ДАТА"));
                if (header == null || !Normal(header.Values[2]).Contains("ФАЙЛ") || !Normal(header.Values[7]).Contains("КОНТРОЛЬ"))
                    throw new InvalidDataException(System.IO.Path.GetFileName(path) + ": не совпадает структура A:T листа замечаний.");
                result.HeaderRow = header.Row;
                if (Normal(header.Values[11]) != "КОРОБ" || !Normal(header.Values[18]).Contains("ОПИСИ") || !Normal(header.Values[19]).Contains("ПРИМЕЧАН"))
                    throw new InvalidDataException(System.IO.Path.GetFileName(path) + ": столбцы I:T не соответствуют таблице проверки.");
                bool knownFirstDate = Normal(header.FirstDateSource) == "ПЕРВАЯ ДАТА ИЗ ИСТОЧНИКОВ";
                result.DataStart = header.Row + 1;
                var next = all.FirstOrDefault(r => r.Row == result.DataStart);
                if (next != null && Text(next.Values[0]) == "1" && Text(next.Values[2]) == "3" && Text(next.Values[8]) == "9")
                    result.DataStart++;
                foreach (var row in all.Where(r => r.Row >= result.DataStart)) {
                    row.IsConsolidated = knownFirstDate;
                    if (knownFirstDate) row.PriorFirstDate = DateSerial(row.FirstDateSource);
                    if (Text(row.Values[2]) != "") result.Rows.Add(row);
                    else if (row.Values.Skip(8).Any(v => Text(v) != "")) result.Orphans.Add(row);
                }
                if (result.Rows.Count == 0) throw new InvalidDataException(System.IO.Path.GetFileName(path) + ": нет документов на основном листе.");
            }
            return result;
        }
        static XAttribute NullAttribute(this XElement element, string name) { return element == null ? null : element.Attribute(name); }
    }
}
