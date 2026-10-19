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
    // Values use the original logical A:T schema internally. Only worksheet addresses
    // shift when the optional text field is present after the box column.
    public sealed class RegisterColumns {
        public bool HasVolumeNumber;
        public int Count { get { return HasVolumeNumber ? 21 : 20; } }
        public int Physical(int logical) { return logical >= 13 && HasVolumeNumber ? logical + 1 : logical; }
        public string Letter(int logical) { return MergeEngine.ExcelColumn(Physical(logical)); }
        public string Marks { get { return Letter(13) + ":" + Letter(18); } }
        public string Notes { get { return Letter(19) + ":" + Letter(20); } }
    }
    public sealed class SourceRow {
        public int Row;
        public string File;
        public object[] Values = new object[20];
        // Values as they are in the file when Values were interpreted (a mark "да" read as 1).
        public object[] Written;
        public object FirstDateSource;
        public bool IsConsolidated;
        public string VolumeNumber;
        // «Замечания устранены»: 1, a date or another mark after the notes; empty when the file has no such column.
        public object Resolved;
        public RegisterColumns Columns;
    }
    public sealed class SourceBook {
        public string Path;
        public int HeaderRow, DataStart, ResolvedColumn;
        public RegisterColumns Columns;
        public readonly List<SourceRow> Rows = new List<SourceRow>();
        public readonly List<SourceRow> Orphans = new List<SourceRow>();
        // The row of column numbers (1, 2, 3 …) found among the data, e.g. moved there by a sort in Excel: not a document.
        public readonly List<int> NumberRows = new List<int>();
    }
    // Cell values of a sheet by row number; Values[0] is column A.
    public sealed class RawSheet {
        public string Path;
        public bool Date1904;
        public readonly SortedDictionary<int, object[]> Rows = new SortedDictionary<int, object[]>();
        // Merged ranges as first row, first column, last row, last column; custom row heights.
        public readonly List<int[]> Merges = new List<int[]>();
        public readonly Dictionary<int, double> Heights = new Dictionary<int, double>();
        public object At(int row, int col) { object[] v; return col >= 1 && Rows.TryGetValue(row, out v) && col <= v.Length ? v[col - 1] : null; }
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
            var s = Regex.Replace(Text(value), @"\s*(?:г\.?|года)\s*$", "", RegexOptions.IgnoreCase);
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
        static object RawValue(XElement cell,IList<string> strings) {
            if(cell==null)return null;
            string type=(string)cell.Attribute("t"),raw=(string)cell.Element(Ns+"v");
            if(type=="inlineStr")return string.Concat(cell.Descendants(Ns+"t").Select(t=>t.Value));
            int index;if(type=="s"&&int.TryParse(raw,out index)&&index>=0&&index<strings.Count)return strings[index];
            double number;if(type!="str"&&type!="e"&&double.TryParse(raw,NumberStyles.Float,Inv,out number))return number;
            return raw;
        }
        static int Column(string address) {
            int result = 0;
            foreach (char ch in address) { if (ch < 'A' || ch > 'Z') break; result = result * 26 + ch - 'A' + 1; }
            return result;
        }
        // The worksheet of a sheet name (null when the book has no such sheet), the shared strings and the date system of the book.
        static XDocument SheetXml(ZipArchive zip, string sheetName, out List<string> strings, out bool date1904) {
            var workbook = Load(zip, "xl/workbook.xml");
            strings = new List<string>();
            date1904 = (string)workbook.Root.Element(Ns + "workbookPr").NullAttribute("date1904") == "1";
            var sheet = workbook.Descendants(Ns + "sheet").FirstOrDefault(s =>
                string.Equals(((string)s.Attribute("name") ?? "").Trim(), sheetName.Trim(), StringComparison.OrdinalIgnoreCase));
            if (sheet == null) return null;
            var rels = Load(zip, "xl/_rels/workbook.xml.rels");
            string id = (string)sheet.Attribute(Rel + "id");
            var relationship = rels.Root.Elements().FirstOrDefault(e => (string)e.Attribute("Id") == id);
            if (relationship == null) throw new InvalidDataException("Не удалось найти основной лист Excel.");
            string target = ((string)relationship.Attribute("Target")).Replace('\\', '/');
            target = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
            if (target.Contains("../")) throw new InvalidDataException("Неподдерживаемая ссылка на лист Excel.");
            if (zip.GetEntry("xl/sharedStrings.xml") != null)
                strings.AddRange(Load(zip, "xl/sharedStrings.xml").Root.Elements(Ns + "si")
                    .Select(e => string.Concat(e.Descendants(Ns + "t").Select(t => t.Value))));
            var data = Load(zip, target);
            Rules.NumberCells(data);
            return data;
        }
        // Values of the first width columns of every row of a sheet, as SourceRow reads them; null when there is no such sheet.
        public static RawSheet ReadSheet(string path, string sheetName, int width) {
            if (!string.Equals(System.IO.Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Поддерживаются файлы .xlsx: " + System.IO.Path.GetFileName(path));
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read)) {
                List<string> strings; bool date1904;
                var data = SheetXml(zip, sheetName, out strings, out date1904);
                if (data == null) return null;
                var result = new RawSheet { Path = path, Date1904 = date1904 };
                foreach (var xmlRow in data.Descendants(Ns + "sheetData").Elements(Ns + "row")) {
                    int rn; if (!int.TryParse((string)xmlRow.Attribute("r"), out rn)) continue;
                    var values = new object[width];
                    foreach (var cell in xmlRow.Elements(Ns + "c")) {
                        int col = Column((string)cell.Attribute("r") ?? "");
                        if (col >= 1 && col <= width) values[col - 1] = RawValue(cell, strings);
                    }
                    result.Rows[rn] = values;
                    double height; string custom = (string)xmlRow.Attribute("customHeight"); if ((custom == "1" || custom == "true") && double.TryParse((string)xmlRow.Attribute("ht"), NumberStyles.Float, Inv, out height)) result.Heights[rn] = height;
                }
                foreach (var merge in data.Descendants(Ns + "mergeCell")) {
                    var m = Regex.Match((string)merge.Attribute("ref") ?? "", @"^([A-Z]+)(\d+):([A-Z]+)(\d+)$");
                    if (m.Success) result.Merges.Add(new[] { int.Parse(m.Groups[2].Value, Inv), Column(m.Groups[1].Value), int.Parse(m.Groups[4].Value, Inv), Column(m.Groups[3].Value) });
                }
                return result;
            }
        }
        public static SourceBook Read(string path, string sheetName) {
            if (!string.Equals(System.IO.Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Поддерживаются файлы .xlsx: " + System.IO.Path.GetFileName(path));
            var result = new SourceBook { Path = path };
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read)) {
                List<string> strings; bool date1904;
                var data = SheetXml(zip, sheetName, out strings, out date1904);
                if (data == null) throw new InvalidDataException(System.IO.Path.GetFileName(path) + ": нет листа «" + sheetName + "».");
                var all = new List<SourceRow>();
                foreach (var xmlRow in data.Descendants(Ns + "sheetData").Elements(Ns + "row")) {
                    int rn;
                    if (!int.TryParse((string)xmlRow.Attribute("r"), out rn)) continue;
                    var row = new SourceRow { Row = rn, File = path, Values = new object[Width] };
                    foreach (var cell in xmlRow.Elements(Ns + "c")) {
                        int col = Column((string)cell.Attribute("r") ?? "");
                        if (col < 1 || col > Width) continue;
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
                        if (date1904 && (col == 6 || col == 10 || col == 27 || col == 28) && value is double) value = (double)value + 1462;
                        row.Values[col - 1] = value;
                    }
                    all.Add(row);
                }
                var header = all.FirstOrDefault(r => r.Row <= 30 && Normal(r.Values[8]) == "ПРОВЕРИЛ" && Normal(r.Values[9]).Contains("ДАТА"));
                if (header == null || !Normal(header.Values[2]).Contains("ФАЙЛ") || !Normal(header.Values[7]).Contains("КОНТРОЛЬ"))
                    throw new InvalidDataException(System.IO.Path.GetFileName(path) + ": не совпадает структура основного листа замечаний.");
                result.HeaderRow = header.Row;
                string volumeHeader = Normal(header.Values[12]).Replace(" ", "");
                result.Columns = new RegisterColumns { HasVolumeNumber = volumeHeader == "НОМЕРТОМА" || volumeHeader == "№ТОМА" };
                if (Normal(header.Values[11]) != "КОРОБ" || !Normal(header.Values[result.Columns.Physical(19)-1]).Contains("ОПИСИ") || !Normal(header.Values[result.Columns.Physical(20)-1]).Contains("ПРИМЕЧАН"))
                    throw new InvalidDataException(System.IO.Path.GetFileName(path) + ": после «Короб» допускается «Номер тома», затем шесть отметок, ошибки в описи и примечание.");
                for (int c = result.Columns.Physical(20) + 1; c <= Width; c++) if (Rules.IsResolvedHeader(header.Values[c - 1])) { result.ResolvedColumn = c; break; }
                foreach (var row in all) {
                    var raw = row.Values;
                    row.Columns = result.Columns;
                    row.VolumeNumber = result.Columns.HasVolumeNumber ? Text(raw[12]) : null;
                    row.FirstDateSource = raw[result.Columns.Count + 6];
                    if (result.ResolvedColumn > 0) { row.Resolved = raw[result.ResolvedColumn - 1]; if (date1904 && row.Resolved is double && (double)row.Resolved > 1000) row.Resolved = (double)row.Resolved + 1462; }
                    row.Values = Enumerable.Range(1,20).Select(c => raw[result.Columns.Physical(c)-1]).ToArray();
                }
                bool knownFirstDate = Normal(header.FirstDateSource) == "ПЕРВАЯ ДАТА ИЗ ИСТОЧНИКОВ";
                result.DataStart = header.Row + 1;
                var next = all.FirstOrDefault(r => r.Row == result.DataStart);
                if (next != null && Text(next.Values[0]) == "1" && Text(next.Values[2]) == "3" && Text(next.Values[8]) == "9")
                    result.DataStart++;
                foreach (var row in all.Where(r => r.Row >= result.DataStart)) {
                    row.IsConsolidated = knownFirstDate;
                    if (Enumerable.Range(0, 9).All(i => Text(row.Values[i]) == (i + 1).ToString(Inv))) { result.NumberRows.Add(row.Row); continue; }
                    if (Text(row.Values[2]) != "") result.Rows.Add(row);
                    else if (row.Values.Skip(8).Any(v => Text(v) != "")) result.Orphans.Add(row);
                }
                if (result.Rows.Count == 0) throw new InvalidDataException(System.IO.Path.GetFileName(path) + ": нет документов на основном листе.");
            }
            return result;
        }
        // Columns read from every row: the register, the resolved-remarks column and room for user columns before it.
        const int Width = 60;
        static XAttribute NullAttribute(this XElement element, string name) { return element == null ? null : element.Attribute(name); }
    }
}
