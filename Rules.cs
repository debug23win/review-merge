using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ReviewMerge {
    // Shared interpretation of reviewer names, box cells and volume parts. Excel formulas mirror these rules.
    public static class Rules {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static readonly XNamespace N = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public const int MaxBoxesInCell = 100;
        static string Text(object v) { return XlsxReader.Text(v); }

        // Names: "Яковлева Александра Алексеевна", "Яковлева А.А.", "АА Яковлева" and "А.А.Яковлева" share the surname "Яковлева".
        public static string CleanName(object value) { return Regex.Replace(Text(value), @"\s+", " "); }
        static string[] NameTokens(string name) {
            string s = Regex.Replace(CleanName(name), @"(?<=\p{Lu}\.)(?=\p{Lu}\p{Ll})", " ");
            return s.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim(',', ';')).Where(t => t != "").ToArray();
        }
        static bool IsInitials(string token) { return Regex.IsMatch(token, @"^(?:\p{Lu}\.?){1,3}$") || Regex.IsMatch(token, @"^(?:\p{L}\.){1,3}$"); }
        public static string Surname(string name) {
            var tokens = NameTokens(name);
            if (tokens.Length == 0) return "";
            string s = tokens.FirstOrDefault(t => !IsInitials(t)) ?? tokens[0];
            s = s.TrimEnd('.');
            if (s.Length > 1 && (s == s.ToLowerInvariant() || s == s.ToUpperInvariant())) s = char.ToUpperInvariant(s[0]) + s.Substring(1).ToLowerInvariant();
            return s;
        }
        public static string SurnameKey(string name) { return Surname(name).ToUpperInvariant().Replace('Ё', 'Е'); }
        // Same person written identically apart from case, spaces after initials and Ё.
        public static string NameKey(string name) { return Regex.Replace(CleanName(name), @"\.\s+", ".").ToUpperInvariant().Replace('Ё', 'Е'); }

        // Boxes: 45; "45 и 46"; "45,46"; "45-48" (consecutive boxes 45..48). Words and signs are ignored: "№2", "короб2", "к. №2" are box 2.
        public static List<int> ParseBoxes(object value, out bool valid) {
            var result = new List<int>(); valid = true;
            if (value == null) return result;
            if (value is double || value is int || value is long) {
                double d = Convert.ToDouble(value, Inv);
                if (d > 0 && d == Math.Floor(d) && d < 1000000) result.Add((int)d); else valid = false;
                return result;
            }
            string s = Text(value);
            if (s == "") return result;
            s = Regex.Replace(s, @"(\d)[.,]0+(?!\d)", "$1");
            s = Regex.Replace(s, @"\s*[-–—]\s*", "-");
            s = Regex.Replace(s, @"[^\d-]+", ",");
            s = Regex.Replace(s, @"(?<!\d)-|-(?!\d)", ",");
            foreach (string part in s.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)) {
                var m = Regex.Match(part, @"^(\d{1,6})(?:-(\d{1,6}))?$");
                if (!m.Success) { valid = false; return new List<int>(); }
                int from = int.Parse(m.Groups[1].Value, Inv), to = m.Groups[2].Success ? int.Parse(m.Groups[2].Value, Inv) : from;
                if (from <= 0 || to < from || to - from >= MaxBoxesInCell) { valid = false; return new List<int>(); }
                for (int n = from; n <= to; n++) if (!result.Contains(n)) result.Add(n);
            }
            if (result.Count == 0 || result.Count > MaxBoxesInCell) { valid = false; return new List<int>(); }
            return result;
        }
        public static string Kind(object kind) { string k = Text(kind).Trim().ToUpperInvariant(); return k == "" ? "Не указан" : k; }
        public static List<string> BoxKeys(object kind, object box) {
            bool valid; string k = Kind(kind);
            return ParseBoxes(box, out valid).Select(n => k + "|" + n.ToString(Inv)).ToList();
        }
        // Several boxes of one record are kept as ";ПД|45;ПД|46;" so formulas can look for ";ПД|45;".
        public static string BoxList(IEnumerable<string> keys) { var list = keys.ToList(); return list.Count == 0 ? "" : ";" + string.Join(";", list) + ";"; }
        public static string BoxDisplay(string key) { int bar = key.IndexOf('|'); return bar < 0 ? key : key.Substring(bar + 1) + key.Substring(0, bar).ToLowerInvariant(); }
        public static string BoxListDisplay(string list) { return string.Join(", ", (list ?? "").Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(BoxDisplay)); }
        public static bool PlainBoxNumber(object box) { bool valid; var list = ParseBoxes(box, out valid); return valid && list.Count == 1 && Regex.IsMatch(Text(box), @"^\d+(?:\.0+)?$"); }

        // Volume: fragments and parts ("Фрагмент 2", "_фрагмент1", "_Часть1", "(Часть 6.2)"), changes ("_изм.1") and the ИУЛ ending ("-УЛ", "-ИУЛ") belong to one volume.
        static string FileStem(object file, object ext) {
            string name = XlsxReader.Normal(file), e = XlsxReader.Normal(ext).TrimStart('.');
            if (e != "" && name.EndsWith("." + e, StringComparison.Ordinal)) name = name.Substring(0, name.Length - e.Length - 1);
            return name;
        }
        public static string VolumeKey(object file, object ext) {
            string name = FileStem(file, ext);
            for (int i = 0; i < 2; i++) {
                name = Regex.Replace(name, @"\s*[\(\[ _,\-]*(?<![\p{L}\p{N}])(?:ФРАГМЕНТ|ЧАСТЬ)\s*(?:№|N)?\s*\d+(?:\.\d+)*(?:\s*ИЗ\s*\d+)?[\)\]]*", "");
                name = Regex.Replace(name, @"[\(\[ _,\-]+ИЗМ[.\s]*\d+[\)\]]*$", "");
                name = Regex.Replace(name, @"[-_ ]+[\(\[]?И?УЛ[\)\]]?$", "");
            }
            return Regex.Replace(name, @"\s+", " ").Trim();
        }
        // ИУЛ is a part of a volume, never a volume of its own: "…-УЛ", "…-ИУЛ", "…-УЛ-ГТСС1".
        public static bool IsUl(object file, object ext) {
            return Regex.IsMatch(FileStem(file, ext), @"(?:^|[-_ \(\[])И?УЛ(?:$|[-_ \)\]])");
        }

        // Excel may omit row and cell addresses; number them the way Excel reads them.
        public static void NumberCells(XDocument sheet) {
            var data = sheet == null || sheet.Root == null ? null : sheet.Root.Element(N + "sheetData");
            if (data == null) return;
            int row = 0;
            foreach (var r in data.Elements(N + "row")) {
                int n; if (int.TryParse((string)r.Attribute("r"), out n)) row = n; else { row++; r.SetAttributeValue("r", row); }
                int col = 0;
                foreach (var c in r.Elements(N + "c")) {
                    string address = (string)c.Attribute("r");
                    if (address == null) { col++; c.SetAttributeValue("r", MergeEngine.ExcelColumn(col) + row); continue; }
                    int k = 0; foreach (char ch in address) { if (ch < 'A' || ch > 'Z') break; k = k * 26 + ch - 'A' + 1; }
                    col = k;
                }
            }
        }
    }
}
