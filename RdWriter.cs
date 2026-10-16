using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;

namespace ReviewMerge {
    // The RD build: merged sets and the formulas of the check sheet, the helper columns of «Ведомость», and the sheets the program owns —
    // «Элементы», «Статистика», «СВОД» and «Итоги проверки» — written again with formulas and the values Excel calculates from them.
    public sealed partial class XlsxWriter {
        const string Gold = "FF996600", Green = "FF00B050", Red = "FFFF0000", Blue = "FF0070C0", Black = "FF000000", Gray = "FFF2F2F2", Head = "FFD9E1F2", Soft = "FF808080", Ink = "FF404040", Hint = "FFC00000";
        const string ReportFont = "Times New Roman", RdElementsSheet = "Элементы", RdStatisticsSheet = "Статистика", RdSummarySheet = "СВОД", Ved = "Ведомость";
        readonly Dictionary<string, int> rdParts = new Dictionary<string, int>(StringComparer.Ordinal);
        RdBook rd;

        // A font, fill, border, cell format or differential format of styles.xml, added once.
        int RdPart(string section, string key, Func<XElement> make) {
            string k = section + "|" + key; int id; if (rdParts.TryGetValue(k, out id)) return id;
            var list = styles.Root.Element(N + section);
            if (list == null) { list = new XElement(N + section); (styles.Root.Element(N + "cellStyles") ?? styles.Root.Element(N + "cellXfs")).AddAfterSelf(list); }
            id = list.Elements().Count(); list.Add(make()); list.SetAttributeValue("count", id + 1); rdParts[k] = id; return id;
        }
        int RdNumberFormat(string code) {
            switch (code) { case null: case "General": return 0; case "0": return 1; case "0%": return 9; case "h:mm": return 20; case "@": return 49; }
            string k = "numFmt|" + code; int id; if (rdParts.TryGetValue(k, out id)) return id;
            var formats = styles.Root.Element(N + "numFmts");
            var same = formats.Elements(N + "numFmt").FirstOrDefault(e => (string)e.Attribute("formatCode") == code);
            if (same != null) id = (int)same.Attribute("numFmtId");
            else {
                id = Math.Max(164, formats.Elements().Select(e => (int)e.Attribute("numFmtId")).DefaultIfEmpty(163).Max() + 1);
                formats.Add(new XElement(N + "numFmt", new XAttribute("numFmtId", id), new XAttribute("formatCode", code))); formats.SetAttributeValue("count", formats.Elements().Count());
            }
            rdParts[k] = id; return id;
        }
        static XElement Edge(string name, string style, string color) { return style == null ? new XElement(N + name) : new XElement(N + name, new XAttribute("style", style), new XElement(N + "color", new XAttribute("rgb", color))); }
        // A cell format of the RD sheets: Times New Roman, centred and wrapped unless told otherwise; borders left, right, top, bottom.
        int Rs(double size = 12, bool bold = false, bool italic = false, string color = Black, string fill = null, string fmt = null, string h = "center", string v = "center", bool wrap = true,
               string l = null, string r = null, string t = null, string b = null, string edge = Black) {
            int font = RdPart("fonts", size + "|" + bold + "|" + italic + "|" + color, () => new XElement(N + "font", bold ? new XElement(N + "b") : null, italic ? new XElement(N + "i") : null,
                new XElement(N + "sz", new XAttribute("val", size.ToString(Inv))), new XElement(N + "color", new XAttribute("rgb", color)), new XElement(N + "name", new XAttribute("val", ReportFont)),
                new XElement(N + "family", new XAttribute("val", 1)), new XElement(N + "charset", new XAttribute("val", 204))));
            int fillId = fill == null ? 0 : RdPart("fills", fill, () => new XElement(N + "fill", new XElement(N + "patternFill", new XAttribute("patternType", "solid"), new XElement(N + "fgColor", new XAttribute("rgb", fill)), new XElement(N + "bgColor", new XAttribute("indexed", 64)))));
            int border = l == null && r == null && t == null && b == null ? 0 : RdPart("borders", l + "|" + r + "|" + t + "|" + b + "|" + edge, () => new XElement(N + "border", Edge("left", l, edge), Edge("right", r, edge), Edge("top", t, edge), Edge("bottom", b, edge), new XElement(N + "diagonal")));
            int number = RdNumberFormat(fmt);
            return RdPart("cellXfs", font + "|" + fillId + "|" + border + "|" + number + "|" + h + "|" + v + "|" + wrap, () => new XElement(N + "xf", new XAttribute("numFmtId", number), new XAttribute("fontId", font),
                new XAttribute("fillId", fillId), new XAttribute("borderId", border), new XAttribute("xfId", 0), new XAttribute("applyNumberFormat", 1), new XAttribute("applyFont", 1), new XAttribute("applyFill", 1),
                new XAttribute("applyBorder", 1), new XAttribute("applyAlignment", 1), new XElement(N + "alignment", new XAttribute("horizontal", h), new XAttribute("vertical", v), new XAttribute("wrapText", wrap ? 1 : 0))));
        }
        int Thin(double size = 11, bool bold = false, string color = Black, string fill = null, string fmt = null, string h = "center", bool wrap = true) {
            return Rs(size, bold, false, color, fill, fmt, h, "center", wrap, "thin", "thin", "thin", "thin", Soft);
        }
        // Differential format of conditional formatting: font colour and weight, fill.
        int RdDxf(string color, bool bold, string fill) {
            return RdPart("dxfs", color + "|" + bold + "|" + fill, () => new XElement(N + "dxf", new XElement(N + "font", new XElement(N + "b", new XAttribute("val", bold ? 1 : 0)), color == null ? null : new XElement(N + "color", new XAttribute("rgb", color))),
                fill == null ? null : new XElement(N + "fill", new XElement(N + "patternFill", new XAttribute("patternType", "solid"), new XElement(N + "fgColor", new XAttribute("rgb", fill)), new XElement(N + "bgColor", new XAttribute("rgb", fill))))));
        }

        // A sheet built from scratch: cells, row heights and outline, merges, columns, conditional formatting and page setup.
        sealed class RdSheet {
            public readonly SortedDictionary<int, SortedDictionary<int, XElement>> Rows = new SortedDictionary<int, SortedDictionary<int, XElement>>();
            public readonly Dictionary<int, double> Heights = new Dictionary<int, double>();
            public readonly HashSet<int> Detail = new HashSet<int>(), Hidden = new HashSet<int>(), Collapsed = new HashSet<int>();
            public readonly List<string> Merges = new List<string>();
            public readonly List<XElement> Formats = new List<XElement>(), Columns = new List<XElement>();
            public string Freeze; public bool Grid = true, Portrait; int priority;
            public void Set(int r, int c, object v, int s, string f = null) {
                SortedDictionary<int, XElement> row; if (!Rows.TryGetValue(r, out row)) Rows[r] = row = new SortedDictionary<int, XElement>();
                row[c] = Cell(r, c, v, s, f);
            }
            public void Merge(int r1, int c1, int r2, int c2) { Merges.Add(MergeEngine.ExcelColumn(c1) + r1 + ":" + MergeEngine.ExcelColumn(c2) + r2); }
            public void Width(int from, int to, double width, bool hidden = false) { Columns.Add(Col(from, to, width, hidden)); }
            public void Rule(string range, string type, string formula, int dxf, string op = null) {
                Formats.Add(new XElement(N + "conditionalFormatting", new XAttribute("sqref", range), new XElement(N + "cfRule", new XAttribute("type", type), new XAttribute("dxfId", dxf),
                    new XAttribute("priority", ++priority), op == null ? null : new XAttribute("operator", op), type == "expression" ? new XAttribute("stopIfTrue", 1) : null, new XElement(N + "formula", formula))));
            }
        }
        static XElement FreezePane(string cell) {
            int col = Column(cell), row = int.Parse(Regex.Match(cell, @"\d+$").Value, Inv);
            return new XElement(N + "pane", col > 1 ? new XAttribute("xSplit", col - 1) : null, row > 1 ? new XAttribute("ySplit", row - 1) : null, new XAttribute("topLeftCell", cell),
                new XAttribute("activePane", col > 1 && row > 1 ? "bottomRight" : row > 1 ? "bottomLeft" : "topRight"), new XAttribute("state", "frozen"));
        }
        static XDocument RdXml(RdSheet sh) {
            int lastRow = sh.Rows.Keys.DefaultIfEmpty(1).Max(), lastCol = sh.Rows.Values.SelectMany(r => r.Keys).DefaultIfEmpty(1).Max();
            var data = new XElement(N + "sheetData");
            foreach (int rn in sh.Rows.Keys.Union(sh.Heights.Keys).Union(sh.Detail).Distinct().OrderBy(x => x)) {
                var row = new XElement(N + "row", new XAttribute("r", rn)); double ht;
                if (sh.Heights.TryGetValue(rn, out ht)) { row.SetAttributeValue("ht", ht.ToString(Inv)); row.SetAttributeValue("customHeight", 1); }
                if (sh.Hidden.Contains(rn)) row.SetAttributeValue("hidden", 1);
                if (sh.Detail.Contains(rn)) row.SetAttributeValue("outlineLevel", 1);
                if (sh.Collapsed.Contains(rn)) row.SetAttributeValue("collapsed", 1);
                SortedDictionary<int, XElement> cells; if (sh.Rows.TryGetValue(rn, out cells)) row.Add(cells.Values);
                data.Add(row);
            }
            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XElement(N + "worksheet", new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
                new XElement(N + "sheetPr", new XElement(N + "outlinePr", new XAttribute("summaryBelow", 0)), new XElement(N + "pageSetUpPr", new XAttribute("fitToPage", 1))),
                new XElement(N + "dimension", new XAttribute("ref", "A1:" + MergeEngine.ExcelColumn(lastCol) + lastRow)),
                new XElement(N + "sheetViews", new XElement(N + "sheetView", sh.Grid ? null : new XAttribute("showGridLines", 0), new XAttribute("zoomScale", 90), new XAttribute("zoomScaleNormal", 90), new XAttribute("workbookViewId", 0), sh.Freeze == null ? null : FreezePane(sh.Freeze))),
                new XElement(N + "sheetFormatPr", new XAttribute("defaultRowHeight", 15), sh.Detail.Count > 0 ? new XAttribute("outlineLevelRow", 1) : null),
                sh.Columns.Count == 0 ? null : new XElement(N + "cols", sh.Columns.OrderBy(c => (int)c.Attribute("min"))),
                data,
                sh.Merges.Count == 0 ? null : new XElement(N + "mergeCells", new XAttribute("count", sh.Merges.Count), sh.Merges.Select(m => new XElement(N + "mergeCell", new XAttribute("ref", m)))),
                sh.Formats,
                new XElement(N + "pageMargins", new XAttribute("left", 0.4), new XAttribute("right", 0.4), new XAttribute("top", 0.5), new XAttribute("bottom", 0.5), new XAttribute("header", 0.3), new XAttribute("footer", 0.3)),
                new XElement(N + "pageSetup", new XAttribute("paperSize", 9), new XAttribute("orientation", sh.Portrait ? "portrait" : "landscape"), new XAttribute("fitToWidth", 1), new XAttribute("fitToHeight", 0))));
        }
        XElement DefinedNames() {
            var names = book.Root.Element(N + "definedNames");
            if (names == null) { names = new XElement(N + "definedNames"); var calc = book.Root.Element(N + "calcPr"); if (calc != null) calc.AddBeforeSelf(names); else book.Root.Add(names); }
            return names;
        }
        int SheetIndex(string name) { return book.Root.Element(N + "sheets").Elements().TakeWhile(e => !string.Equals(((string)e.Attribute("name") ?? "").Trim(), name, StringComparison.OrdinalIgnoreCase)).Count(); }
        // A regenerated sheet replaces the old one in place (its relationships go: the new sheet has no drawings or comments); the print names follow.
        void RdPut(string name, XDocument doc, string printArea) {
            string path = PathFor(name);
            if (path != null) entries.Remove(Path.GetDirectoryName(path).Replace('\\', '/') + "/_rels/" + Path.GetFileName(path) + ".rels");
            PutSheet(name, doc);
            int index = SheetIndex(name); var names = DefinedNames();
            names.Elements().Where(e => (int?)e.Attribute("localSheetId") == index && ((string)e.Attribute("name") ?? "").StartsWith("_xlnm.", StringComparison.Ordinal)).Remove();
            if (printArea != null) names.Add(new XElement(N + "definedName", new XAttribute("name", "_xlnm.Print_Area"), new XAttribute("localSheetId", index), printArea));
        }
        static string Sheet(string name) { return "'" + name.Replace("'", "''") + "'!"; }
        string M(string key) { var l = rd.Layout; return Sheet(rd.MainName) + "$" + l.L(key) + "$" + l.First + ":$" + l.L(key) + "$" + rd.Last; }
        string E(string col) { return Sheet(RdElementsSheet) + "$" + col + "$6:$" + col + "$" + rd.ElementsLast; }
        string A(int col) { return Sheet(rd.Acts.Name) + "$" + MergeEngine.ExcelColumn(col) + "$" + rd.Acts.First + ":$" + MergeEngine.ExcelColumn(col) + "$" + rd.Acts.Last; }
        static string Q(string text) { return "\"" + text.Replace("\"", "\"\"") + "\""; }
        static bool Is(string a, string b) { return string.Equals(a ?? "", b ?? "", StringComparison.OrdinalIgnoreCase); }
        string Stage() {
            string stage = Text(rd.Target.At(2, 1));
            return stage;
        }
        string StageLabel() { var m = Regex.Match(Stage(), @"^\s*(\d+)\s*этап", RegexOptions.IgnoreCase); return m.Success ? "Этап " + m.Groups[1].Value : "Этап"; }

        public void SaveRd(string firstFile, string output, RdBook data, DateTime asOf, Action<int, string> progress, CancellationToken token) {
            rd = data;
            using (var s = File.OpenRead(firstFile)) using (var z = new ZipArchive(s, ZipArchiveMode.Read)) foreach (var e in z.Entries) { using (var m = new MemoryStream()) using (var es = e.Open()) { es.CopyTo(m); entries[e.FullName] = m.ToArray(); } }
            book = Xml("xl/workbook.xml"); rels = Xml("xl/_rels/workbook.xml.rels"); types = Xml("[Content_Types].xml");
            var pr = book.Root.Element(N + "workbookPr"); if (pr != null && ((string)pr.Attribute("date1904") == "1" || (string)pr.Attribute("date1904") == "true")) throw new InvalidDataException("Сводный документ использует систему дат 1904. Нужна книга с системой дат 1900.");
            SetupStyles();
            WriteRdMain(token); progress(60, "Комплекты объединены, формулы листа «" + rd.MainName + "» до строки " + rd.Last);
            WriteVedHelpers(token);
            rd.ElementsLast = 5 + Math.Max(60, rd.Elements.Count + 20);
            WriteRdElements(); WriteRdStatistics(); WriteRdSummary(asOf); WriteRdTotals(asOf);
            GrowMainReferences();
            progress(85, "Пересозданы листы: СВОД, Итоги проверки, Статистика, Элементы.");
            FinishRd(output, progress, token);
        }

        // «Проверка комплектов»: the merged columns of every set, new sets after the data, the formulas of the grey columns down to Last.
        void WriteRdMain(CancellationToken token) {
            var layout = rd.Layout; mainPath = PathFor(rd.MainName); var main = Xml(mainPath); ExpandSharedFormulas(main);
            var data = main.Root.Element(N + "sheetData"); var rows = data.Elements(N + "row").ToDictionary(e => (int)e.Attribute("r"));
            int statusCol = layout["status"];
            rd.OldLast = rows.Where(p => p.Key >= layout.First && p.Value.Elements(N + "c").Any(c => Column((string)c.Attribute("r")) == statusCol && c.Element(N + "f") != null)).Select(p => p.Key).DefaultIfEmpty(layout.First).Max();
            rd.Last = Math.Max(rd.Last, rd.OldLast);
            var template = new Dictionary<int, int>();
            foreach (var c in rows.Where(p => p.Key >= layout.First).OrderBy(p => p.Key).SelectMany(p => p.Value.Elements(N + "c"))) { int col = Column((string)c.Attribute("r")); if (!template.ContainsKey(col)) template[col] = (int?)c.Attribute("s") ?? 0; }
            Func<int, XElement> rowAt = rn => {
                XElement row; if (rows.TryGetValue(rn, out row)) return row;
                row = new XElement(N + "row", new XAttribute("r", rn)); var next = rows.Where(p => p.Key > rn).OrderBy(p => p.Key).Select(p => p.Value).FirstOrDefault();
                if (next != null) next.AddBeforeSelf(row); else data.Add(row);
                rows[rn] = row; return row;
            };
            Action<XElement, int, int, object, string, bool> put = (row, rn, col, value, formula, date) => {
                var old = row.Elements(N + "c").FirstOrDefault(e => Column((string)e.Attribute("r")) == col);
                int style = old != null ? (int?)old.Attribute("s") ?? 0 : template.ContainsKey(col) ? template[col] : bodyStyle;
                if (date) style = MainDateStyle(style);
                SetCell(row, Cell(rn, col, value, style, formula));
            };
            foreach (var set in rd.Sets) {
                token.ThrowIfCancellationRequested();
                int rn = set.SheetRow; var row = rowAt(rn);
                var keys = (set.IsNew ? RdLayout.Identity : new string[0]).Concat(RdLayout.Review).Concat(layout.Flags);
                bool notes = false;
                foreach (string k in keys) {
                    int col = layout[k]; if (col == 0) continue;
                    object value = set[k], before = rd.Target.At(rn, col); bool date = false;
                    if (k == "rev" && value is string) { double n; if (double.TryParse(((string)value).Replace(',', '.'), NumberStyles.Float, Inv, out n)) value = n; }
                    if (k == "date" || k == "fixed") {
                        double? d = XlsxReader.DateSerial(value);
                        if (d.HasValue) { value = d.Value; date = true; }
                        else if (value is string) { double n; if (double.TryParse((string)value, NumberStyles.Float, Inv, out n)) value = n; }
                    }
                    if (!set.IsNew && Text(before) == Text(value) && (!date || before is double)) continue;
                    if (Text(value) == "" && Text(before) == "") continue;
                    if ((k == "text" || k == "note") && Text(value).Contains("\n")) notes = true;
                    put(row, rn, col, Text(value) == "" ? null : value, null, date);
                }
                if (notes) { int lines = Math.Max(Text(set["text"]).Split('\n').Length, Text(set["note"]).Split('\n').Length); row.SetAttributeValue("ht", Math.Min(120, Math.Max(30, lines * 15))); row.SetAttributeValue("customHeight", 1); }
            }
            var byRow = rd.Sets.ToDictionary(s => s.SheetRow);
            for (int r = layout.First; r <= rd.Last; r++) {
                if (r % 200 == 0) token.ThrowIfCancellationRequested();
                RdSet set; byRow.TryGetValue(r, out set); var row = rowAt(r); var f = RdFormulas(r);
                string elkey; rd.RowElKey.TryGetValue(r, out elkey); int order; bool first = rd.RowOrder.TryGetValue(r, out order);
                put(row, r, layout["ve"], set == null ? "" : set.Ve, f["ve"], false);
                put(row, r, layout["te"], set == null ? "" : set.Te, f["te"], false);
                put(row, r, layout["tedate"], set == null ? "" : set.TeDate, f["tedate"], false);
                put(row, r, layout["teletter"], set == null ? "" : set.TeLetter, f["teletter"], false);
                put(row, r, layout["cnt"], set == null ? "" : (object)(double)set.Cnt, f["cnt"], false);
                put(row, r, layout["status"], set == null ? "" : set.Status, f["status"], false);
                put(row, r, layout["latest"], set == null ? "" : (object)(set.Latest ? 1.0 : 0.0), f["latest"], false);
                put(row, r, layout["order"], first ? (object)(double)order : "", f["order"], false);
                put(row, r, layout["elkey"], elkey ?? "", f["elkey"], false);
            }
            foreach (var row in data.Elements(N + "row")) { var sorted = row.Elements(N + "c").OrderBy(e => Column((string)e.Attribute("r"))).ToList(); sorted.Remove(); row.Add(sorted); }
            // Validation lists, conditional formatting and the filter of the form reach the new last row.
            if (rd.Last > rd.OldLast) {
                string old = rd.OldLast.ToString(Inv), now = rd.Last.ToString(Inv);
                Func<string, string> grow = refs => Regex.Replace(refs ?? "", @"(:\$?[A-Z]{1,3}\$?)" + old + @"(?!\d)", "${1}" + now);
                foreach (var e in main.Root.Descendants().Where(e => e.Name == N + "dataValidation" || e.Name == N + "conditionalFormatting")) e.SetAttributeValue("sqref", grow((string)e.Attribute("sqref")));
                foreach (var e in main.Root.Descendants(N + "conditionalFormatting").Descendants(N + "formula")) e.Value = grow(e.Value);
                var filter = main.Root.Element(N + "autoFilter"); if (filter != null) filter.SetAttributeValue("ref", grow((string)filter.Attribute("ref")));
            }
            int maxCol = data.Elements(N + "row").Elements(N + "c").Select(e => Column((string)e.Attribute("r"))).DefaultIfEmpty(1).Max(), maxRow = rows.Keys.DefaultIfEmpty(1).Max();
            var dimension = main.Root.Element(N + "dimension"); if (dimension != null) dimension.SetAttributeValue("ref", "A1:" + MergeEngine.ExcelColumn(maxCol) + maxRow);
            Store(mainPath, main);
        }
        // The formulas of the grey columns of the form (the same as in the form the program was made for).
        Dictionary<string, string> RdFormulas(int r) {
            var l = rd.Layout;
            Func<string, string> c = k => "$" + l.L(k) + r;
            Func<string, string> range = k => "$" + l.L(k) + "$" + l.First + ":$" + l.L(k) + "$" + rd.Last;
            string v = Sheet(Ved), code = "TRIM(" + c("rd") + "&\"\")", found = "MATCH(" + code + "," + v + "$AG:$AG,0)", empty = "IF(" + c("rd") + "=\"\",\"\",";
            string agreed = "OR(" + string.Join(",", RdBook.Agreed.Select(a => c("te") + "=" + Q(a))) + ")";
            string above = r > l.First ? "$" + l.L("elkey") + (r - 1) : "\"\"";
            string flags = "$" + MergeEngine.ExcelColumn(l.Flags.Min(f => l[f])) + r + ":$" + MergeEngine.ExcelColumn(l.Flags.Max(f => l[f])) + r;
            string fx = "TRIM(" + c("fixed") + "&\"\")", unresolved = "OR(" + fx + "=\"\"," + fx + "=\"0\",UPPER(" + fx + ")=\"НЕТ\")";
            return new Dictionary<string, string> {
                { "ve", empty + "IF(COUNTIF(" + v + "$AG:$AG," + code + ")=0,\"нет в ВЭ РД\",IF(INDEX(" + v + "$AF:$AF," + found + ")<>" + c("elkey") + ",\"элемент \"&INDEX(" + v + "$AF:$AF," + found + "),\"да\")))" },
                { "te", empty + "IFERROR(INDEX(" + v + "$I:$I," + found + "),\"—\"))" },
                { "tedate", empty + "IFERROR(IF(INDEX(" + v + "$K:$K," + found + ")=\"\",\"\",INDEX(" + v + "$K:$K," + found + ")),\"\"))" },
                { "teletter", empty + "IFERROR(INDEX(" + v + "$L:$L," + found + ")&\"\",\"\"))" },
                { "cnt", empty + "COUNTIF(" + flags + ",1)+IF(TRIM(" + c("note") + "&\"\")<>\"\",1,0))" },
                { "status", empty + "IF(" + c("ve") + "<>\"да\",\"не соответствует ВЭ РД\",IF(" + c("who") + "=\"\",\"не проверен\",IF(AND(" + c("cnt") + ">0," + unresolved + "),\"замечания\",IF(NOT(" + agreed + "),\"не согласован ТЭ\",\"готово к ВПР\")))))" },
                { "latest", empty + "IF(COUNTIFS(" + range("rd") + "," + c("rd") + "," + range("rev") + ",\">\"&N(" + c("rev") + "))=0,1,0))" },
                { "order", "IF(" + c("elkey") + "=\"\",\"\",IF(COUNTIF($" + l.L("elkey") + "$" + l.First + ":" + c("elkey") + "," + c("elkey") + ")=1,MAX($" + l.L("order") + "$" + (l.First - 1) + ":$" + l.L("order") + (r - 1) + ")+1,\"\"))" },
                { "elkey", "IF(TRIM(" + c("el") + "&\"\")<>\"\",TRIM(" + c("el") + "&\"\"),IF(" + c("rd") + "=\"\",\"\"," + above + "))" } };
        }
        // «Ведомость»: the element of every row carried down (AF), the trimmed code (AG) and the first row of an element (AH), 200 rows past the data.
        void WriteVedHelpers(CancellationToken token) {
            var ved = rd.Ved; string path = PathFor(Ved); var doc = Xml(path); ExpandSharedFormulas(doc);
            var data = doc.Root.Element(N + "sheetData"); var rows = data.Elements(N + "row").ToDictionary(e => (int)e.Attribute("r"));
            int oldLast = rows.Where(p => p.Value.Elements(N + "c").Any(c => Column((string)c.Attribute("r")) == 33 && c.Element(N + "f") != null)).Select(p => p.Key).DefaultIfEmpty(0).Max();
            ved.HelperLast = Math.Max(oldLast, ved.Last + 200);
            int style = bodyStyle; var sample = At(doc, ved.First, 33); if (sample != null) style = (int?)sample.Attribute("s") ?? bodyStyle;
            int header = ved.First - (Text(ved.At(ved.First - 1, 4)) == "4" ? 2 : 1);
            XElement headRow; if (rows.TryGetValue(header, out headRow)) {
                string[] names = { "Элемент строки", "Шифр РД без пробелов", "Первая строка элемента" };
                for (int k = 0; k < 3; k++) if (Text(Value(At(doc, header, 32 + k))) == "") SetCell(headRow, Cell(header, 32 + k, names[k], StyleAt(doc, header, 4, headStyle)));
            }
            for (int r = ved.First; r <= ved.HelperLast; r++) {
                if (r % 500 == 0) token.ThrowIfCancellationRequested();
                XElement row;
                if (!rows.TryGetValue(r, out row)) { row = new XElement(N + "row", new XAttribute("r", r)); var next = rows.Where(p => p.Key > r).OrderBy(p => p.Key).Select(p => p.Value).FirstOrDefault(); if (next != null) next.AddBeforeSelf(row); else data.Add(row); rows[r] = row; }
                string element; if (!ved.Element.TryGetValue(r, out element)) element = "";
                string previous; bool start = element != "" && (r == ved.First || !ved.Element.TryGetValue(r - 1, out previous) || !RdBook.Same(previous, element));
                string above = r > ved.First ? "$AF" + (r - 1) : "\"\"";
                SetCell(row, Cell(r, 32, element, style, "IF(TRIM($A" + r + "&\"\")<>\"\",TRIM($A" + r + "&\"\"),IF(TRIM($D" + r + "&\"\")=\"\",\"\"," + above + "))"));
                SetCell(row, Cell(r, 33, Text(ved.At(r, 4)), style, "TRIM($D" + r + "&\"\")"));
                SetCell(row, Cell(r, 34, start ? 1.0 : 0.0, style, r == ved.First ? "IF($AF" + r + "<>\"\",1,0)" : "IF(AND($AF" + r + "<>\"\",$AF" + r + "<>$AF" + (r - 1) + "),1,0)"));
                var sorted = row.Elements(N + "c").OrderBy(e => Column((string)e.Attribute("r"))).ToList(); sorted.Remove(); row.Add(sorted);
            }
            var cols = doc.Root.Element(N + "cols"); if (cols == null) { cols = new XElement(N + "cols"); data.AddBeforeSelf(cols); }
            var helpers = cols.Elements(N + "col").Where(c => (int)c.Attribute("max") >= 32 && (int)c.Attribute("min") <= 34).ToList();
            if (helpers.Count == 0) { cols.Add(Col(32, 34, 12, true)); var sortedCols = cols.Elements().OrderBy(c => (int)c.Attribute("min")).ToList(); sortedCols.Remove(); cols.Add(sortedCols); }
            else foreach (var c in helpers) if ((int)c.Attribute("min") >= 32 && (int)c.Attribute("max") <= 34) c.SetAttributeValue("hidden", 1);
            var dimension = doc.Root.Element(N + "dimension");
            if (dimension != null) { int maxRow = rows.Keys.DefaultIfEmpty(1).Max(), maxCol = Math.Max(34, data.Elements(N + "row").Elements(N + "c").Select(e => Column((string)e.Attribute("r"))).DefaultIfEmpty(1).Max()); dimension.SetAttributeValue("ref", "A1:" + MergeEngine.ExcelColumn(maxCol) + maxRow); }
            Store(path, doc);
        }
        static readonly string[][] ElementColors = { new[] { "готов к ВПР", "FFC6EFCE", "FF006100" }, new[] { "замечания", "FFFFC7CE", "FF9C0006" }, new[] { "не согласован ТЭ", "FFFFEB9C", "FF9C5700" },
            new[] { "неполный комплект", "FFFCE4D6", "FF833C0B" }, new[] { "расхождение с ВЭ РД", "FFFF9999", "FF7F0000" }, new[] { "в проверке", "FFF2F2F2", "FF808080" } };
        void Titles(RdSheet sh, string title, string hint) {
            sh.Set(1, 1, title, Rs(14, true, h: "left", wrap: false)); sh.Heights[1] = 22;
            sh.Set(2, 1, Stage(), Rs(11, italic: true, h: "left", wrap: false));
            if (hint != null) sh.Set(3, 1, hint, Rs(9, color: Hint, h: "left", wrap: false));
        }
        // «Элементы»: the elements of the check sheet in their order with the counts of their current sets, the status, the act and the day the check finished.
        void WriteRdElements() {
            var sh = new RdSheet { Freeze = "C6" };
            Titles(sh, "Элементы на проверке — рассчитывается автоматически", "Лист заполнять не нужно: программа пересоздаёт его при каждой сборке; элементы берутся из листа «" + rd.MainName + "», комплекты считаются по действующим редакциям.");
            string[] heads = { "Номер элемента", "Наименование элемента", "№ элемента по ДГП", "Комплектов в ВЭ РД", "Представлено", "Не представлено", "Не из ВЭ РД", "Проверено", "С замечаниями", "Не согласовано ТЭ",
                "Готово к ВПР", "Статус элемента", "Проверка акта", "Принято по акту (ВЭ)", "Выдано «В производство работ» (ВЭ)", "Дата проверки элемента", "Проверил элемент", "Комплектов с замечаниями ТЭ" };
            double[] widths = { 10, 46, 14, 10, 11, 11, 10, 10, 11, 11, 10, 18, 13, 11, 13, 12, 16, 12 };
            int head = Thin(11, true, fill: Head);
            for (int i = 0; i < heads.Length; i++) { sh.Set(5, i + 1, heads[i], head); sh.Width(i + 1, i + 1, widths[i]); }
            sh.Heights[5] = 60;
            int formula = Thin(11, color: Ink, fill: Gray), name = Thin(11, color: Ink, fill: Gray, h: "left"), day = Thin(11, color: Ink, fill: Gray, fmt: "dd.mm.yyyy");
            var acts = rd.Acts; string v = Sheet(Ved);
            for (int r = 6; r <= rd.ElementsLast; r++) {
                var e = r - 6 < rd.Elements.Count ? rd.Elements[r - 6] : null;
                string a = "$A" + r, blank = "IF(" + a + "=\"\",\"\",", live = M("elkey") + "," + a + "," + M("latest") + ",1";
                Func<string, string> look = col => "IF(" + a + "=\"\",\"\",IFERROR(INDEX(" + v + "$" + col + ":$" + col + ",MATCH(" + a + "," + v + "$AF:$AF,0)),\"\"))";
                string act = acts.Exists && acts.Status > 0 ? blank + "IFERROR(INDEX(" + A(acts.Status) + ",MATCH(" + a + "," + A(acts.El) + ",0)),\"не поступал\"))" : blank + "\"не поступал\")";
                var cells = new object[][] {
                    new object[] { e == null ? "" : e.Key, "IFERROR(INDEX(" + M("elkey") + ",MATCH(ROW()-5," + M("order") + ",0)),\"\")" },
                    new object[] { e == null ? "" : e.Name, "IF(" + a + "=\"\",\"\",IFERROR(INDEX(" + v + "$B:$B,MATCH(" + a + "," + v + "$AF:$AF,0)),INDEX(" + M("elname") + ",MATCH(" + a + "," + M("elkey") + ",0))))" },
                    new object[] { e == null ? "" : e.Dgp, look("G") },
                    new object[] { e == null ? "" : (object)(double)e.InVe, blank + "COUNTIF(" + v + "$AF:$AF," + a + "))" },
                    new object[] { e == null ? "" : (object)(double)e.Presented, blank + "COUNTIFS(" + live + "," + M("ve") + ",\"да\"))" },
                    new object[] { e == null ? "" : (object)(double)e.Missing, blank + "MAX(0,D" + r + "-E" + r + "))" },
                    new object[] { e == null ? "" : (object)(double)e.NotInVe, blank + "COUNTIFS(" + live + "," + M("ve") + ",\"<>да\"))" },
                    new object[] { e == null ? "" : (object)(double)e.Checked, blank + "COUNTIFS(" + live + "," + M("ve") + ",\"да\"," + M("status") + ",\"<>не проверен\"," + M("who") + ",\"<>" + RdBook.Expert + "\"))" },
                    new object[] { e == null ? "" : (object)(double)e.Remarks, blank + "COUNTIFS(" + live + "," + M("status") + ",\"замечания\"))" },
                    new object[] { e == null ? "" : (object)(double)e.NotAgreed, blank + "COUNTIFS(" + live + "," + M("status") + ",\"не согласован ТЭ\"))" },
                    new object[] { e == null ? "" : (object)(double)e.Ready, blank + "COUNTIFS(" + live + "," + M("status") + ",\"готово к ВПР\"))" },
                    new object[] { e == null ? "" : e.Status, blank + "IF(AND(F" + r + "=0,G" + r + "=0,K" + r + "=D" + r + "),\"готов к ВПР\",IF(F" + r + ">0,\"неполный комплект\",IF(G" + r + ">0,\"расхождение с ВЭ РД\",IF(I" + r + ">0,\"замечания\",IF(J" + r + ">0,\"не согласован ТЭ\",\"в проверке\"))))))" },
                    new object[] { e == null ? "" : e.Act, act },
                    new object[] { e == null ? "" : e.ActVe, look("N") },
                    new object[] { e == null ? "" : e.Vpr, look("R") },
                    // The day of the last current set when every current set has a date and none is a remark of ТЭ; who checked that set.
                    new object[] { e == null || !e.Date.HasValue ? "" : (object)e.Date.Value, blank + "IF(AND(COUNTIFS(" + live + ")=COUNTIFS(" + live + "," + M("date") + ",\">0\"),COUNTIFS(" + live + "," + M("who") + ",\"" + RdBook.Expert + "\")=0),SUMPRODUCT(MAX((" + M("elkey") + "=" + a + ")*(" + M("latest") + "=1)*" + M("date") + ")),\"\"))" },
                    new object[] { e == null ? "" : e.Who, "IF(OR(" + a + "=\"\",N(P" + r + ")=0),\"\",IFERROR(LOOKUP(2,1/((" + M("elkey") + "=" + a + ")*(" + M("latest") + "=1)*(" + M("date") + "=P" + r + "))," + M("who") + "),\"\"))" },
                    new object[] { e == null ? "" : (object)(double)e.Expert, blank + "COUNTIFS(" + live + "," + M("who") + ",\"" + RdBook.Expert + "\"))" } };
                for (int i = 0; i < cells.Length; i++) sh.Set(r, i + 1, cells[i][0], i == 1 ? name : i == 15 ? day : formula, (string)cells[i][1]);
            }
            string range = "L6:L" + rd.ElementsLast;
            foreach (var color in ElementColors) sh.Rule(range, "expression", "$L6=" + Q(color[0]), RdDxf(color[2], true, color[1]));
            sh.Rule("F6:G" + rd.ElementsLast, "expression", "AND(ISNUMBER(F6),F6>0)", RdDxf("FF833C0B", true, "FFFCE4D6"));
            RdPut(RdElementsSheet, RdXml(sh), Sheet(RdElementsSheet) + "$A$1:$R$" + Math.Max(15, 5 + rd.Elements.Count));
        }
        // «Статистика»: statuses of the current sets and of the elements, remark types and the acts, on the report date of «СВОД».
        void WriteRdStatistics() {
            var sh = new RdSheet { Portrait = true }; var sets = rd.Sets.Where(s => s.Latest).ToList(); double report = rd.Days.Count == 0 ? 0 : rd.Days.Max();
            Titles(sh, "СТАТИСТИКА ПРОВЕРКИ РАБОЧЕЙ ДОКУМЕНТАЦИИ: СТАТУСЫ И ВИДЫ ЗАМЕЧАНИЙ", null);
            sh.Width(1, 1, 46); sh.Width(2, 8, 11);
            sh.Set(4, 1, "По состоянию на:", Rs(11, true, h: "left", wrap: false)); sh.Set(4, 2, report, Thin(11, color: Ink, fill: Gray, fmt: "dd.mm.yyyy"), Sheet(RdSummarySheet) + "$U$2");
            sh.Set(4, 3, "← дата отчёта с листа «СВОД»", Rs(9, color: Hint, h: "left", wrap: false));
            int head = Thin(11, true, fill: Head), headLeft = Thin(11, true, fill: Head, h: "left"), label = Thin(11, h: "left"), number = Thin(11, color: Ink, fill: Gray), percent = Thin(11, color: Ink, fill: Gray, fmt: "0.0");
            string live = M("latest") + ",1", done = M("status") + ",\"<>замечания\"," + M("status") + ",\"<>не проверен\"," + M("status") + ",\"<>не соответствует ВЭ РД\"";
            Func<RdSet, bool> resolved = s => s.Status != "замечания" && s.Status != "не проверен" && s.Status != "не соответствует ВЭ РД";
            int r = 6;
            Action<string, object[][]> block = (title, lines) => {
                sh.Set(r, 1, title, headLeft); sh.Set(r, 2, "Шт.", head); sh.Set(r, 3, "%", head); r++;
                int first = r; double total = Convert.ToDouble(lines[0][2], Inv);
                foreach (var line in lines) {
                    double value = Convert.ToDouble(line[2], Inv);
                    sh.Set(r, 1, line[0], label); sh.Set(r, 2, value, number, (string)line[1]);
                    sh.Set(r, 3, total == 0 ? 0 : value / total * 100, percent, "IFERROR(B" + r + "/B$" + first + "*100,0)"); r++;
                }
                r++;
            };
            string expert = M("who") + ",\"" + RdBook.Expert + "\"";
            block("Комплекты РД", new[] {
                new object[] { "В ведомости элементов РД (ВЭ РД)", "COUNTIF(" + Sheet(Ved) + "$AG$" + rd.Ved.First + ":$AG$" + rd.Ved.HelperLast + ",\"?*\")", rd.Ved.Sets },
                new object[] { "Представлено на проверку (действующие редакции из ВЭ РД)", "COUNTIFS(" + live + "," + M("ve") + ",\"да\")", sets.Count(s => s.Ve == "да") },
                new object[] { "Проверено по состоянию на дату отчёта", "COUNTIFS(" + live + "," + M("ve") + ",\"да\"," + M("date") + ",\"<=\"&$B$4," + M("who") + ",\"<>" + RdBook.Expert + "\")", sets.Count(s => s.Ve == "да" && s.Date.HasValue && s.Date.Value <= report && !s.Expert) },
                new object[] { "Готово к ВПР", "COUNTIFS(" + live + "," + M("status") + ",\"готово к ВПР\")", sets.Count(s => s.Status == "готово к ВПР") },
                new object[] { "С неустранёнными замечаниями", "COUNTIFS(" + live + "," + M("status") + ",\"замечания\")", sets.Count(s => s.Status == "замечания") },
                new object[] { "Замечания устранены", "COUNTIFS(" + live + "," + M("cnt") + ",\">0\"," + done + ")", sets.Count(s => s.Cnt > 0 && resolved(s)) },
                new object[] { "Не согласовано ТЭ", "COUNTIFS(" + live + "," + M("status") + ",\"не согласован ТЭ\")", sets.Count(s => s.Status == "не согласован ТЭ") },
                new object[] { "Замечания ТЭ («Проверил: ТЭ»)", "COUNTIFS(" + live + "," + expert + ")", sets.Count(s => s.Expert) },
                new object[] { "Не соответствуют ВЭ РД (нет в ведомости или другой элемент)", "COUNTIFS(" + live + "," + M("status") + ",\"не соответствует ВЭ РД\")", sets.Count(s => s.Status == "не соответствует ВЭ РД") } });
            Func<string, object[]> element = status => new object[] { status, "COUNTIF(" + E("L") + "," + Q(status) + ")", rd.Elements.Count(e => e.Status == status) };
            block("Элементы", new[] {
                new object[] { "В ведомости элементов РД", "SUM(" + Sheet(Ved) + "$AH$" + rd.Ved.First + ":$AH$" + rd.Ved.HelperLast + ")", rd.Ved.Elements },
                new object[] { "На проверке", "COUNTIF(" + E("A") + ",\"?*\")", rd.Elements.Count },
                Rename(element("готов к ВПР"), "Готовы к ВПР"), Rename(element("замечания"), "С замечаниями"), Rename(element("неполный комплект"), "Неполный комплект"),
                Rename(element("расхождение с ВЭ РД"), "Расхождение с ВЭ РД"), Rename(element("не согласован ТЭ"), "Не согласованы ТЭ"), Rename(element("в проверке"), "В проверке"),
                new object[] { "С замечаниями ТЭ", "COUNTIF(" + E("R") + ",\">0\")", rd.Elements.Count(e => e.Expert > 0) } });
            sh.Set(r, 1, "Виды замечаний (комплектов, действующие редакции)", headLeft); sh.Set(r, 2, "Всего", head); sh.Set(r, 3, "Устранено", head); sh.Set(r, 4, "Осталось", head); r++;
            Action<string, string, int, int> type = (text, criterion, all, fixedCount) => {
                sh.Set(r, 1, text, label); sh.Set(r, 2, all, number, "COUNTIFS(" + live + "," + criterion + ")");
                sh.Set(r, 3, fixedCount, number, "COUNTIFS(" + live + "," + criterion + "," + done + ")"); sh.Set(r, 4, all - fixedCount, number, "B" + r + "-C" + r); r++;
            };
            foreach (string code in rd.Layout.Flags) {
                Func<RdSet, bool> marked = s => { object x = s[code]; double n; return x is double ? (double)x == 1 : double.TryParse(Text(x), NumberStyles.Float, Inv, out n) && n == 1; };
                string header = Text(rd.Target.At(rd.Layout.Header, rd.Layout[code]));
                type(header == "" ? code : header, M(code) + ",1", sets.Count(marked), sets.Count(s => marked(s) && resolved(s)));
            }
            type("Примечание (текст в столбце «Примечание»)", M("note") + ",\"?*\"", sets.Count(s => Text(s["note"]) != ""), sets.Count(s => Text(s["note"]) != "" && resolved(s)));
            r++;
            sh.Set(r, 1, "Акты сдачи-приёмки", headLeft); sh.Set(r, 2, "Шт.", head); r++;
            var acts = rd.Acts;
            string results = acts.Exists && acts.Status > 0 ? A(acts.Status) : null;
            sh.Set(r, 1, "Проверено актов", label); sh.Set(r, 2, acts.Conform + acts.Remarks, number, results == null ? null : "COUNTIF(" + results + ",\"соответствует\")+COUNTIF(" + results + ",\"замечания\")"); r++;
            sh.Set(r, 1, "Из них с замечаниями", label); sh.Set(r, 2, acts.Remarks, number, results == null ? null : "COUNTIF(" + results + ",\"замечания\")");
            RdPut(RdStatisticsSheet, RdXml(sh), Sheet(RdStatisticsSheet) + "$A$1:$D$" + r);
        }
        static object[] Rename(object[] line, string label) { line[0] = label; return line; }
        // Left and right edges of a «СВОД» column: thick outside, medium between blocks, dashed between Шт. and %.
        static string[] SummaryEdges(int c) {
            string left = c == 1 ? "thick" : c <= 3 ? "medium" : (c - 4) % 4 == 0 ? "medium" : (c - 4) % 2 == 1 ? "dashed" : "thin";
            string right = c == 19 ? "thick" : c <= 3 ? "medium" : (c - 4) % 2 == 0 ? "dashed" : (c - 4) % 4 == 3 ? "medium" : "thin";
            return new[] { left, right };
        }
        static readonly string[] BlockColors = { Gold, Green, Red, Blue };
        int SummaryStyle(int c, string top, string bottom, bool bold, bool gray, string fmt = null, double? size = null, bool italic = false, string color = null, string h = "center") {
            var e = SummaryEdges(c);
            return Rs(size ?? 12, bold, italic, color ?? (c < 4 ? Black : BlockColors[(c - 4) / 4]), gray ? Gray : null, fmt, h, "center", true, e[0], e[1], top, bottom);
        }
        // «СВОД»: rows by date with the reviewers below (collapsed except under the report date), as the PD «СВОД»; elements and sets for boxes and volumes.
        void WriteRdSummary(DateTime asOf) {
            var sh = new RdSheet { Grid = false }; double report = asOf.Date.ToOADate();
            var latest = rd.Sets.Where(s => s.Latest).ToList(); var people = rd.People; var days = rd.Days;
            string team = M("who") + ",\"<>" + RdBook.Expert + "\"", lat = M("latest"), date = M("date"), who = M("who");
            double[] widths = { 10, 28, 10 };
            for (int c = 1; c <= 19; c++) sh.Width(c, c, c <= 3 ? widths[c - 1] : (c - 4) % 2 == 0 ? 11 : 7);
            sh.Width(21, 21, 14);
            sh.Merge(1, 1, 1, 19); sh.Set(1, 1, "СВОДНАЯ ТАБЛИЦА ПО РЕЗУЛЬТАТАМ ПРОВЕРКИ РАБОЧЕЙ ДОКУМЕНТАЦИИ", Rs(16, true));
            sh.Set(2, 1, Stage(), Rs(12, h: "left", wrap: false));
            sh.Merge(3, 1, 3, 19); sh.Set(3, 1, "Проверено — накопительно по датам и проверяющим; проверено за сутки — за указанный день. Комплекты — действующие редакции. Замечания ТЭ — отдельной строкой под таблицей.", Rs(12, h: "left", wrap: false));
            sh.Set(1, 21, "Дата отчёта", Rs(12, true)); sh.Set(2, 21, report, Rs(12, true, fill: "FFFFFF00", fmt: "dd.mm.yyyy"));
            sh.Set(3, 21, "← дата сборки, в печать не попадает", Rs(9, color: Red, h: "left", wrap: false));
            double[] heights = { 32, 24, 24, 34, 32, 30, 26 }; for (int i = 0; i < heights.Length; i++) sh.Heights[i + 1] = heights[i];
            // Sets with remarks above the table, as the volumes with remarks above the PD «СВОД».
            string checkedSets = lat + ",1," + M("cnt") + ",\">0\"," + date + ",\">0\"," + date + ",\"<=\"&$U$2," + team;
            string remarked = "COUNTIFS(" + checkedSets + ")", resolved = "(COUNTIFS(" + checkedSets + "," + M("status") + ",\"готово к ВПР\")+COUNTIFS(" + checkedSets + "," + M("status") + ",\"не согласован ТЭ\"))";
            var withRemarks = latest.Where(s => s.Cnt > 0 && s.Date.HasValue && s.Date.Value <= report && !s.Expert).ToList();
            int remarkedCount = withRemarks.Count, resolvedCount = withRemarks.Count(s => s.Status == "готово к ВПР" || s.Status == "не согласован ТЭ");
            var above = new[] { new object[] { 4, Gold, "Всего комплектов", "с замечаниями: ", remarked, remarkedCount }, new object[] { 8, Green, "Устранено замечаний", "по комплектам: ", resolved, resolvedCount },
                new object[] { 12, Red, "Устраняется комплектов", "с замечаниями: ", remarked + "-" + resolved, remarkedCount - resolvedCount } };
            foreach (var item in above) {
                int c = (int)item[0]; sh.Merge(4, c, 4, c + 3);
                for (int k = 0; k < 4; k++) sh.Set(4, c + k, k > 0 ? null : (string)item[2] + "\n" + (string)item[3] + item[5], Rs(12, true, color: (string)item[1]), k > 0 ? null : "\"" + item[2] + "\"&CHAR(10)&\"" + item[3] + "\"&(" + item[4] + ")");
            }
            // Header: three rows like PD, «элементов» and «комплектов» in place of «коробов» and «томов».
            string[] blocks = { "Всего", "Проверено", "Осталось проверить", "Проверено за сутки" };
            for (int r = 5; r <= 7; r++) for (int c = 1; c <= 19; c++) {
                object text = null;
                if (r == 5 && c <= 3) text = new[] { "Этап", "Дата / проверяющий", "Время" }[c - 1];
                if (c >= 4) { int k = (c - 4) % 4; if (r == 5 && k == 0) text = blocks[(c - 4) / 4]; if (r == 6 && k % 2 == 0) text = k == 0 ? "элементов" : "комплектов"; if (r == 7) text = k % 2 == 0 ? "Шт." : "%"; }
                sh.Set(r, c, text, SummaryStyle(c, r == 5 ? "thick" : "thin", "thin", true, false));
            }
            for (int c = 1; c <= 3; c++) sh.Merge(5, c, 7, c);
            for (int c = 4; c <= 16; c += 4) { sh.Merge(5, c, 5, c + 3); sh.Merge(6, c, 6, c + 1); sh.Merge(6, c + 2, 6, c + 3); }
            int last = 7 + days.Count * (1 + people.Count), row = 8;
            Func<int, bool> person = c => c >= 8 && c < 12 || c >= 16;
            for (int index = 0; index < days.Count; index++) {
                double day = days[index]; bool gray = index % 2 == 1, lastDay = index == days.Count - 1; int top = row; string d = "$B" + top;
                Func<string, string> cut = f => "IF(" + d + ">$U$2,\"\"," + f + ")";
                var counts = new Dictionary<int, double> {
                    { 4, rd.Elements.Count }, { 6, latest.Count },
                    { 8, rd.Elements.Count(e => e.Date.HasValue && e.Date.Value > 0 && e.Date.Value <= day) }, { 10, latest.Count(s => s.Date.HasValue && s.Date.Value <= day && !s.Expert) } };
                counts[12] = counts[4] - counts[8]; counts[14] = counts[6] - counts[10];
                counts[16] = rd.Elements.Count(e => e.Date == day); counts[18] = latest.Count(s => s.Date == day && !s.Expert);
                var formulas = new Dictionary<int, string> {
                    { 4, "COUNTIF(" + E("A") + ",\"?*\")" }, { 6, "COUNTIF(" + lat + ",1)" }, { 8, "COUNTIFS(" + E("P") + ",\">0\"," + E("P") + ",\"<=\"&" + d + ")" },
                    { 10, "COUNTIFS(" + lat + ",1," + date + ",\">0\"," + date + ",\"<=\"&" + d + "," + team + ")" }, { 12, "D" + top + "-H" + top }, { 14, "F" + top + "-J" + top },
                    { 16, "COUNTIF(" + E("P") + "," + d + ")" }, { 18, "COUNTIFS(" + lat + ",1," + date + "," + d + "," + team + ")" } };
                sh.Set(row, 2, day, SummaryStyle(2, "thin", row == last ? "thick" : "thin", true, gray, "dd.mm.yyyy"));
                sh.Set(row, 3, 17.0 / 24, SummaryStyle(3, "thin", row == last ? "thick" : "thin", true, gray, "h:mm"));
                SummaryNumbers(sh, row, top, counts, formulas, cut, true, gray, row == last, c => true);
                row++;
                // A row per reviewer: what they checked up to the day and on the day.
                foreach (string name in people) {
                    string n = "$B" + row;
                    var mine = new Dictionary<int, double> {
                        { 8, rd.Elements.Count(e => Is(e.Who, name) && e.Date.HasValue && e.Date.Value > 0 && e.Date.Value <= day) }, { 10, latest.Count(s => Is(s.Who, name) && s.Date.HasValue && s.Date.Value <= day) },
                        { 16, rd.Elements.Count(e => Is(e.Who, name) && e.Date == day) }, { 18, latest.Count(s => Is(s.Who, name) && s.Date == day) } };
                    var mineFormulas = new Dictionary<int, string> {
                        { 8, "COUNTIFS(" + E("Q") + "," + n + "," + E("P") + ",\">0\"," + E("P") + ",\"<=\"&" + d + ")" }, { 10, "COUNTIFS(" + lat + ",1," + who + "," + n + "," + date + ",\">0\"," + date + ",\"<=\"&" + d + ")" },
                        { 16, "COUNTIFS(" + E("Q") + "," + n + "," + E("P") + "," + d + ")" }, { 18, "COUNTIFS(" + lat + ",1," + who + "," + n + "," + date + "," + d + ")" } };
                    sh.Set(row, 2, name, SummaryStyle(2, "thin", row == last ? "thick" : "thin", false, gray));
                    sh.Set(row, 3, 17.0 / 24, SummaryStyle(3, "thin", row == last ? "thick" : "thin", false, gray, "h:mm"));
                    SummaryNumbers(sh, row, top, mine, mineFormulas, cut, false, gray, row == last, person);
                    sh.Detail.Add(row); if (!lastDay) sh.Hidden.Add(row);
                    row++;
                }
                if (!lastDay && people.Count > 0) sh.Collapsed.Add(top);
            }
            if (days.Count > 0) {
                sh.Merge(8, 1, last, 1);
                for (int r = 8; r <= last; r++) sh.Set(r, 1, r == 8 ? StageLabel() : null, SummaryStyle(1, "thin", r == last ? "thick" : "thin", false, false));
            }
            // Remarks of the Technical expert under the table: not in the reviewers' statistics, the sets wait for ТЭ to agree.
            int t = last + 2, lastTop = last - people.Count;
            sh.Merge(t, 1, t, 3); int box = Rs(11, true, color: Red, l: "thin", r: "thin", t: "thin", b: "thin");
            for (int c = 1; c <= 3; c++) sh.Set(t, c, c == 1 ? "Замечания ТЭ («Проверил: ТЭ»), в статистику проверяющих не входят" : null, box);
            double expertElements = rd.Elements.Count(e => e.Expert > 0), expertSets = latest.Count(s => s.Expert), lastElements = rd.Elements.Count, lastSets = latest.Count;
            sh.Set(t, 4, expertElements, Rs(12, true, color: Red, fmt: "0", l: "thin", r: "thin", t: "thin", b: "thin"), "COUNTIF(" + E("R") + ",\">0\")");
            sh.Set(t, 5, lastElements == 0 ? 0 : expertElements / lastElements * 100, Rs(9, true, true, Red, fmt: "0.0", l: "thin", r: "thin", t: "thin", b: "thin"), "IFERROR(D" + t + "/$D$" + lastTop + "*100,0)");
            sh.Set(t, 6, expertSets, Rs(12, true, color: Red, fmt: "0", l: "thin", r: "thin", t: "thin", b: "thin"), "COUNTIFS(" + lat + ",1," + who + ",\"" + RdBook.Expert + "\")");
            sh.Set(t, 7, lastSets == 0 ? 0 : expertSets / lastSets * 100, Rs(9, true, true, Red, fmt: "0.0", l: "thin", r: "thin", t: "thin", b: "thin"), "IFERROR(F" + t + "/$F$" + lastTop + "*100,0)");
            sh.Set(t, 8, "элементов / комплектов, ждут согласования ТЭ", Rs(10, color: Red, h: "left", wrap: false)); sh.Heights[t] = 30;
            sh.Merge(t + 2, 1, t + 2, 19); sh.Set(t + 2, 1, "Примечание:", Rs(12, h: "left", wrap: false));
            RdPut(RdSummarySheet, RdXml(sh), Sheet(RdSummarySheet) + "$A$1:$S$" + (t + 2));
        }
        // The Шт. and % columns of a «СВОД» row: an empty cell where the row shows nothing, the share of the date row's totals in %.
        void SummaryNumbers(RdSheet sh, int row, int top, Dictionary<int, double> counts, Dictionary<int, string> formulas, Func<string, string> cut, bool bold, bool gray, bool bottom, Func<int, bool> shown) {
            var totals = sh.Rows[top];
            for (int c = 4; c <= 19; c++) {
                string edge = bottom ? "thick" : "thin";
                if (!shown(c)) { sh.Set(row, c, null, SummaryStyle(c, "thin", edge, bold, gray)); continue; }
                if (c % 2 == 0) { double value; counts.TryGetValue(c, out value); sh.Set(row, c, value, SummaryStyle(c, "thin", edge, bold, gray, "0"), cut(formulas[c])); continue; }
                int totalCol = (c - 4) % 4 < 2 ? 4 : 6; double count; counts.TryGetValue(c - 1, out count);
                double total = row == top ? (totalCol == 4 ? counts[4] : counts[6]) : Convert.ToDouble(Value(totals[totalCol]) ?? 0.0, Inv);
                string totalRef = row == top ? "$" + MergeEngine.ExcelColumn(totalCol) + top : "$" + MergeEngine.ExcelColumn(totalCol) + "$" + top;
                sh.Set(row, c, total == 0 ? 0 : count / total * 100, SummaryStyle(c, "thin", edge, bold, gray, "0.0", 9, true), cut("IFERROR(" + MergeEngine.ExcelColumn(c - 1) + row + "/" + totalRef + "*100,0)"));
            }
        }
        // «Итоги проверки»: per day and reviewer — elements finished and sets checked; below, the same by reviewer and day.
        void WriteRdTotals(DateTime asOf) {
            var sh = new RdSheet { Grid = false }; var people = rd.People; double report = asOf.Date.ToOADate();
            int width = 2 + 2 * (people.Count + 1);
            sh.Width(1, 1, 10); sh.Width(2, 2, 18); sh.Width(3, width, 13);
            sh.Set(1, 1, "СВОДНАЯ ТАБЛИЦА ПО ИТОГАМ РАБОТЫ ПРОВЕРЯЮЩИХ", Rs(16, true, h: "left", wrap: false));
            sh.Set(2, 1, Stage(), Rs(12, h: "left", wrap: false));
            sh.Set(3, 1, "Элементов — элементы, проверка которых завершена в этот день; комплектов — все проверки за день, включая повторные. Замечания ТЭ не входят.", Rs(10, italic: true, h: "left", wrap: false));
            sh.Heights[1] = 32; sh.Heights[2] = 24;
            Func<int, string[]> edges = c => c == 1 ? new[] { "thick", "thin" } : c == 2 ? new[] { "thin", "medium" } : c % 2 == 1 ? new[] { "medium", "thin" } : new[] { "thin", c == width ? "thick" : "medium" };
            Func<int, string, string, bool, bool, string, string, int> style = (c, top, bottom, bold, gray, fmt, color) => { var e = edges(c); return Rs(12, bold, false, color ?? Black, gray ? Gray : null, fmt, "center", "center", true, e[0], e[1], top, bottom); };
            var names = people.Concat(new[] { "ВСЕГО" }).ToList();
            for (int r = 4; r <= 5; r++) {
                sh.Heights[r] = 30;
                for (int c = 1; c <= width; c++) {
                    object text = null;
                    if (r == 4 && c == 1) text = "Этап"; if (r == 4 && c == 2) text = "Дата";
                    if (c >= 3) { int k = (c - 3) / 2; if (r == 4 && (c - 3) % 2 == 0) text = names[k]; if (r == 5) text = (c - 3) % 2 == 0 ? "элементов" : "комплектов"; }
                    sh.Set(r, c, text, style(c, r == 4 ? "thick" : "thin", "thin", true, false, null, null));
                }
            }
            sh.Merge(4, 1, 5, 1); sh.Merge(4, 2, 5, 2); for (int k = 0; k < names.Count; k++) sh.Merge(4, 3 + 2 * k, 4, 4 + 2 * k);
            string who = M("who"), date = M("date"), cnt = M("cnt"), reportDate = Sheet(RdSummarySheet) + "$U$2";
            int last = 6 + rd.Days.Count;
            Action<int, double?> counts = (r, day) => {
                double sumElements = 0, sumSets = 0; bool gray = day.HasValue && (r - 7) % 2 == 1, bottom = r == last;
                for (int k = 0; k < people.Count; k++) {
                    int c = 3 + 2 * k; string name = MergeEngine.ExcelColumn(c) + "$4", person = people[k];
                    double elements = rd.Elements.Count(e => Is(e.Who, person) && e.Date.HasValue && (day.HasValue ? e.Date == day : e.Date.Value > 0 && e.Date.Value <= report));
                    double sets = rd.Sets.Count(s => Is(s.Who, person) && s.Date.HasValue && (day.HasValue ? s.Date == day : s.Date.Value <= report));
                    string[] when = day.HasValue ? new[] { E("P") + ",$B" + r, date + ",$B" + r } : new[] { E("P") + ",\">0\"," + E("P") + ",\"<=\"&" + reportDate, date + ",\">0\"," + date + ",\"<=\"&" + reportDate };
                    sh.Set(r, c, elements, style(c, "thin", bottom ? "thick" : "thin", !day.HasValue || elements != 0, gray, "0", day.HasValue ? null : Green), "COUNTIFS(" + E("Q") + "," + name + "," + when[0] + ")");
                    sh.Set(r, c + 1, sets, style(c + 1, "thin", bottom ? "thick" : "thin", !day.HasValue || sets != 0, gray, "0", day.HasValue ? null : Green), "COUNTIFS(" + who + "," + name + "," + when[1] + ")");
                    sumElements += elements; sumSets += sets;
                }
                int t = 3 + 2 * people.Count;
                string first = people.Count == 0 ? "0" : string.Join("+", Enumerable.Range(0, people.Count).Select(k => MergeEngine.ExcelColumn(3 + 2 * k) + r)), second = people.Count == 0 ? "0" : string.Join("+", Enumerable.Range(0, people.Count).Select(k => MergeEngine.ExcelColumn(4 + 2 * k) + r));
                sh.Set(r, t, sumElements, style(t, "thin", bottom ? "thick" : "thin", !day.HasValue || sumElements != 0, gray, "0", day.HasValue ? null : Green), first);
                sh.Set(r, t + 1, sumSets, style(t + 1, "thin", bottom ? "thick" : "thin", !day.HasValue || sumSets != 0, gray, "0", day.HasValue ? null : Green), second);
            };
            sh.Set(6, 2, "Всего", style(2, "thin", last == 6 ? "thick" : "thin", true, false, null, null)); counts(6, null); sh.Heights[6] = 24;
            for (int i = 0; i < rd.Days.Count; i++) {
                int r = 7 + i; bool gray = i % 2 == 1; sh.Heights[r] = 24;
                sh.Set(r, 2, rd.Days[i], style(2, "thin", r == last ? "thick" : "thin", false, gray, "dd.mm.yyyy", null)); counts(r, rd.Days[i]);
            }
            sh.Merge(6, 1, last, 1); for (int r = 6; r <= last; r++) sh.Set(r, 1, r == 6 ? StageLabel() : null, style(1, "thin", r == last ? "thick" : "thin", false, false, null, null));
            // Non-zero numbers are bold, zeros regular, as in the PD «Итоги работы»; the rules keep it so when Excel recalculates.
            if (last >= 7) { string area = "C7:" + MergeEngine.ExcelColumn(width) + last; sh.Rule(area, "cellIs", "0", RdDxf(null, true, null), "notEqual"); sh.Rule(area, "cellIs", "0", RdDxf(null, false, null), "equal"); }
            // The same work by reviewer, with the days each of them worked.
            int row = last + 3; sh.Merge(row, 1, row, 7); sh.Set(row, 1, "Альтернативное представление по проверяющим", Rs(12, true, h: "left", wrap: false)); row++;
            string[] heads = { "Этап", "Проверяющий / дата", "Проверено комплектов", "Комплектов с замечаниями", "Элементов завершено", "Проверено актов" };
            int alt = row; for (int c = 0; c < heads.Length; c++) sh.Set(row, c + 1, heads[c], Rs(12, true, l: "thin", r: "thin", t: "thin", b: "thin")); sh.Heights[row] = 62; row++;
            int firstAlt = row; var acts = rd.Acts;
            foreach (string name in people) {
                SortedSet<double> worked; rd.Activity.TryGetValue(name, out worked);
                int personRow = row;
                foreach (double? day in new double?[] { null }.Concat((worked ?? new SortedSet<double>()).Select(d => (double?)d))) {
                    string crit = day.HasValue ? ",$B" + row : ",\">0\"", n = "$B$" + personRow; bool bold = !day.HasValue;
                    Func<double?, bool> on = x => x.HasValue && (day.HasValue ? x == day : x.Value > 0);
                    var whose = rd.Sets.Where(s => Is(s.Who, name) && on(s.Date)).ToList();
                    sh.Set(row, 2, day.HasValue ? (object)day.Value : name, Rs(12, bold, fmt: day.HasValue ? "dd.mm.yyyy" : null, l: "thin", r: "thin", t: "thin", b: "thin"));
                    int number = Rs(12, bold, fmt: "0", l: "thin", r: "thin", t: "thin", b: "thin");
                    sh.Set(row, 3, whose.Count, number, "COUNTIFS(" + who + "," + n + "," + date + crit + ")");
                    sh.Set(row, 4, whose.Count(s => s.Cnt > 0), number, "COUNTIFS(" + who + "," + n + "," + date + crit + "," + cnt + ",\">0\")");
                    sh.Set(row, 5, rd.Elements.Count(e => Is(e.Who, name) && on(e.Date)), number, "COUNTIFS(" + E("Q") + "," + n + "," + E("P") + crit + ")");
                    sh.Set(row, 6, acts.Checks.Count(x => Is(x.Item1, name) && on(x.Item2)), number, acts.Exists ? "COUNTIFS(" + A(acts.Who) + "," + n + "," + A(acts.Date) + crit + ")" : null);
                    row++;
                }
            }
            if (row > firstAlt) { sh.Merge(firstAlt, 1, row - 1, 1); for (int r = firstAlt; r < row; r++) sh.Set(r, 1, r == firstAlt ? StageLabel() : null, Rs(12, l: "thin", r: "thin", t: "thin", b: "thin")); }
            // The first range is the table of the PDF, as the first table of the PD «Итоги работы».
            string sheet = Sheet(RdEngine.RdWork);
            RdPut(RdEngine.RdWork, RdXml(sh), sheet + "$A$1:$" + MergeEngine.ExcelColumn(width) + "$" + last + "," + sheet + "$A$" + (alt - 1) + ":$F$" + Math.Max(alt, row - 1));
        }
        // Ranges of other sheets that end at the old last row of the check sheet follow it down (e.g. sheets the user added).
        void GrowMainReferences() {
            if (rd.Last <= rd.OldLast) return;
            string name = Regex.Escape(rd.MainName.Replace("'", "''"));
            var pattern = new Regex(@"((?:'" + name + "'|" + name + @")!\$[A-Z]{1,3}\$" + rd.Layout.First + @":\$[A-Z]{1,3}\$)" + rd.OldLast + @"(?!\d)");
            var own = new[] { rd.MainName, RdElementsSheet, RdStatisticsSheet, RdSummarySheet, RdEngine.RdWork };
            foreach (var sheet in book.Root.Element(N + "sheets").Elements().ToList()) {
                string sheetName = (string)sheet.Attribute("name"); if (own.Any(o => Is(o, sheetName))) continue;
                string path = PathFor(sheetName); if (path == null || !entries.ContainsKey(path)) continue;
                var doc = Xml(path); bool changed = false;
                foreach (var f in doc.Descendants(N + "f")) { string grown = pattern.Replace(f.Value, "${1}" + rd.Last); if (grown != f.Value) { f.Value = grown; changed = true; } }
                if (changed) Store(path, doc);
            }
        }
        void FinishRd(string output, Action<int, string> progress, CancellationToken token) {
            var calc = book.Root.Element(N + "calcPr"); if (calc == null) { calc = new XElement(N + "calcPr"); book.Root.Add(calc); }
            calc.SetAttributeValue("calcMode", "auto"); calc.SetAttributeValue("fullCalcOnLoad", 1); calc.SetAttributeValue("forceFullCalc", null); calc.SetAttributeValue("calcId", 0);
            rels.Root.Elements().Where(e => ((string)e.Attribute("Type") ?? "").EndsWith("/calcChain")).Remove(); types.Root.Elements().Where(e => ((string)e.Attribute("PartName") ?? "").EndsWith("/calcChain.xml")).Remove(); entries.Remove("xl/calcChain.xml");
            var names = book.Root.Element(N + "definedNames"); if (names != null && !names.Elements().Any()) names.Remove();
            int active = SheetIndex(RdSummarySheet);
            var view = book.Root.Element(N + "bookViews"); if (view == null) { view = new XElement(N + "bookViews"); book.Root.Element(N + "sheets").AddBeforeSelf(view); }
            if (!view.Elements().Any()) view.Add(new XElement(N + "workbookView"));
            foreach (var workbookView in view.Elements()) { workbookView.SetAttributeValue("activeTab", active); workbookView.SetAttributeValue("firstSheet", null); }
            foreach (var sheet in book.Root.Element(N + "sheets").Elements()) {
                string sheetName = (string)sheet.Attribute("name"), path = PathFor(sheetName); if (path == null || !entries.ContainsKey(path)) continue;
                var doc = Xml(path); bool changed = false;
                foreach (var sheetView in doc.Descendants(N + "sheetView")) { string selected = Is(sheetName, RdSummarySheet) ? "1" : null; if ((string)sheetView.Attribute("tabSelected") != selected) { sheetView.SetAttributeValue("tabSelected", selected); changed = true; } }
                if (changed) Store(path, doc);
            }
            Store("xl/workbook.xml", book); Store("xl/_rels/workbook.xml.rels", rels); Store("[Content_Types].xml", types); Store("xl/styles.xml", styles);
            progress(90, "Сохранение XLSX с формулами");
            using (var s = new FileStream(output, FileMode.CreateNew)) using (var z = new ZipArchive(s, ZipArchiveMode.Create)) {
                foreach (var item in entries) { token.ThrowIfCancellationRequested(); var e = z.CreateEntry(item.Key, CompressionLevel.Optimal); using (var stream = e.Open()) stream.Write(item.Value, 0, item.Value.Length); }
            }
        }
    }
}
