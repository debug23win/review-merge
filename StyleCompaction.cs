using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace ReviewMerge {
    public sealed partial class XlsxWriter {
        // Every build adds the fonts, fills, borders and formats it needs; the identical ones of earlier builds are merged here and the
        // cells, rows and columns of every sheet point at the kept ones, so styles.xml does not grow from build to build.
        void CompactStyles() {
            var root = styles.Root;
            var fonts = Dedupe(root.Element(N + "fonts")); var fills = Dedupe(root.Element(N + "fills")); var borders = Dedupe(root.Element(N + "borders"));
            // Number formats are the same when their codes are: the first id of a code is kept.
            var numbers = new Dictionary<int, int>(); var formats = root.Element(N + "numFmts");
            if (formats != null) {
                var first = new Dictionary<string, int>();
                foreach (var f in formats.Elements(N + "numFmt").ToList()) {
                    int id; string code = (string)f.Attribute("formatCode") ?? ""; if (!int.TryParse((string)f.Attribute("numFmtId"), out id)) continue;
                    int kept; if (first.TryGetValue(code, out kept)) { numbers[id] = kept; f.Remove(); } else first[code] = id;
                }
                formats.SetAttributeValue("count", formats.Elements().Count());
            }
            foreach (var xf in root.Elements(N + "cellStyleXfs").Elements().Concat(root.Elements(N + "cellXfs").Elements())) {
                Remap(xf, "fontId", fonts); Remap(xf, "fillId", fills); Remap(xf, "borderId", borders); Remap(xf, "numFmtId", numbers);
            }
            foreach (var f in root.Elements(N + "dxfs").Elements().Elements(N + "numFmt")) Remap(f, "numFmtId", numbers);
            var cells = Dedupe(root.Element(N + "cellXfs"));
            // Table styles, tables, pivot tables and slicers refer to differential formats by number too; such books keep their dxfs as they are.
            var tables = root.Element(N + "tableStyles");
            bool dxfUsers = tables != null && tables.Elements().Any() || entries.Keys.Any(k => k.StartsWith("xl/tables/") || k.StartsWith("xl/pivot") || k.StartsWith("xl/slicer") || k.StartsWith("xl/timeline"));
            var dxfs = dxfUsers ? null : Dedupe(root.Element(N + "dxfs"));
            bool cellsMoved = cells != null && cells.Any(p => p.Key != p.Value), dxfsMoved = dxfs != null && dxfs.Any(p => p.Key != p.Value);
            if (!cellsMoved && !dxfsMoved) return;
            foreach (string path in entries.Keys.Where(k => k.StartsWith("xl/worksheets/") && k.EndsWith(".xml")).ToList()) {
                var doc = Xml(path); bool changed = false;
                if (cellsMoved) {
                    foreach (var e in doc.Root.Descendants().Where(e => e.Name == N + "c" || e.Name == N + "row")) changed |= Remap(e, "s", cells);
                    foreach (var e in doc.Root.Descendants(N + "col")) changed |= Remap(e, "style", cells);
                }
                if (dxfsMoved) foreach (var e in doc.Root.Descendants(N + "cfRule")) changed |= Remap(e, "dxfId", dxfs);
                if (changed) Store(path, doc);
            }
        }
        // Merges identical children of a style list in place; the map gives the new position of every old one (null without the list).
        static Dictionary<int, int> Dedupe(XElement list) {
            if (list == null) return null;
            var map = new Dictionary<int, int>(); var seen = new Dictionary<string, int>(); var kept = new List<XElement>(); int index = 0;
            foreach (var e in list.Elements().ToList()) {
                string key = e.ToString(SaveOptions.DisableFormatting); int at;
                if (!seen.TryGetValue(key, out at)) { at = kept.Count; seen[key] = at; kept.Add(e); }
                map[index++] = at;
            }
            if (kept.Count == index) return map;
            list.RemoveNodes(); list.Add(kept); list.SetAttributeValue("count", kept.Count);
            return map;
        }
        static bool Remap(XElement e, string attribute, Dictionary<int, int> map) {
            var a = e.Attribute(attribute); int from, to;
            if (map == null || a == null || !int.TryParse(a.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out from) || !map.TryGetValue(from, out to) || to == from) return false;
            a.Value = to.ToString(CultureInfo.InvariantCulture); return true;
        }
    }
}
