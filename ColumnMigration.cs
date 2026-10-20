using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ReviewMerge {
    public sealed partial class XlsxWriter {
        static string ShiftColumnAddress(string value) {
            return Regex.Replace(value??"",@"(?<![\p{L}\p{N}_])(?<d>\$?)(?<c>[A-Z]{1,3})(?=\$?\d+|:|$|[\s,])",m=>{
                int col=Column(m.Groups["c"].Value);if(col<13)return m.Value;
                if(col>=16384)throw new System.IO.InvalidDataException("Недостаточно столбцов Excel для добавления номера тома.");
                return m.Groups["d"].Value+MergeEngine.ExcelColumn(col+1);
            });
        }
        static string ShiftMainFormula(string formula,string sheetName,bool local) {
            return Regex.Replace(formula??"",@"""(?:[^""]|"""")*""|(?<![\p{L}\p{N}_])(?:(?<sheet>'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_. ]*)!)?(?<a>\$?[A-Z]{1,3}(?:\$?\d+)?)(?::(?<b>\$?[A-Z]{1,3}(?:\$?\d+)?))?(?![\p{L}\p{N}_.]|\s*\()",m=>{
                if(!m.Groups["a"].Success)return m.Value;
                if(!Regex.IsMatch(m.Groups["a"].Value,@"\d")&&!m.Groups["b"].Success)return m.Value;
                string sheet=m.Groups["sheet"].Value.Trim('\'').Replace("''","'");
                if(sheet==""&&!local||sheet!=""&&!string.Equals(sheet,sheetName,StringComparison.OrdinalIgnoreCase))return m.Value;
                string prefix=m.Groups["sheet"].Success?m.Groups["sheet"].Value+"!":"";
                return prefix+ShiftColumnAddress(m.Groups["a"].Value)+(m.Groups["b"].Success?":"+ShiftColumnAddress(m.Groups["b"].Value):"");
            });
        }
        // A filter criterion is numbered from the first column of its range: criteria at M or to the right follow the inserted field.
        static void ShiftFilterColumns(XElement filter) {
            if(filter==null)return;int start=Column(Regex.Match((string)filter.Attribute("ref")??"",@"[A-Z]+").Value);if(start==0||start>=13)return;
            foreach(var c in filter.Elements(N+"filterColumn")){int id;if(int.TryParse((string)c.Attribute("colId"),out id)&&start+id>=13)c.SetAttributeValue("colId",id+1);}
        }
        // The part a relationship of part from points to: "/xl/comments/comment1.xml" or "../comments1.xml".
        static string PartPath(string from,string target) {
            target=target.Replace('\\','/');if(target.StartsWith("/",StringComparison.Ordinal))return target.TrimStart('/');
            var parts=from.Substring(0,Math.Max(0,from.LastIndexOf('/'))).Split(new[]{'/'},StringSplitOptions.RemoveEmptyEntries).ToList();
            foreach(string p in target.Split('/')){if(p==".."){if(parts.Count>0)parts.RemoveAt(parts.Count-1);}else if(p!="."&&p!="")parts.Add(p);}
            return string.Join("/",parts);
        }
        // Legacy drawing of notes and controls: zero-based columns of the anchors (VML is not always well-formed XML, so it is edited as text).
        static byte[] ShiftVml(byte[] data) {
            var latin=System.Text.Encoding.GetEncoding(28591);string text=latin.GetString(data);
            Func<string,string> shift=s=>{int n=int.Parse(s.Trim(),System.Globalization.CultureInfo.InvariantCulture);return n>=12?s.Replace(n.ToString(System.Globalization.CultureInfo.InvariantCulture),(n+1).ToString(System.Globalization.CultureInfo.InvariantCulture)):s;};
            text=Regex.Replace(text,@"(<(?:\w+:)?Column>)(\s*\d+\s*)(</(?:\w+:)?Column>)",m=>m.Groups[1].Value+shift(m.Groups[2].Value)+m.Groups[3].Value);
            text=Regex.Replace(text,@"(<(?:\w+:)?Anchor>)([^<]*)(</(?:\w+:)?Anchor>)",m=>{var v=m.Groups[2].Value.Split(',');if(v.Length!=8||!v.All(p=>Regex.IsMatch(p,@"^\s*\d+\s*$")))return m.Value;v[0]=shift(v[0]);v[4]=shift(v[4]);return m.Groups[1].Value+string.Join(",",v)+m.Groups[3].Value;});
            return latin.GetBytes(text);
        }
        void InsertVolumeColumn(XDocument main,string name,int header,int first) {
            ExpandSharedFormulas(main);
            foreach(var filter in main.Descendants(N+"autoFilter"))ShiftFilterColumns(filter);
            int headerStyle=StyleAt(main,header,13,headStyle);
            var rows=main.Root.Element(N+"sheetData").Elements(N+"row").ToList();
            foreach(var row in rows)foreach(var cell in row.Elements(N+"c")) {
                cell.SetAttributeValue("r",ShiftColumnAddress((string)cell.Attribute("r")));
                var f=cell.Element(N+"f");if(f!=null){f.Value=ShiftMainFormula(f.Value,name,true);if(f.Attribute("ref")!=null)f.SetAttributeValue("ref",ShiftColumnAddress((string)f.Attribute("ref")));}
            }
            foreach(var element in main.Descendants().Where(e=>e.Name.LocalName!="c"&&e.Name.LocalName!="f"))foreach(string attribute in new[]{"ref","sqref","topLeftCell","activeCell"}) {
                var a=element.Attribute(attribute);if(a!=null)a.Value=ShiftColumnAddress(a.Value);
            }
            foreach(var f in main.Descendants().Where(e=>e.Name==N+"formula"||e.Name==N+"formula1"||e.Name==N+"formula2"))f.Value=ShiftMainFormula(f.Value,name,true);
            // Frozen columns that include M keep the inserted field frozen as well.
            foreach(var pane in main.Descendants(N+"pane").Where(p=>((string)p.Attribute("state")??"").StartsWith("frozen",StringComparison.Ordinal))){int x;if(int.TryParse((string)pane.Attribute("xSplit"),out x)&&x>=13)pane.SetAttributeValue("xSplit",x+1);}
            var cols=main.Root.Element(N+"cols");
            if(cols==null){cols=new XElement(N+"cols");main.Root.Element(N+"sheetData").AddBeforeSelf(cols);}
            foreach(var c in cols.Elements().ToList()) {
                int min=(int)c.Attribute("min"),max=(int)c.Attribute("max");
                if(min>=16384){c.Remove();continue;}
                if(min>=13){c.SetAttributeValue("min",min+1);c.SetAttributeValue("max",Math.Min(16384,max+1));}
                else if(max>=13){var tail=new XElement(c);tail.SetAttributeValue("min",14);tail.SetAttributeValue("max",Math.Min(16384,max+1));c.SetAttributeValue("max",12);c.AddAfterSelf(tail);}
            }
            cols.Add(Col(13,13,14));var sorted=cols.Elements().OrderBy(c=>(int)c.Attribute("min")).ToList();sorted.Remove();cols.Add(sorted);
            var title=rows.First(r=>(int)r.Attribute("r")==header);SetCell(title,Cell(header,13,"Номер тома",headerStyle));
            var numbers=rows.FirstOrDefault(r=>(int)r.Attribute("r")==header+1);
            if(numbers!=null&&Text(Value(numbers.Elements(N+"c").FirstOrDefault(c=>(string)c.Attribute("r")=="A"+(header+1))))=="1") {
                SetCell(numbers,Cell(header+1,13,13,StyleAt(main,header+1,14,bodyStyle)));
                foreach(var cell in numbers.Elements(N+"c").ToList()){int c=Column((string)cell.Attribute("r"));var v=Value(cell);
                    if(c>13&&cell.Element(N+"f")==null&&v is double&&(double)v==c-1)SetCell(numbers,Cell(header+1,c,c,(int?)cell.Attribute("s")??bodyStyle));}
            }
            foreach(var row in rows){var cells=row.Elements(N+"c").OrderBy(c=>Column((string)c.Attribute("r"))).ToList();cells.Remove();row.Add(cells);}
            foreach(var sheet in book.Root.Element(N+"sheets").Elements()) {
                string path=PathFor((string)sheet.Attribute("name"));if(path==mainPath)continue;var doc=Xml(path);bool changed=false;
                foreach(var f in doc.Descendants(N+"f")){string next=ShiftMainFormula(f.Value,name,false);if(next!=f.Value){f.Value=next;changed=true;}}
                if(changed)Store(path,doc);
            }
            foreach(var defined in book.Descendants(N+"definedName"))defined.Value=ShiftMainFormula(defined.Value,name,false);
            // Chart series and pivot sources that read the main sheet.
            foreach(string part in entries.Keys.Where(k=>k.StartsWith("xl/charts/",StringComparison.Ordinal)&&k.EndsWith(".xml",StringComparison.Ordinal)&&!k.Contains("/_rels/")).ToList()) {
                var chart=Xml(part);bool changed=false;
                foreach(var f in chart.Descendants().Where(e=>e.Name.LocalName=="f"&&e.Name.NamespaceName.Contains("drawingml"))){string next=ShiftMainFormula(f.Value,name,false);if(next!=f.Value){f.Value=next;changed=true;}}
                if(changed)Store(part,chart);
            }
            foreach(string part in entries.Keys.Where(k=>k.StartsWith("xl/pivotCache/pivotCacheDefinition",StringComparison.Ordinal)&&k.EndsWith(".xml",StringComparison.Ordinal)).ToList()) {
                var cache=Xml(part);bool changed=false;
                foreach(var source in cache.Descendants(N+"worksheetSource").Where(s=>s.Attribute("ref")!=null&&string.Equals(((string)s.Attribute("sheet")??"").Trim(),name.Trim(),StringComparison.OrdinalIgnoreCase))){source.SetAttributeValue("ref",ShiftColumnAddress((string)source.Attribute("ref")));changed=true;}
                if(changed)Store(part,cache);
            }
            // Native Excel table ranges and column definitions follow the inserted field.
            var sheetRels=mainPath.Replace("/sheet","/_rels/sheet")+".rels";
            if(entries.ContainsKey(sheetRels))foreach(var rel in Xml(sheetRels).Root.Elements()) {
                string type=(string)rel.Attribute("Type")??"",target=(string)rel.Attribute("Target")??"";
                if((string)rel.Attribute("TargetMode")=="External")continue;
                string part=PartPath(mainPath,target);
                if(!entries.ContainsKey(part))continue;
                if(type.EndsWith("/table",StringComparison.Ordinal)) {
                    var table=Xml(part);var range=(string)table.Root.Attribute("ref");table.Root.SetAttributeValue("ref",ShiftColumnAddress(range));
                    var filter=table.Root.Element(N+"autoFilter");if(filter!=null){ShiftFilterColumns(filter);filter.SetAttributeValue("ref",ShiftColumnAddress((string)filter.Attribute("ref")));}
                    int start=Column(Regex.Match(range??"",@"[A-Z]+").Value),offset=13-start;var columns=table.Root.Element(N+"tableColumns");
                    if(columns!=null&&offset>=0&&offset<=columns.Elements().Count()) {
                        var list=columns.Elements().ToList();var added=new XElement(N+"tableColumn",new XAttribute("id",0),new XAttribute("name","Номер тома"));list.Insert(offset,added);columns.RemoveNodes();
                        for(int i=0;i<list.Count;i++){list[i].SetAttributeValue("id",i+1);columns.Add(list[i]);}columns.SetAttributeValue("count",list.Count);
                    }
                    Store(part,table);
                } else if(type.EndsWith("/comments",StringComparison.Ordinal)) {
                    var comments=Xml(part);foreach(var c in comments.Descendants(N+"comment"))c.SetAttributeValue("ref",ShiftColumnAddress((string)c.Attribute("ref")));Store(part,comments);
                } else if(type.EndsWith("/vmlDrawing",StringComparison.Ordinal))entries[part]=ShiftVml(entries[part]);
            }
        }
    }
}
