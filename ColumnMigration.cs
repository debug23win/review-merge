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
        void InsertVolumeColumn(XDocument main,string name,int header,int first) {
            ExpandSharedFormulas(main);
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
            if(numbers!=null&&Text(Value(numbers.Elements(N+"c").FirstOrDefault(c=>(string)c.Attribute("r")=="A"+(header+1))))=="1")
                for(int c=13;c<=21;c++)SetCell(numbers,Cell(header+1,c,c,StyleAt(main,header+1,c,bodyStyle)));
            foreach(var row in rows){var cells=row.Elements(N+"c").OrderBy(c=>Column((string)c.Attribute("r"))).ToList();cells.Remove();row.Add(cells);}
            foreach(var sheet in book.Root.Element(N+"sheets").Elements()) {
                string path=PathFor((string)sheet.Attribute("name"));if(path==mainPath)continue;var doc=Xml(path);bool changed=false;
                foreach(var f in doc.Descendants(N+"f")){string next=ShiftMainFormula(f.Value,name,false);if(next!=f.Value){f.Value=next;changed=true;}}
                if(changed)Store(path,doc);
            }
            foreach(var defined in book.Descendants(N+"definedName"))defined.Value=ShiftMainFormula(defined.Value,name,false);
            // Native Excel table ranges and column definitions follow the inserted field.
            var sheetRels=mainPath.Replace("/sheet","/_rels/sheet")+".rels";
            if(entries.ContainsKey(sheetRels))foreach(var rel in Xml(sheetRels).Root.Elements()) {
                string type=(string)rel.Attribute("Type")??"",target=(string)rel.Attribute("Target")??"";
                string part=System.IO.Path.GetFullPath(System.IO.Path.Combine("C:/xlsx/",mainPath.Substring(0,mainPath.LastIndexOf('/')),target)).Replace('\\','/').Substring("C:/xlsx/".Length);
                if(!entries.ContainsKey(part))continue;
                if(type.EndsWith("/table",StringComparison.Ordinal)) {
                    var table=Xml(part);var range=(string)table.Root.Attribute("ref");table.Root.SetAttributeValue("ref",ShiftColumnAddress(range));
                    var filter=table.Root.Element(N+"autoFilter");if(filter!=null)filter.SetAttributeValue("ref",ShiftColumnAddress((string)filter.Attribute("ref")));
                    int start=Column(Regex.Match(range??"",@"[A-Z]+").Value),offset=13-start;var columns=table.Root.Element(N+"tableColumns");
                    if(columns!=null&&offset>=0&&offset<=columns.Elements().Count()) {
                        var list=columns.Elements().ToList();var added=new XElement(N+"tableColumn",new XAttribute("id",0),new XAttribute("name","Номер тома"));list.Insert(offset,added);columns.RemoveNodes();
                        for(int i=0;i<list.Count;i++){list[i].SetAttributeValue("id",i+1);columns.Add(list[i]);}columns.SetAttributeValue("count",list.Count);
                    }
                    Store(part,table);
                } else if(type.EndsWith("/comments",StringComparison.Ordinal)) {
                    var comments=Xml(part);foreach(var c in comments.Descendants(N+"comment"))c.SetAttributeValue("ref",ShiftColumnAddress((string)c.Attribute("ref")));Store(part,comments);
                }
            }
        }
    }
}
