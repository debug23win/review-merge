using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace ReviewMerge {
    // Native XLSX packaging. Formulas are retained; cached values make the file readable before Excel recalculates it.
    public sealed partial class XlsxWriter {
        static readonly XNamespace N="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly XNamespace R="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static readonly XNamespace P="http://schemas.openxmlformats.org/package/2006/relationships";
        static readonly XNamespace C="http://schemas.openxmlformats.org/package/2006/content-types";
        static readonly CultureInfo Inv=CultureInfo.InvariantCulture;
        readonly Dictionary<string,byte[]> entries=new Dictionary<string,byte[]>(StringComparer.Ordinal);
        readonly MergeOptions options;
        public XlsxWriter(MergeOptions options=null){this.options=options??new MergeOptions();}
        XDocument book,rels,types,styles;
        int bodyStyle,headStyle,dateStyle,dateHeadStyle,intStyle,percentStyle,titleStyle,noteStyle;
        readonly Dictionary<int,int> mainDateStyles=new Dictionary<int,int>();
        int MainDateStyle(int original) {
            int id;if(mainDateStyles.TryGetValue(original,out id))return id;
            var xfs=styles.Root.Element(N+"cellXfs");var source=xfs.Elements().ElementAt(original);var xf=new XElement(source);xf.SetAttributeValue("numFmtId",(int)xfs.Elements().ElementAt(dateStyle).Attribute("numFmtId"));xf.SetAttributeValue("applyNumberFormat",1);id=xfs.Elements().Count();xfs.Add(xf);xfs.SetAttributeValue("count",id+1);mainDateStyles[original]=id;return id;
        }
        string mainPath;
        static string Text(object v) {return XlsxReader.Text(v);}
        XDocument Xml(string path) {
            XDocument doc;using(var m=new MemoryStream(entries[path]))using(var r=XmlReader.Create(m,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))doc=XDocument.Load(r);
            if(path.StartsWith("xl/worksheets/",StringComparison.Ordinal))Rules.NumberCells(doc);
            return doc;
        }
        void Store(string path,XDocument xml) {using(var m=new MemoryStream()){xml.Save(m);entries[path]=m.ToArray();}}
        int AddStyle(int font,int fill,int num,bool wrap) {
            var xfs=styles.Root.Element(N+"cellXfs");int id=xfs.Elements().Count();
            xfs.Add(new XElement(N+"xf",new XAttribute("numFmtId",num),new XAttribute("fontId",font),new XAttribute("fillId",fill),new XAttribute("borderId",0),new XAttribute("xfId",0),new XAttribute("applyNumberFormat",1),new XAttribute("applyAlignment",1),new XElement(N+"alignment",new XAttribute("vertical","center"),new XAttribute("wrapText",wrap?1:0))));
            xfs.SetAttributeValue("count",id+1);return id;
        }
        void SetupStyles() {
            styles=Xml("xl/styles.xml");var fonts=styles.Root.Element(N+"fonts");var fills=styles.Root.Element(N+"fills");
            int normal=fonts.Elements().Count();fonts.Add(new XElement(N+"font",new XElement(N+"sz",new XAttribute("val",10)),new XElement(N+"name",new XAttribute("val","Arial")),new XElement(N+"color",new XAttribute("rgb","FF1C2E41"))));
            int bold=normal+1;fonts.Add(new XElement(N+"font",new XElement(N+"b"),new XElement(N+"sz",new XAttribute("val",10)),new XElement(N+"name",new XAttribute("val","Arial")),new XElement(N+"color",new XAttribute("rgb","FFFFFFFF"))));
            int title=normal+2;fonts.Add(new XElement(N+"font",new XElement(N+"b"),new XElement(N+"sz",new XAttribute("val",16)),new XElement(N+"name",new XAttribute("val","Arial"))));fonts.SetAttributeValue("count",normal+3);
            int dark=fills.Elements().Count();fills.Add(new XElement(N+"fill",new XElement(N+"patternFill",new XAttribute("patternType","solid"),new XElement(N+"fgColor",new XAttribute("rgb","FF234564")),new XElement(N+"bgColor",new XAttribute("indexed",64)))));fills.SetAttributeValue("count",dark+1);
            var formats=styles.Root.Element(N+"numFmts");if(formats==null){formats=new XElement(N+"numFmts");styles.Root.AddFirst(formats);}
            int fmt=Math.Max(164,formats.Elements().Select(e=>(int)e.Attribute("numFmtId")).DefaultIfEmpty(163).Max()+1);
            formats.Add(new XElement(N+"numFmt",new XAttribute("numFmtId",fmt),new XAttribute("formatCode","dd.mm.yyyy")));formats.SetAttributeValue("count",formats.Elements().Count());
            bodyStyle=AddStyle(normal,0,0,false);noteStyle=AddStyle(normal,0,0,true);headStyle=AddStyle(bold,dark,0,true);dateStyle=AddStyle(normal,0,fmt,false);dateHeadStyle=AddStyle(bold,dark,fmt,false);intStyle=AddStyle(normal,0,1,false);percentStyle=AddStyle(normal,0,10,false);titleStyle=AddStyle(title,0,0,false);
        }
        static XElement Cell(int row,int col,object value,int style,string formula=null) {
            var cell=new XElement(N+"c",new XAttribute("r",MergeEngine.ExcelColumn(col)+row),new XAttribute("s",style));
            if(formula!=null) {
                cell.Add(new XElement(N+"f",formula.TrimStart('=')));
                if(value is string || value==null)cell.SetAttributeValue("t","str");
                cell.Add(new XElement(N+"v",value==null?"":Convert.ToString(value,Inv)));return cell;
            }
            if(value==null)return cell;
            if(value is double || value is int || value is long || value is decimal)cell.Add(new XElement(N+"v",Convert.ToString(value,Inv)));
            else {cell.SetAttributeValue("t","inlineStr");cell.Add(new XElement(N+"is",new XElement(N+"t",new XAttribute(XNamespace.Xml+"space","preserve"),Convert.ToString(value,Inv))));}
            return cell;
        }
        sealed class Grid {
            public readonly SortedDictionary<int,SortedDictionary<int,XElement>> Rows=new SortedDictionary<int,SortedDictionary<int,XElement>>();
            public readonly Dictionary<int,double> Heights=new Dictionary<int,double>();
            public readonly Dictionary<int,XElement> Templates=new Dictionary<int,XElement>();
            public void Set(int r,int c,object v,int style,string f=null) {if(!Rows.ContainsKey(r))Rows[r]=new SortedDictionary<int,XElement>();Rows[r][c]=Cell(r,c,v,style,f);}
            public XElement Data() {return new XElement(N+"sheetData",Rows.Select(r=>{XElement template;bool exists=Templates.TryGetValue(r.Key,out template);var row=exists?new XElement(N+"row",template.Attributes()):new XElement(N+"row",new XAttribute("r",r.Key));if(Heights.ContainsKey(r.Key)||!exists){row.SetAttributeValue("ht",Heights.ContainsKey(r.Key)?Heights[r.Key]:24);row.SetAttributeValue("customHeight",1);}row.Add(r.Value.Values);return row;}));}
        }
        static XElement Col(int from,int to,double width,bool hidden=false){return new XElement(N+"col",new XAttribute("min",from),new XAttribute("max",to),new XAttribute("width",width),new XAttribute("customWidth",1),new XAttribute("hidden",hidden?1:0));}
        XDocument Sheet(Grid grid,IEnumerable<XElement> columns,bool freeze,int lastCol) {
            int last=grid.Rows.Keys.DefaultIfEmpty(1).Max();
            return new XDocument(new XDeclaration("1.0","utf-8","yes"),new XElement(N+"worksheet",
              new XElement(N+"sheetPr",new XElement(N+"tabColor",new XAttribute("rgb","FF234564"))),
              new XElement(N+"dimension",new XAttribute("ref","A1:"+MergeEngine.ExcelColumn(lastCol)+last)),
              new XElement(N+"sheetViews",new XElement(N+"sheetView",new XAttribute("workbookViewId",0),new XAttribute("showGridLines",0),freeze?new XElement(N+"pane",new XAttribute("ySplit",5),new XAttribute("xSplit",1),new XAttribute("topLeftCell","B6"),new XAttribute("activePane","bottomRight"),new XAttribute("state","frozen")):null)),
              new XElement(N+"sheetFormatPr",new XAttribute("defaultRowHeight",24)),new XElement(N+"cols",columns),grid.Data()));
        }
        string PathFor(string sheetName) {
            var sh=book.Root.Element(N+"sheets").Elements().FirstOrDefault(e=>string.Equals(((string)e.Attribute("name")).Trim(),sheetName.Trim(),StringComparison.OrdinalIgnoreCase));
            if(sh==null)return null;string id=(string)sh.Attribute(R+"id");string p=(string)rels.Root.Elements().First(e=>(string)e.Attribute("Id")==id).Attribute("Target");return p.StartsWith("/")?p.TrimStart('/') : "xl/"+p;
        }
        void PutSheet(string name,XDocument data) {
            string path=PathFor(name);
            if(path!=null)book.Root.Element(N+"sheets").Elements().First(e=>PathFor((string)e.Attribute("name"))==path).SetAttributeValue("name",name);
            if(path==null){
                int n=1;while(entries.ContainsKey("xl/worksheets/sheet"+n+".xml"))n++;
                path="xl/worksheets/sheet"+n+".xml";string rid="mergeSheet"+n;
                while(rels.Root.Elements().Any(e=>(string)e.Attribute("Id")==rid))rid+="x";
                int sid=book.Root.Element(N+"sheets").Elements().Select(e=>(int)e.Attribute("sheetId")).DefaultIfEmpty(0).Max()+1;
                book.Root.Element(N+"sheets").Add(new XElement(N+"sheet",new XAttribute("name",name),new XAttribute("sheetId",sid),new XAttribute(R+"id",rid)));
                rels.Root.Add(new XElement(P+"Relationship",new XAttribute("Id",rid),new XAttribute("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),new XAttribute("Target","worksheets/sheet"+n+".xml")));
                types.Root.Add(new XElement(C+"Override",new XAttribute("PartName","/"+path),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            }
            Store(path,data);
        }
        void RemoveGeneratedLogs() {
            foreach(var sheet in book.Root.Element(N+"sheets").Elements().ToList()) {
                string name=(string)sheet.Attribute("name"),path=PathFor(name);string[] expected=null;
                if(name=="История проверок")expected=new[]{"Исходный файл","Полный путь"};
                if(name=="Конфликты свода")expected=new[]{"Документ","Поле","Различия в источниках","Применённое правило"};
                if(name=="Проблемы данных")expected=new[]{"Исходный файл","Строка","Документ","Проблема"};
                if(expected==null)continue;
                var first=Xml(path).Descendants(N+"sheetData").Elements(N+"row").FirstOrDefault();
                var headers=first==null?new string[0]:first.Elements(N+"c").Take(expected.Length).Select(e=>string.Concat(e.Descendants(N+"t").Select(t=>t.Value))).ToArray();
                if(!headers.SequenceEqual(expected))continue;
                int index=book.Root.Element(N+"sheets").Elements().TakeWhile(e=>e!=sheet).Count();
                foreach(var n in book.Descendants(N+"definedName").Where(e=>e.Attribute("localSheetId")!=null).ToList()) {
                    int local=(int)n.Attribute("localSheetId");if(local==index)n.Remove();else if(local>index)n.SetAttributeValue("localSheetId",local-1);
                }
                string id=(string)sheet.Attribute(R+"id");sheet.Remove();rels.Root.Elements().Where(e=>(string)e.Attribute("Id")==id).Remove();
                types.Root.Elements().Where(e=>(string)e.Attribute("PartName")=="/"+path).Remove();entries.Remove(path);
                entries.Remove(Path.GetDirectoryName(path).Replace('\\','/')+"/_rels/"+Path.GetFileName(path)+".rels");
            }
        }
        static void SetCell(XElement row,XElement value) {
            var old=row.Elements(N+"c").FirstOrDefault(e=>(string)e.Attribute("r")== (string)value.Attribute("r"));
            if(old!=null)old.ReplaceWith(value);else row.Add(value);
        }
        string Ref(string name,string col,int start,int last){return "'"+name.Replace("'","''")+"'!$"+col+"$"+start+":$"+col+"$"+last;}
        public void Save(string firstFile,string output,string mainName,DateTime asOf,MergeResult result,Action<int,string> progress,CancellationToken token) {
            using(var s=File.OpenRead(firstFile))using(var z=new ZipArchive(s,ZipArchiveMode.Read))foreach(var e in z.Entries){using(var m=new MemoryStream())using(var es=e.Open()){es.CopyTo(m);entries[e.FullName]=m.ToArray();}}
            book=Xml("xl/workbook.xml");rels=Xml("xl/_rels/workbook.xml.rels");types=Xml("[Content_Types].xml");
            var pr=book.Root.Element(N+"workbookPr");if(pr!=null && ((string)pr.Attribute("date1904")=="1" || (string)pr.Attribute("date1904")=="true"))throw new InvalidDataException("Первый файл использует систему дат 1904. Выберите первым файл с системой дат 1900.");
            SetupStyles();mainPath=PathFor(mainName);var main=Xml(mainPath);string actualName=(string)book.Root.Element(N+"sheets").Elements().First(e=>PathFor((string)e.Attribute("name"))==mainPath).Attribute("name");
            int first=result.StartRow;var data=main.Root.Element(N+"sheetData");
            var existing=data.Elements(N+"row").ToDictionary(e=>(int)e.Attribute("r"));
            int next=Math.Max(first-1,existing.Keys.DefaultIfEmpty(first-1).Max())+1;
            var rowNumbers=result.Rows.Select(r=>r.TargetRow!=null?r.TargetRow.Row:next++).ToList();
            int last=Math.Max(existing.Keys.DefaultIfEmpty(first).Max(),rowNumbers.DefaultIfEmpty(first).Max());
            if(last>1048576)throw new InvalidDataException("В сводном документе недостаточно свободных строк Excel для новых документов.");
            string[] legacyHeaders={"Ключ короба (формула)","Первая строка короба","Первая проверка короба","Документ проверен","Дата для статистики","Есть замечания","Первая дата из источников"};
            bool ourHelpers=Enumerable.Range(0,7).All(i=>Text(Value(At(main,result.HeaderRow,21+i)))==legacyHeaders[i]);
            if(ourHelpers) RemoveLegacyHelpers(main);
            var sourceStyles=new Dictionary<int,int>();
            var sample=data.Elements(N+"row").FirstOrDefault(e=>(int)e.Attribute("r")==first);
            if(sample!=null)foreach(var c in sample.Elements(N+"c")){int col=Column((string)c.Attribute("r"));sourceStyles[col]=(int?)c.Attribute("s")??0;}
            var boxKeys=new List<List<string>>();var dates=new List<double?>();var issues=new List<int>();var checkedFlags=new List<int>();
            foreach(var r in result.Rows){
                int checked_=Text(r.Values[8])!=""?1:0;double? date=checked_==1?XlsxReader.DateSerial(r.Values[9]):null;
                if(date.HasValue && r.FirstReviewDate.HasValue)date=Math.Min(date.Value,r.FirstReviewDate.Value);
                int issue=r.Values.Skip(12).Take(6).Any(v=>Text(v)!="" && Text(v)!="0") || Text(r.Values[18])!="" || Text(r.Values[19])!=""?1:0;
                boxKeys.Add(Rules.BoxKeys(r.Values[10],r.Values[11]));dates.Add(date);issues.Add(issue);checkedFlags.Add(checked_);
            }
            for(int i=0;i<result.Rows.Count;i++){
                token.ThrowIfCancellationRequested();int rn=rowNumbers[i];var row=result.Rows[i];XElement xmlrow;
                bool keep=existing.TryGetValue(rn,out xmlrow);if(!keep){xmlrow=new XElement(N+"row",new XAttribute("r",rn));existing[rn]=xmlrow;data.Add(xmlrow);}
                WriteMainRow(xmlrow,rn,keep,row,row.TargetRow,sourceStyles);
                // A document repeated in the consolidated book gets the same review data in every repeated row.
                foreach(var copy in row.TargetDuplicates){XElement other;if(existing.TryGetValue(copy.Row,out other))WriteMainRow(other,copy.Row,true,row,copy,sourceStyles);}
            }
            int maxCol=Math.Max(20,data.Elements(N+"row").Elements(N+"c").Select(e=>Column((string)e.Attribute("r"))).DefaultIfEmpty(20).Max());
            var dimension=main.Root.Element(N+"dimension");if(dimension==null){dimension=new XElement(N+"dimension");var sheetPr=main.Root.Element(N+"sheetPr");if(sheetPr!=null)sheetPr.AddAfterSelf(dimension);else main.Root.AddFirst(dimension);}dimension.SetAttributeValue("ref","A1:"+MergeEngine.ExcelColumn(maxCol)+last);
            var filter=main.Root.Element(N+"autoFilter");
            if(filter!=null){string right=Regex.Match((string)filter.Attribute("ref")??"",@":\$?([A-Z]+)").Groups[1].Value;filter.SetAttributeValue("ref","A"+result.HeaderRow+":"+MergeEngine.ExcelColumn(Math.Max(20,Column(right)))+last);}
            Store(mainPath,main);progress(65,"Данные проверки объединены. Статистика рассчитывается на листе «Свод»");
            RemoveGeneratedLogs();WriteCurrentSummary(actualName,result,asOf,rowNumbers,boxKeys,dates,issues,checkedFlags);
            Finish(output,result,asOf,progress,token);
        }
        void WriteMainRow(XElement xmlrow,int rn,bool keep,MergedRow row,SourceRow previous,Dictionary<int,int> sourceStyles) {
            int noteLines=Math.Max(Text(row.Values[18]).Split('\n').Length,Text(row.Values[19]).Split('\n').Length);
            bool notesChanged=!keep||previous==null||Text(previous.Values[18])!=Text(row.Values[18])||Text(previous.Values[19])!=Text(row.Values[19]);
            if(noteLines>1&&notesChanged){xmlrow.SetAttributeValue("ht",Math.Min(120,Math.Max(30,noteLines*15)));xmlrow.SetAttributeValue("customHeight",1);}
            for(int col=keep?9:1;col<=20;col++){
                var old=xmlrow.Elements(N+"c").FirstOrDefault(e=>Column((string)e.Attribute("r"))==col);
                if(col!=10&&keep&&previous!=null&&Text(previous.Values[col-1])==Text(row.Values[col-1]))continue;
                int style=old!=null?(int?)old.Attribute("s")??0:sourceStyles.ContainsKey(col)?sourceStyles[col]:bodyStyle;
                if(old==null){if(col==10)style=dateStyle;if(col>=13&&col<=18)style=intStyle;if(col>=19)style=noteStyle;}
                object value=row.Values[col-1];if(col==10){style=MainDateStyle(style);if(old!=null&&old.Element(N+"f")!=null&&previous!=null&&Text(previous.Values[col-1])==Text(value)){old.SetAttributeValue("s",style);continue;}double? serial=XlsxReader.DateSerial(value);if(serial.HasValue)value=serial.Value;}
                SetCell(xmlrow,Cell(rn,col,value,style));
            }
            /* Calculation cells belong on the reference sheet, outside the main register. */
            var sorted=xmlrow.Elements(N+"c").OrderBy(e=>Column((string)e.Attribute("r"))).ToList();sorted.Remove();xmlrow.Add(sorted);
        }
        void Finish(string output,MergeResult result,DateTime asOf,Action<int,string> progress,CancellationToken token) {
            // Excel recalculates on open; a stored forced full calculation would recalculate the whole book after every edit.
            var calc=book.Root.Element(N+"calcPr");if(calc==null){calc=new XElement(N+"calcPr");book.Root.Add(calc);}calc.SetAttributeValue("calcMode","auto");calc.SetAttributeValue("fullCalcOnLoad",1);calc.SetAttributeValue("forceFullCalc",null);calc.SetAttributeValue("calcId",0);
            rels.Root.Elements().Where(e=>((string)e.Attribute("Type")??"").EndsWith("/calcChain")).Remove();types.Root.Elements().Where(e=>((string)e.Attribute("PartName")??"").EndsWith("/calcChain.xml")).Remove();entries.Remove("xl/calcChain.xml");
            int svIndex=book.Root.Element(N+"sheets").Elements().Select((e,i)=>new {e,i}).First(x=>(string)x.e.Attribute("name")=="Свод").i;
            var view=book.Root.Element(N+"bookViews");if(view==null){view=new XElement(N+"bookViews");book.Root.Element(N+"sheets").AddBeforeSelf(view);}if(!view.Elements().Any())view.Add(new XElement(N+"workbookView"));foreach(var workbookView in view.Elements())workbookView.SetAttributeValue("activeTab",svIndex);
            foreach(var sheet in book.Root.Element(N+"sheets").Elements()) {
                string sheetName=(string)sheet.Attribute("name"),path=PathFor(sheetName);var doc=Xml(path);
                bool changed=false;
                foreach(var sheetView in doc.Descendants(N+"sheetView")) {
                    string selected=sheetName=="Свод"?"1":null;
                    if((string)sheetView.Attribute("tabSelected")!=selected){sheetView.SetAttributeValue("tabSelected",selected);changed=true;}
                }
                if(changed)Store(path,doc);
            }
            Store("xl/workbook.xml",book);Store("xl/_rels/workbook.xml.rels",rels);Store("[Content_Types].xml",types);Store("xl/styles.xml",styles);
            progress(90,"Сохранение XLSX с формулами");using(var s=new FileStream(output,FileMode.CreateNew))using(var z=new ZipArchive(s,ZipArchiveMode.Create)){foreach(var item in entries){token.ThrowIfCancellationRequested();var e=z.CreateEntry(item.Key,CompressionLevel.Optimal);using(var stream=e.Open())stream.Write(item.Value,0,item.Value.Length);}}
            result.Boxes=referenceLayout.Completed.Count(e=>e.Value.HasValue&&e.Value.Value<=asOf.Date.ToOADate());
        }
        static int Column(string s){int n=0;foreach(char c in s){if(c<'A'||c>'Z')break;n=n*26+c-'A'+1;}return n;}
    }
}
