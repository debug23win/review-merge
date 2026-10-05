using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace ReviewMerge {
    // Native XLSX packaging. Formulas are retained; cached values make the file readable before Excel recalculates it.
    public sealed class XlsxWriter {
        static readonly XNamespace N="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly XNamespace R="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static readonly XNamespace P="http://schemas.openxmlformats.org/package/2006/relationships";
        static readonly XNamespace C="http://schemas.openxmlformats.org/package/2006/content-types";
        static readonly CultureInfo Inv=CultureInfo.InvariantCulture;
        readonly Dictionary<string,byte[]> entries=new Dictionary<string,byte[]>(StringComparer.Ordinal);
        XDocument book,rels,types,styles;
        int bodyStyle,headStyle,dateStyle,dateHeadStyle,intStyle,percentStyle,titleStyle,noteStyle;
        string mainPath;
        static string Text(object v) {return XlsxReader.Text(v);}
        XDocument Xml(string path) {using(var m=new MemoryStream(entries[path]))using(var r=XmlReader.Create(m,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))return XDocument.Load(r);}
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
            public void Set(int r,int c,object v,int style,string f=null) {if(!Rows.ContainsKey(r))Rows[r]=new SortedDictionary<int,XElement>();Rows[r][c]=Cell(r,c,v,style,f);}
            public XElement Data() {return new XElement(N+"sheetData",Rows.Select(r=>new XElement(N+"row",new XAttribute("r",r.Key),new XAttribute("ht",Heights.ContainsKey(r.Key)?Heights[r.Key]:24),new XAttribute("customHeight",1),r.Value.Values)));}
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
            var helperHeader=existing[result.HeaderRow].Elements(N+"c").FirstOrDefault(e=>Column((string)e.Attribute("r"))==27);
            bool ourHelpers=helperHeader!=null&&string.Concat(helperHeader.Descendants(N+"t").Select(t=>t.Value))=="Первая дата из источников";
            if(!ourHelpers&&data.Elements(N+"row").Elements(N+"c").Any(e=>Column((string)e.Attribute("r"))>=21&&Column((string)e.Attribute("r"))<=27&&(e.Element(N+"v")!=null||e.Element(N+"f")!=null||e.Descendants(N+"t").Any())))
                throw new InvalidDataException("Столбцы U:AA сводного документа заняты. Они нужны для формул статистики; перенесите ваши данные в другие столбцы перед сборкой.");
            var sourceStyles=new Dictionary<int,int>();
            var sample=data.Elements(N+"row").FirstOrDefault(e=>(int)e.Attribute("r")==first);
            if(sample!=null)foreach(var c in sample.Elements(N+"c")){int col=Column((string)c.Attribute("r"));sourceStyles[col]=(int?)c.Attribute("s")??0;}
            var firstBoxes=new HashSet<string>();var firstBoxDates=new Dictionary<string,double>();
            var boxKeys=new List<string>();var dates=new List<double?>();var issues=new List<int>();var checkedFlags=new List<int>();
            foreach(var r in result.Rows){
                string kind=Text(r.Values[10]);double box;string key=(kind=="ПД"||kind=="ИИ"||kind=="ДПТ") && double.TryParse(Text(r.Values[11]),NumberStyles.Float,Inv,out box) && box>0 && box==Math.Floor(box) ? kind+"|"+box.ToString("0",Inv) : "";
                int checked_=Text(r.Values[8])!=""?1:0;double? date=checked_==1?XlsxReader.DateSerial(r.Values[9]):null;
                if(date.HasValue && r.FirstReviewDate.HasValue)date=Math.Min(date.Value,r.FirstReviewDate.Value);
                int issue=r.Values.Skip(12).Take(6).Any(v=>Text(v)!="" && Text(v)!="0") || Text(r.Values[18])!="" || Text(r.Values[19])!=""?1:0;
                boxKeys.Add(key);dates.Add(date);issues.Add(issue);checkedFlags.Add(checked_);
                if(key!="" && date.HasValue && (!firstBoxDates.ContainsKey(key) || firstBoxDates[key]>date.Value))firstBoxDates[key]=date.Value;
            }
            var usedBoxDate=new HashSet<string>();
            for(int i=0;i<result.Rows.Count;i++){
                token.ThrowIfCancellationRequested();int rn=rowNumbers[i];var row=result.Rows[i];XElement xmlrow;
                bool keep=existing.TryGetValue(rn,out xmlrow);if(!keep){xmlrow=new XElement(N+"row",new XAttribute("r",rn));existing[rn]=xmlrow;data.Add(xmlrow);}
                int noteLines=Math.Max(Text(row.Values[18]).Split('\n').Length,Text(row.Values[19]).Split('\n').Length);
                if(noteLines>1){xmlrow.SetAttributeValue("ht",Math.Min(120,Math.Max(30,noteLines*15)));xmlrow.SetAttributeValue("customHeight",1);}
                for(int col=keep?9:1;col<=20;col++){
                    var old=xmlrow.Elements(N+"c").FirstOrDefault(e=>Column((string)e.Attribute("r"))==col);
                    if(keep&&row.TargetRow!=null&&Text(row.TargetRow.Values[col-1])==Text(row.Values[col-1]))continue;
                    int style=old!=null?(int?)old.Attribute("s")??0:sourceStyles.ContainsKey(col)?sourceStyles[col]:bodyStyle;
                    if(old==null){if(col==10)style=dateStyle;if(col>=13&&col<=18)style=intStyle;if(col>=19)style=noteStyle;}
                    SetCell(xmlrow,Cell(rn,col,row.Values[col-1],style));
                }
                xmlrow.Elements(N+"c").Where(e=>Column((string)e.Attribute("r"))>=21&&Column((string)e.Attribute("r"))<=27).Remove();
                string k=boxKeys[i];int boxFirst=k!="" && firstBoxes.Add(k)?1:0;
                double? boxDate=k!="" && dates[i].HasValue && firstBoxDates[k]==dates[i].Value && usedBoxDate.Add(k)?dates[i]:null;
                string ur="$U$"+first+":$U$"+last,yr="$Y$"+first+":$Y$"+last;
                xmlrow.Add(Cell(rn,21,k,bodyStyle,"IF(AND(OR(K"+rn+"=\"ИИ\",K"+rn+"=\"ПД\",K"+rn+"=\"ДПТ\"),ISNUMBER(L"+rn+"),L"+rn+">0),K"+rn+"&\"|\"&TEXT(L"+rn+",\"0\"),\"\")"));
                xmlrow.Add(Cell(rn,22,boxFirst,intStyle,"IF(U"+rn+"=\"\",0,IF(COUNTIFS($U$"+first+":U"+rn+",U"+rn+")=1,1,0))"));
                xmlrow.Add(Cell(rn,23,boxDate.HasValue?(object)boxDate.Value:"",dateStyle,"IF(OR(U"+rn+"=\"\",Y"+rn+"=\"\"),\"\",IF(COUNTIFS("+ur+",U"+rn+","+yr+",\">0\","+yr+",\"<\"&Y"+rn+")=0,IF(COUNTIFS($U$"+first+":U"+rn+",U"+rn+",$Y$"+first+":Y"+rn+",Y"+rn+")=1,Y"+rn+",\"\"),\"\"))"));
                xmlrow.Add(Cell(rn,24,checkedFlags[i],intStyle,"IF(AND(C"+rn+"<>\"\",I"+rn+"<>\"\"),1,0)"));
                xmlrow.Add(Cell(rn,25,dates[i].HasValue?(object)dates[i].Value:"",dateStyle,"IF(AND(X"+rn+"=1,ISNUMBER(J"+rn+"),J"+rn+">0),IF(ISNUMBER(AA"+rn+"),MIN(INT(J"+rn+"),AA"+rn+"),INT(J"+rn+")),\"\")"));
                xmlrow.Add(Cell(rn,26,issues[i],intStyle,"IF(OR(COUNTIFS(M"+rn+":R"+rn+",1)>0,COUNTIFS(M"+rn+":R"+rn+",\"?*\")>0,S"+rn+"<>\"\",T"+rn+"<>\"\"),1,0)"));
                xmlrow.Add(Cell(rn,27,row.FirstReviewDate.HasValue?(object)row.FirstReviewDate.Value:null,dateStyle));
                var sorted=xmlrow.Elements(N+"c").OrderBy(e=>Column((string)e.Attribute("r"))).ToList();sorted.Remove();xmlrow.Add(sorted);
            }
            var header=data.Elements(N+"row").First(e=>(int)e.Attribute("r")==result.HeaderRow);string[] helpers={"Ключ короба (формула)","Первая строка короба","Первая проверка короба","Документ проверен","Дата для статистики","Есть замечания","Первая дата из источников"};
            header.Elements(N+"c").Where(e=>Column((string)e.Attribute("r"))>=21 && Column((string)e.Attribute("r"))<=27).Remove();for(int i=0;i<7;i++)header.Add(Cell(result.HeaderRow,21+i,helpers[i],headStyle));
            int maxCol=Math.Max(27,data.Elements(N+"row").Elements(N+"c").Select(e=>Column((string)e.Attribute("r"))).DefaultIfEmpty(27).Max());
            var dimension=main.Root.Element(N+"dimension");if(dimension==null){dimension=new XElement(N+"dimension");var sheetPr=main.Root.Element(N+"sheetPr");if(sheetPr!=null)sheetPr.AddAfterSelf(dimension);else main.Root.AddFirst(dimension);}dimension.SetAttributeValue("ref","A1:"+MergeEngine.ExcelColumn(maxCol)+last);
            var cols=main.Root.Element(N+"cols");if(cols==null){cols=new XElement(N+"cols");data.AddBeforeSelf(cols);}foreach(var col in cols.Elements().Where(e=>(int)e.Attribute("max")>=21&&(int)e.Attribute("min")<=27).ToList()){int min=(int)col.Attribute("min"),max=(int)col.Attribute("max");if(max>27){var rest=new XElement(col);rest.SetAttributeValue("min",28);cols.Add(rest);}if(min<21)col.SetAttributeValue("max",20);else col.Remove();}cols.Add(Col(21,27,24,true));var sortedCols=cols.Elements().OrderBy(e=>(int)e.Attribute("min")).ToList();sortedCols.Remove();cols.Add(sortedCols);
            var filter=main.Root.Element(N+"autoFilter");if(filter==null){filter=new XElement(N+"autoFilter");data.AddAfterSelf(filter);}filter.RemoveNodes();filter.SetAttributeValue("ref","A"+result.HeaderRow+":T"+last);
            Store(mainPath,main);progress(65,"Основной лист объединён. Создание формул статистики");
            RemoveGeneratedLogs();WriteSummary(actualName,result,asOf,first,last,boxKeys,dates,issues,checkedFlags,firstBoxDates);
            var calc=book.Root.Element(N+"calcPr");if(calc==null){calc=new XElement(N+"calcPr");book.Root.Add(calc);}calc.SetAttributeValue("calcMode","auto");calc.SetAttributeValue("fullCalcOnLoad",1);calc.SetAttributeValue("forceFullCalc",1);calc.SetAttributeValue("calcId",0);
            rels.Root.Elements().Where(e=>((string)e.Attribute("Type")??"").EndsWith("/calcChain")).Remove();types.Root.Elements().Where(e=>((string)e.Attribute("PartName")??"").EndsWith("/calcChain.xml")).Remove();entries.Remove("xl/calcChain.xml");
            int svIndex=book.Root.Element(N+"sheets").Elements().Select((e,i)=>new {e,i}).First(x=>(string)x.e.Attribute("name")=="Свод").i;
            var view=book.Root.Element(N+"bookViews");if(view!=null && view.Elements().Any())view.Elements().First().SetAttributeValue("activeTab",svIndex);
            Store("xl/workbook.xml",book);Store("xl/_rels/workbook.xml.rels",rels);Store("[Content_Types].xml",types);Store("xl/styles.xml",styles);
            progress(90,"Сохранение XLSX с формулами");using(var s=new FileStream(output,FileMode.CreateNew))using(var z=new ZipArchive(s,ZipArchiveMode.Create)){foreach(var item in entries){token.ThrowIfCancellationRequested();var e=z.CreateEntry(item.Key,CompressionLevel.Optimal);using(var stream=e.Open())stream.Write(item.Value,0,item.Value.Length);}}
            result.Boxes=firstBoxDates.Count(e=>e.Value<=asOf.Date.ToOADate());
        }
        static int Column(string s){int n=0;foreach(char c in s){if(c<'A'||c>'Z')break;n=n*26+c-'A'+1;}return n;}
        void WriteSummary(string name,MergeResult result,DateTime asOf,int first,int last,List<string> keys,List<double?> dates,List<int> issues,List<int> checked_,Dictionary<string,double> boxes) {
            var g=new Grid();int end=result.Dates.Count+1;string ec=MergeEngine.ExcelColumn(end);
            string cr=Ref(name,"C",first,last),ir=Ref(name,"I",first,last),kr=Ref(name,"K",first,last),vr=Ref(name,"V",first,last),wr=Ref(name,"W",first,last),xr=Ref(name,"X",first,last),yr=Ref(name,"Y",first,last),zr=Ref(name,"Z",first,last);
            g.Set(2,1,"Свод результатов проверки",titleStyle);g.Heights[2]=32;
            g.Set(3,1,"Накопительный итог и прирост — по первой проверке. Короб учтён при наличии проверенного документа.",bodyStyle);
            g.Set(5,1,"Показатель",headStyle);g.Heights[5]=32;
            string[] labels={"Документов в реестре","Документов проверено на дату","Документов без отметки проверки на дату","Проверенных документов с замечаниями","Проверенных документов без замечаний","Доля проверенных документов","Документов впервые проверено за сутки","Коробов указано в реестре","Коробов с проверенными документами на дату","Коробов впервые отмечено за сутки"};
            for(int i=0;i<10;i++)g.Set(6+i,1,labels[i],noteStyle);g.Heights[8]=32;
            int totalBoxes=keys.Where(k=>k!="").Distinct().Count();
            for(int i=0;i<result.Dates.Count;i++){
                int c=i+2;string cs=MergeEngine.ExcelColumn(c),d=cs+"$5";double serial=result.Dates[i].ToOADate();
                int reviewed=dates.Count(x=>x.HasValue&&x.Value<=serial),remark=dates.Select((x,j)=>new{x,j}).Count(v=>v.x.HasValue&&v.x.Value<=serial&&issues[v.j]==1);
                g.Set(5,c,serial,dateHeadStyle);
                g.Set(6,c,result.Documents,intStyle,"COUNTA("+cr+")");g.Set(7,c,reviewed,intStyle,"COUNTIFS("+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+d+")");
                g.Set(8,c,result.Documents-reviewed,intStyle,cs+"6-"+cs+"7");g.Set(9,c,remark,intStyle,"COUNTIFS("+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+d+","+zr+",1)");g.Set(10,c,reviewed-remark,intStyle,cs+"7-"+cs+"9");
                g.Set(11,c,(double)reviewed/result.Documents,percentStyle,"IF("+cs+"6=0,0,"+cs+"7/"+cs+"6)");g.Set(12,c,dates.Count(x=>x==serial),intStyle,"COUNTIFS("+xr+",1,"+yr+","+d+")");g.Set(13,c,totalBoxes,intStyle,"SUM("+vr+")");
                g.Set(14,c,boxes.Count(x=>x.Value<=serial),intStyle,"COUNTIFS("+wr+",\">0\","+wr+",\"<=\"&"+d+")");g.Set(15,c,boxes.Count(x=>x.Value==serial),intStyle,"COUNTIFS("+wr+","+d+")");
            }
            g.Set(18,1,"Проверено без корректной даты (не входит в статистику по дням)",noteStyle);g.Heights[18]=32;g.Set(18,2,dates.Select((d,i)=>new{d,i}).Count(x=>checked_[x.i]==1&&!x.d.HasValue),intStyle,"SUMPRODUCT(("+xr+"=1)*("+yr+"=\"\"))");
            g.Set(19,1,"Документы с замечаниями без проверяющего",noteStyle);g.Set(19,2,issues.Select((v,i)=>new{v,i}).Count(x=>x.v==1&&checked_[x.i]==0),intStyle,"COUNTIFS("+xr+",0,"+zr+",1)");
            g.Set(20,1,"Проверки с датой позже даты свода",noteStyle);g.Set(20,2,dates.Count(d=>d.HasValue&&d.Value>asOf.ToOADate()),intStyle,"COUNTIFS("+xr+",1,"+yr+",\">\"&"+ec+"$5)");
            string[] heads={"Вид документации","В реестре","Проверено на дату","С замечаниями","Осталось"};for(int j=0;j<5;j++)g.Set(22,j+1,heads[j],headStyle);g.Heights[22]=32;
            int sumTotal=0,sumCheck=0,sumIssue=0;string[] kinds={"ИИ","ПД","ДПТ"};
            for(int k=0;k<4;k++){
                int row=23+k;string kind=k<3?kinds[k]:"Не указан / иной";g.Set(row,1,kind,bodyStyle);
                int tot=0,count=0,issue=0;for(int i=0;i<result.Rows.Count;i++)if(k<3?Text(result.Rows[i].Values[10])==kind:!kinds.Contains(Text(result.Rows[i].Values[10]))){tot++;if(dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()){count++;issue+=issues[i];}}
                g.Set(row,2,tot,intStyle,k<3?"COUNTIFS("+kr+",$A"+row+")":ec+"6-SUM(B23:B25)");
                g.Set(row,3,count,intStyle,k<3?"COUNTIFS("+kr+",$A"+row+","+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+ec+"$5)":ec+"7-SUM(C23:C25)");
                g.Set(row,4,issue,intStyle,k<3?"COUNTIFS("+kr+",$A"+row+","+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+ec+"$5,"+zr+",1)":ec+"9-SUM(D23:D25)");g.Set(row,5,tot-count,intStyle,"B"+row+"-C"+row);sumTotal+=tot;sumCheck+=count;sumIssue+=issue;
            }
            string[] rh={"Последний проверяющий","Документов на дату","С замечаниями","Без даты"};for(int j=0;j<4;j++)g.Set(29,j+1,rh[j],headStyle);g.Heights[29]=32;
            var reviewers=result.Rows.Select(r=>Text(r.Values[8])).Where(v=>v!="").Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v=>v,StringComparer.CurrentCulture).ToList();
            for(int n=0;n<reviewers.Count;n++){
                int row=30+n;string who=reviewers[n];g.Set(row,1,who,bodyStyle);string crit="SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($A"+row+",\"~\",\"~~\"),\"*\",\"~*\"),\"?\",\"~?\")";
                int count=0,issue=0,missing=0;for(int i=0;i<result.Rows.Count;i++)if(string.Equals(Text(result.Rows[i].Values[8]),who,StringComparison.OrdinalIgnoreCase)){if(!dates[i].HasValue)missing++;else if(dates[i].Value<=asOf.ToOADate()){count++;issue+=issues[i];}}
                g.Set(row,2,count,intStyle,"COUNTIFS("+ir+","+crit+","+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+ec+"$5)");g.Set(row,3,issue,intStyle,"COUNTIFS("+ir+","+crit+","+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+ec+"$5,"+zr+",1)");g.Set(row,4,missing,intStyle,"SUMPRODUCT(("+ir+"=$A"+row+")*("+xr+"=1)*("+yr+"=\"\"))");
            }
            PutSheet("Свод",Sheet(g,new []{Col(1,1,64),Col(2,Math.Max(5,end),19)},true,Math.Max(5,end)));
        }
        void WriteLogs(MergeResult result) {
            var h=new Grid();string[] names={"Исходный файл","Полный путь","Строка источника","№ п/п","Раздел","Документ (файл)","Формат","Вес","Дата выгрузки","Время выгрузки","CRC32","Проверил","Дата проверки","Вид","Короб","Название","Шифр","Страницы","Титул: подписи/печати","CRC32: замечание","ИУЛ: подписи","Ошибки в описи","Примечание"};for(int j=0;j<names.Length;j++)h.Set(1,j+1,names[j],headStyle);h.Heights[1]=32;
            for(int i=0;i<result.History.Count;i++){var r=result.History[i];int rn=i+2;h.Set(rn,1,Path.GetFileName(r.File),bodyStyle);h.Set(rn,2,r.File,bodyStyle);h.Set(rn,3,r.Row,intStyle);for(int j=0;j<20;j++)h.Set(rn,j+4,r.Values[j],j==9?dateStyle:bodyStyle);}
            PutSheet("История проверок",Sheet(h,new []{Col(1,1,30),Col(2,2,55),Col(3,5,15),Col(6,6,60),Col(7,21,18),Col(22,23,75)},false,23));
            var c=new Grid();string[] cn={"Документ","Поле","Различия в источниках","Применённое правило"};for(int j=0;j<4;j++)c.Set(1,j+1,cn[j],headStyle);c.Heights[1]=32;
            if(result.Conflicts.Count==0)c.Set(2,1,"Конфликтов не обнаружено",bodyStyle);
            for(int i=0;i<result.Conflicts.Count;i++){var cf=result.Conflicts[i];int rn=i+2;c.Set(rn,1,cf.Document,noteStyle);c.Set(rn,2,cf.Field,noteStyle);c.Set(rn,3,cf.Details,noteStyle);c.Set(rn,4,cf.Decision,noteStyle);c.Heights[rn]=Math.Min(180,Math.Max(45,cf.Details.Split('\n').Length*16));}
            PutSheet("Конфликты свода",Sheet(c,new []{Col(1,1,55),Col(2,2,30),Col(3,3,85),Col(4,4,75)},false,4));
            var p=new Grid();string[] pn={"Исходный файл","Строка","Документ","Проблема"};for(int j=0;j<4;j++)p.Set(1,j+1,pn[j],headStyle);p.Heights[1]=32;
            if(result.Problems.Count==0)p.Set(2,1,"Проблем не обнаружено",bodyStyle);
            for(int i=0;i<result.Problems.Count;i++){var pr=result.Problems[i];int rn=i+2;p.Set(rn,1,Path.GetFileName(pr.File),bodyStyle);p.Set(rn,2,pr.Row,intStyle);p.Set(rn,3,pr.Document,noteStyle);p.Set(rn,4,pr.Detail,noteStyle);p.Heights[rn]=45;}
            PutSheet("Проблемы данных",Sheet(p,new []{Col(1,1,32),Col(2,2,10),Col(3,3,55),Col(4,4,95)},false,4));
        }
    }
}
