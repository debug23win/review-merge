using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ReviewMerge {
    public sealed partial class XlsxWriter {
        sealed class ReportAddress {public string Sheet;public int Row,Column;public string Address{get{return MergeEngine.ExcelColumn(Column)+Row;}}}
        readonly Dictionary<string,ReportAddress> reportAddresses=new Dictionary<string,ReportAddress>();
        readonly Dictionary<string,int> reportStyleCache=new Dictionary<string,int>();
        static int CellStyle(XElement cell,int fallback){return cell==null?fallback:(int?)cell.Attribute("s")??fallback;}
        static void ExpandSharedFormulas(XDocument doc) {
            var cells=doc.Descendants(N+"c").ToList();var anchors=new Dictionary<string,Tuple<string,string>>();
            foreach(var c in cells){var f=c.Element(N+"f");if(f!=null&&(string)f.Attribute("t")=="shared"&&f.Value!="")anchors[(string)f.Attribute("si")]=Tuple.Create((string)c.Attribute("r"),f.Value);}
            foreach(var c in cells){var f=c.Element(N+"f");if(f==null||(string)f.Attribute("t")!="shared")continue;Tuple<string,string> anchor;
                if(!anchors.TryGetValue((string)f.Attribute("si"),out anchor))throw new System.IO.InvalidDataException("Не найден исходный текст общей формулы Excel: "+(string)c.Attribute("r"));
                string target=(string)c.Attribute("r");int dc=Column(target)-Column(anchor.Item1),dr=int.Parse(Regex.Match(target,@"\d+$").Value,Inv)-int.Parse(Regex.Match(anchor.Item1,@"\d+$").Value,Inv);
                f.Value=Regex.Replace(anchor.Item2,@"""(?:[^""]|"""")*""|(?<![\p{L}\p{N}_])(?<ca>\$?)(?<c>[A-Z]{1,3})(?<ra>\$?)(?<r>\d+)(?![\p{L}\p{N}_.]|\s*\()",m=>{
                    if(!m.Groups["c"].Success)return m.Value;int col=Column(m.Groups["c"].Value)+(m.Groups["ca"].Value==""?dc:0),row=int.Parse(m.Groups["r"].Value,Inv)+(m.Groups["ra"].Value==""?dr:0);
                    return col<1||row<1?"#REF!":m.Groups["ca"].Value+MergeEngine.ExcelColumn(col)+m.Groups["ra"].Value+row;
                });f.RemoveAttributes();
            }
        }
        int CaptionRow(XDocument doc,string text) {
            return doc.Root.Element(N+"sheetData").Elements(N+"row").Where(r=>XlsxReader.Normal(Value(r.Elements(N+"c").FirstOrDefault(c=>Column((string)c.Attribute("r"))==1))).Contains(text))
                .Select(r=>(int)r.Attribute("r")).FirstOrDefault();
        }
        int ReportStyle(int original,string left=null,string right=null,string top=null,string bottom=null,string color=null,bool? bold=null,int? number=null,int? size=null,bool? italic=null) {
            string key=original+"|"+left+"|"+right+"|"+top+"|"+bottom+"|"+color+"|"+bold+"|"+number+"|"+size+"|"+italic;int found;if(reportStyleCache.TryGetValue(key,out found))return found;
            var xfs=styles.Root.Element(N+"cellXfs");var xf=new XElement(xfs.Elements().ElementAt(original));
            if(left!=null||right!=null||top!=null||bottom!=null){
                var borders=styles.Root.Element(N+"borders");var border=new XElement(borders.Elements().ElementAt((int?)xf.Attribute("borderId")??0));
                string[] edges={"left","right","top","bottom"},values={left,right,top,bottom};
                for(int i=0;i<4;i++)if(values[i]!=null){var e=border.Element(N+edges[i]);var replacement=new XElement(N+edges[i],new XAttribute("style",values[i]),new XElement(N+"color",new XAttribute("rgb","FF000000")));if(e==null)border.Add(replacement);else e.ReplaceWith(replacement);}
                int id=borders.Elements().Count();borders.Add(border);borders.SetAttributeValue("count",id+1);xf.SetAttributeValue("borderId",id);xf.SetAttributeValue("applyBorder",1);
            }
            if(color!=null||bold.HasValue||size.HasValue||italic.HasValue){
                var fonts=styles.Root.Element(N+"fonts");var font=new XElement(fonts.Elements().ElementAt((int?)xf.Attribute("fontId")??0));
                if(color!=null){font.Elements(N+"color").Remove();font.Add(new XElement(N+"color",new XAttribute("rgb",color)));}
                if(bold.HasValue){font.Elements(N+"b").Remove();if(bold.Value)font.AddFirst(new XElement(N+"b"));}
                if(italic.HasValue){font.Elements(N+"i").Remove();if(italic.Value){var b=font.Element(N+"b");if(b!=null)b.AddAfterSelf(new XElement(N+"i"));else font.AddFirst(new XElement(N+"i"));}}
                if(size.HasValue){font.Elements(N+"sz").Remove();font.Add(new XElement(N+"sz",new XAttribute("val",size.Value)));}
                int id=fonts.Elements().Count();fonts.Add(font);fonts.SetAttributeValue("count",id+1);xf.SetAttributeValue("fontId",id);xf.SetAttributeValue("applyFont",1);
            }
            if(number.HasValue){xf.SetAttributeValue("numFmtId",number.Value);xf.SetAttributeValue("applyNumberFormat",1);}
            found=xfs.Elements().Count();xfs.Add(xf);xfs.SetAttributeValue("count",found+1);reportStyleCache[key]=found;return found;
        }
        void TableEdges(Grid g,int first,int last,int width,IEnumerable<int> percentages=null,IEnumerable<int> sections=null,int dashFirst=0) {
            var pct=new HashSet<int>(percentages??new int[0]);var breaks=new HashSet<int>(sections??new int[0]);
            for(int r=first;r<=last;r++)for(int c=1;c<=width;c++){
                if(!g.Rows.ContainsKey(r))g.Rows[r]=new SortedDictionary<int,XElement>();XElement cell;
                if(!g.Rows[r].TryGetValue(c,out cell)){cell=Cell(r,c,null,bodyStyle);g.Rows[r][c]=cell;}
                string left=c==1?"thick":breaks.Contains(c)?"medium":r>=dashFirst&&pct.Contains(c)?"dashed":"thin";
                string right=c==width?"thick":breaks.Contains(c+1)?"medium":r>=dashFirst&&pct.Contains(c+1)?"dashed":"thin";
                cell.SetAttributeValue("s",ReportStyle((int?)cell.Attribute("s")??0,left,right,r==first?"thick":"thin",r==last?"thick":"thin"));
            }
        }
        string RemapReportFormula(string formula,string destination,bool external=false) {
            // Quoted Excel strings are consumed whole; cell-like text inside them is never changed.
            return Regex.Replace(formula,@"""(?:[^""]|"""")*""|(?<![\p{L}\p{N}_])(?:(?<sheet>'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_. ]*)!)?(?<a>\$?[A-Z]{1,3}\$?\d+)(?::(?<b>\$?[A-Z]{1,3}\$?\d+))?",m=>{
                if(!m.Groups["a"].Success)return m.Value;
                string sheet=m.Groups["sheet"].Value.Trim('\'').Replace("''","'");
                if(sheet!=""&&!string.Equals(sheet,"Свод",StringComparison.OrdinalIgnoreCase)||sheet==""&&external)return m.Value;
                ReportAddress a,b;string rawA=m.Groups["a"].Value,rawB=m.Groups["b"].Value;
                if(!reportAddresses.TryGetValue(rawA.Replace("$",""),out a))return m.Value;
                if(rawB!=""&&(!reportAddresses.TryGetValue(rawB.Replace("$",""),out b)||a.Sheet!=b.Sheet))throw new System.IO.InvalidDataException("Формула с диапазоном между разными таблицами: "+m.Value+". Сводный документ не изменён.");
                Func<string,ReportAddress,string> address=(raw,target)=>{var p=Regex.Match(raw,@"^(\$?)[A-Z]+(\$?)\d+$");return p.Groups[1].Value+MergeEngine.ExcelColumn(target.Column)+p.Groups[2].Value+target.Row;};
                string result=address(rawA,a);if(rawB!=""){b=reportAddresses[rawB.Replace("$","")];result+=":"+address(rawB,b);}
                return a.Sheet==destination?result:"'"+a.Sheet.Replace("'","''")+"'!"+result;
            });
        }
        void SaveReport(string name,Grid g,int width,IEnumerable<XElement> columns,List<string> merges,string[] printAreas,XElement conditional=null) {
            var doc=Sheet(g,columns,false,width);var pr=doc.Root.Element(N+"sheetPr");pr.Add(new XElement(N+"outlinePr",new XAttribute("summaryBelow",0)),new XElement(N+"pageSetUpPr",new XAttribute("fitToPage",1)));
            doc.Root.Element(N+"sheetData").AddAfterSelf(new XElement(N+"mergeCells",new XAttribute("count",merges.Count),merges.Select(a=>new XElement(N+"mergeCell",new XAttribute("ref",a)))));
            if(conditional!=null)doc.Root.Element(N+"mergeCells").AddAfterSelf(conditional);
            doc.Root.Add(new XElement(N+"pageMargins",new XAttribute("left",.25),new XAttribute("right",.25),new XAttribute("top",.35),new XAttribute("bottom",.35),new XAttribute("header",.15),new XAttribute("footer",.15)),
                new XElement(N+"pageSetup",new XAttribute("paperSize",9),new XAttribute("orientation","landscape"),new XAttribute("fitToWidth",1),new XAttribute("fitToHeight",1)));
            PutSheet(name,doc);
            int index=book.Root.Element(N+"sheets").Elements().TakeWhile(s=>!string.Equals((string)s.Attribute("name"),name,StringComparison.OrdinalIgnoreCase)).Count();
            var names=book.Root.Element(N+"definedNames");if(names==null){names=new XElement(N+"definedNames");var calc=book.Root.Element(N+"calcPr");if(calc!=null)calc.AddBeforeSelf(names);else book.Root.Add(names);}
            names.Elements().Where(e=>(int?)e.Attribute("localSheetId")==index&&((string)e.Attribute("name")=="_xlnm.Print_Area"||(string)e.Attribute("name")=="_xlnm.Print_Titles")).Remove();
            names.Add(new XElement(N+"definedName",new XAttribute("name","_xlnm.Print_Area"),new XAttribute("localSheetId",index),string.Join(",",printAreas.Select(a=>"'"+name+"'!"+Regex.Replace(a,@"([A-Z]+)(\d+)","$$$1$$$2")))));
        }
        void PolishSplitReport(XDocument doc,string name) {
            var rows=doc.Root.Element(N+"sheetData").Elements(N+"row").ToList();
            int small=name=="Стат. по зам."?rows.Where(r=>XlsxReader.Normal(Value(r.Elements(N+"c").FirstOrDefault(c=>Column((string)c.Attribute("r"))==2)))=="ВИД ДОКУМЕНТАЦИИ").Select(r=>(int)r.Attribute("r")).DefaultIfEmpty(int.MaxValue).Min():int.MaxValue;
            int alt=name=="Итоги работы"?CaptionRow(doc,"АЛЬТЕРНАТИВНОЕ"):0;
            foreach(var row in rows){int r=(int)row.Attribute("r");foreach(var cell in row.Elements(N+"c")){
                int c=Column((string)cell.Attribute("r")),style=CellStyle(cell,bodyStyle);
                if(name=="Стат. по зам."&&r>=4&&r<small-1&&c>=3&&c<=16){string color=c<=4?"FF996600":c<=8?"FF00B050":c<=12?"FFFF0000":"FF0070C0";style=ReportStyle(style,color:color,size:r>=6&&c%2==0?(int?)10:null);}
                if(name=="Стат. по зам."&&r>=small&&c>6||name=="Итоги работы"&&alt>0&&r>=alt&&c>8)style=ReportStyle(style,"none","none","none","none");
                if(style!=CellStyle(cell,bodyStyle))cell.SetAttributeValue("s",style);
            }}
        }
        void SplitFormattedReport(XDocument source) {
            int typesRow=CaptionRow(source,"КОЛИЧЕСТВО ЗАМЕЧАНИЙ ПО ТИПАМ"),remarksRow=CaptionRow(source,"СТАТИСТИКА ПО ОСНОВНОМУ ЛИСТУ"),workRow=CaptionRow(source,"ПРОВЕРЕНО ПО ПРОВЕРЯЮЩИМ"),altRow=CaptionRow(source,"АЛЬТЕРНАТИВНОЕ ПРЕДСТАВЛЕНИЕ");
            if(remarksRow==0)remarksRow=CaptionRow(source,"СТАТИСТИКА ПО ЗАМЕЧАНИЯМ");
            // Already split (or another custom template): retain it and update all report links.
            if(typesRow==0||remarksRow==0||workRow==0||altRow==0){
                PutSheet("СВОД",source);
                foreach(string name in new[]{"Стат. по типам","Стат. по зам.","Итоги работы"}){string path=PathFor(name);if(path==null)continue;var doc=Xml(path);RefreshFormattedSummary(doc,referenceLayout.Helper,referenceLayout.BoxFiles.Count+6,referenceLayout.Volumes.Count+6);PolishSplitReport(doc,name);PutSheet(name,doc);}
                return;
            }
            string[] sheetNames={"СВОД","Стат. по типам","Стат. по зам.","Итоги работы"};var grids=sheetNames.ToDictionary(n=>n,n=>new Grid());var mergeLists=sheetNames.ToDictionary(n=>n,n=>new List<string>());
            var cells=source.Root.Element(N+"sheetData").Elements(N+"row").Elements(N+"c").ToDictionary(c=>(string)c.Attribute("r"));
            var rows=source.Root.Element(N+"sheetData").Elements(N+"row").ToDictionary(r=>(int)r.Attribute("r"));
            Func<int,int,XElement> at=(r,c)=>{XElement e;cells.TryGetValue(MergeEngine.ExcelColumn(c)+r,out e);return e;};
            Func<int,int,object> val=(r,c)=>Value(at(r,c));
            var days=rows.Keys.Where(r=>r>=workRow+4&&r<altRow&&XlsxReader.DateSerial(val(r,2)).HasValue).OrderBy(r=>r).ToList();
            var starts=Enumerable.Range(5,Column((string)cells.Values.Where(c=>Regex.IsMatch((string)c.Attribute("r"),@"^[A-Z]+"+(workRow+1)+"$")).OrderBy(c=>Column((string)c.Attribute("r"))).Last().Attribute("r"))-4).Where(c=>Text(val(workRow+1,c))!=""||at(workRow+1,c)!=null&&at(workRow+1,c).Element(N+"f")!=null).ToList();
            int matrixLast=6+days.Count,altTitle=matrixLast+3,altShift=altTitle-altRow,firstNote=rows.Keys.Where(r=>r>6&&r<typesRow&&XlsxReader.Normal(val(r,1)).StartsWith("ПРИМЕЧАНИЕ",StringComparison.Ordinal)).DefaultIfEmpty(typesRow-3).Min();
            int mainLast=firstNote,secondLast=remarksRow-typesRow+2,thirdLast=workRow-remarksRow+2,totalLast=rows.Keys.Max()+altShift;
            Action<XElement,string,int,int> copy=(cell,name,r,c)=>{
                if(cell==null)return;int oldRow=int.Parse(Regex.Match((string)cell.Attribute("r"),@"\d+$").Value,Inv);var g=grids[name];
                if(!g.Rows.ContainsKey(r))g.Rows[r]=new SortedDictionary<int,XElement>();var newCell=new XElement(cell);newCell.SetAttributeValue("r",MergeEngine.ExcelColumn(c)+r);g.Rows[r][c]=newCell;
                reportAddresses[(string)cell.Attribute("r")]=new ReportAddress{Sheet=name,Row=r,Column=c};
                if(!g.Templates.ContainsKey(r)){g.Templates[r]=new XElement(rows[oldRow]);g.Templates[r].SetAttributeValue("r",r);}
            };
            foreach(var cell in cells.Values){int r=int.Parse(Regex.Match((string)cell.Attribute("r"),@"\d+$").Value,Inv),c=Column((string)cell.Attribute("r"));
                if(r<typesRow&&r!=4&&r!=5)copy(cell,sheetNames[0],r>=5?r+1:r,c);
                else if(r>=typesRow&&r<remarksRow)copy(cell,sheetNames[1],r==typesRow?1:r-typesRow+3,c);
                else if(r>=remarksRow&&r<workRow)copy(cell,sheetNames[2],r==remarksRow?1:r-remarksRow+3,c);
                else if(r>=altRow)copy(cell,sheetNames[3],r+altShift,c);
            }
            var main=grids[sheetNames[0]];int headerStyle=CellStyle(at(4,4),headStyle);
            for(int c=1;c<=3;c++){main.Set(4,c,c==1?"Этап":c==2?"Дата / проверяющий":"Время",CellStyle(at(4,c),headerStyle));mergeLists[sheetNames[0]].Add(MergeEngine.ExcelColumn(c)+"4:"+MergeEngine.ExcelColumn(c)+"6");}
            for(int block=0;block<4;block++){int start=4+block*4;string color=new[]{"FF996600","FF00B050","FFFF0000","FF0070C0"}[block];int style=ReportStyle(headerStyle,color:color,bold:true);
                main.Set(4,start,new[]{"Всего","Проверено","Осталось проверить","Проверено за сутки"}[block],style);mergeLists[sheetNames[0]].Add(MergeEngine.ExcelColumn(start)+"4:"+MergeEngine.ExcelColumn(start+3)+"4");
                for(int c=start;c<=start+2;c+=2){main.Set(5,c,c==start?"коробов":"томов",style);main.Set(6,c,"Шт.",style);main.Set(6,c+1,"%",CellStyle(at(5,c+1),style));mergeLists[sheetNames[0]].Add(MergeEngine.ExcelColumn(c)+"5:"+MergeEngine.ExcelColumn(c+1)+"5");}
            }
            main.Heights[4]=29;main.Heights[5]=29;main.Heights[6]=29;
            for(int r=7;r<=mainLast;r++)if(main.Rows.ContainsKey(r)&&main.Rows[r].ContainsKey(1))main.Rows[r][1].RemoveNodes();main.Set(7,1,"Этап 1",CellStyle(at(6,1),bodyStyle));mergeLists[sheetNames[0]].Add("A7:A"+mainLast);
            foreach(var merge in source.Descendants(N+"mergeCell")){string raw=(string)merge.Attribute("ref");var p=Regex.Match(raw,@"^([A-Z]+)(\d+):([A-Z]+)(\d+)$");if(!p.Success)continue;int r=int.Parse(p.Groups[2].Value,Inv),end=int.Parse(p.Groups[4].Value,Inv),c=Column(p.Groups[1].Value);string name=null;int shift=0;
                if(r<=3&&end<=3){name=sheetNames[0];}
                else if(r>=typesRow+1&&end<remarksRow){name=sheetNames[1];shift=3-typesRow;}
                else if(r>=remarksRow+1&&end<workRow){name=sheetNames[2];shift=3-remarksRow;}
                else if(r>=altRow+1){name=sheetNames[3];shift=altShift;}
                if(name!=null)mergeLists[name].Add(p.Groups[1].Value+(r+shift)+":"+(name==sheetNames[0]&&end<=3?"S":p.Groups[3].Value)+(end+shift));
            }
            var work=grids[sheetNames[3]];int workHeader=ReportStyle(CellStyle(at(workRow+1,5),headerStyle),color:"FF000000",bold:true);
            work.Set(1,1,"СВОДНАЯ ТАБЛИЦА ПО ИТОГАМ РАБОТЫ ПРОВЕРЯЮЩИХ",CellStyle(at(workRow,1),titleStyle));
            work.Set(2,1,val(2,1),CellStyle(at(2,1),bodyStyle));mergeLists[sheetNames[3]].Add("A1:"+MergeEngine.ExcelColumn(2+starts.Count*2)+"1");mergeLists[sheetNames[3]].Add("A"+altTitle+":H"+altTitle);
            work.Set(4,1,"Этап",workHeader);work.Set(4,2,"Дата",workHeader);work.Set(6,1,"Этап 1",bodyStyle);work.Set(6,2,"Всего",workHeader);mergeLists[sheetNames[3]].Add("A4:A5");mergeLists[sheetNames[3]].Add("B4:B5");mergeLists[sheetNames[3]].Add("A6:A"+matrixLast);
            for(int i=0;i<starts.Count;i++){int old=starts[i],c=3+2*i;copy(at(workRow+1,old),sheetNames[3],4,c);mergeLists[sheetNames[3]].Add(MergeEngine.ExcelColumn(c)+"4:"+MergeEngine.ExcelColumn(c+1)+"4");
                work.Set(5,c,"коробов",workHeader);work.Set(5,c+1,"томов",workHeader);copy(at(workRow+3,old),sheetNames[3],6,c);copy(at(workRow+2,old),sheetNames[3],6,c+1);
                for(int d=0;d<days.Count;d++){int r=7+d;copy(at(days[d]+1,old),sheetNames[3],r,c);copy(at(days[d],old),sheetNames[3],r,c+1);}
            }
            for(int i=0;i<days.Count;i++){copy(at(days[i],2),sheetNames[3],7+i,2);reportAddresses["B"+(days[i]+1)]=new ReportAddress{Sheet=sheetNames[3],Row=7+i,Column=2};}
            for(int r=7;r<=matrixLast;r++)foreach(var cell in work.Rows[r].Values)cell.SetAttributeValue("s",ReportStyle((int?)cell.Attribute("s")??bodyStyle,color:"FF000000",bold:false));
            foreach(string name in sheetNames.Where(n=>n!=sheetNames[0])){var g=grids[name];g.Set(2,1,val(2,1),CellStyle(at(2,1),bodyStyle));}
            grids[sheetNames[2]].Set(1,1,"СТАТИСТИКА ПО ЗАМЕЧАНИЯМ",CellStyle(at(remarksRow,1),titleStyle));
            foreach(string name in new[]{sheetNames[1],sheetNames[2]})for(int c=3;c<=(name==sheetNames[1]?22:16);c+=2){var g=grids[name];g.Set(5,c,"Шт.",CellStyle(g.Rows[5][c],headerStyle));g.Set(5,c+1,"%",CellStyle(g.Rows[5][c+1],headerStyle));}
            mergeLists[sheetNames[1]].Add("A1:V1");mergeLists[sheetNames[2]].Add("A1:P1");
            foreach(string name in sheetNames)foreach(var cell in grids[name].Rows.Values.SelectMany(r=>r.Values)){var f=cell.Element(N+"f");if(f!=null)f.Value=RemapReportFormula(f.Value,name);}
            foreach(string path in entries.Keys.Where(p=>p.StartsWith("xl/worksheets/",StringComparison.Ordinal)&&p.EndsWith(".xml",StringComparison.Ordinal)&&p!=PathFor("Свод")).ToList()){
                var doc=Xml(path);bool changed=false;foreach(var f in doc.Descendants(N+"f"))if(f.Value.Contains("Свод")||f.Value.Contains("СВОД")){string next=RemapReportFormula(f.Value,null,true);if(next!=f.Value){f.Value=next;changed=true;}}if(changed)Store(path,doc);
            }
            var sourceCols=source.Root.Element(N+"cols");var sourceColumns=sourceCols==null?new List<XElement>{Col(1,1,10),Col(2,2,28),Col(3,22,11)}:sourceCols.Elements().ToList();Func<int,List<XElement>> columns=width=>sourceColumns.Where(c=>(int)c.Attribute("min")<=width).Select(c=>{var e=new XElement(c);if((int)e.Attribute("max")>width)e.SetAttributeValue("max",width);return e;}).ToList();
            // Empty cells beside the two small tables must have no table borders.
            foreach(var r in work.Rows.Where(r=>r.Key>=altTitle))foreach(var c in r.Value.Where(c=>c.Key>8))c.Value.SetAttributeValue("s",ReportStyle(CellStyle(c.Value,bodyStyle),"none","none","none","none"));
            foreach(var r in grids[sheetNames[2]].Rows.Where(r=>r.Key>=workRow-remarksRow-2))foreach(var c in r.Value.Where(c=>c.Key>6))c.Value.SetAttributeValue("s",ReportStyle(CellStyle(c.Value,bodyStyle),"none","none","none","none"));
            TableEdges(main,4,mainLast,19,Enumerable.Range(0,8).Select(i=>5+i*2),new[]{2,3,4,8,12,16},6);
            TableEdges(grids[sheetNames[1]],4,secondLast-1,22,Enumerable.Range(0,10).Select(i=>4+i*2),null,5);
            int kindHeader=CaptionRow(source,"ВИД ДОКУМЕНТАЦИИ");if(kindHeader==0)kindHeader=workRow-5;int kindTarget=kindHeader-remarksRow+3;
            foreach(var r in grids[sheetNames[2]].Rows.Where(r=>r.Key>=4&&r.Key<kindTarget-1))foreach(var c in r.Value.Where(c=>c.Key>=3&&c.Key<=16)){
                string color=c.Key<=4?"FF996600":c.Key<=8?"FF00B050":c.Key<=12?"FFFF0000":"FF0070C0";
                c.Value.SetAttributeValue("s",ReportStyle(CellStyle(c.Value,bodyStyle),color:color,size:r.Key>=6&&c.Key%2==0?(int?)10:null));
            }
            TableEdges(grids[sheetNames[2]],4,kindTarget-2,16,Enumerable.Range(0,7).Select(i=>4+i*2),null,5);TableEdges(grids[sheetNames[2]],kindTarget,thirdLast,6);
            TableEdges(work,4,matrixLast,2+starts.Count*2,null,Enumerable.Range(0,starts.Count).Select(i=>3+i*2));TableEdges(work,altTitle+1,totalLast,8);
            SaveReport(sheetNames[0],main,19,columns(19),mergeLists[sheetNames[0]],new[]{"A1:S"+(typesRow)});
            SaveReport(sheetNames[1],grids[sheetNames[1]],22,columns(22),mergeLists[sheetNames[1]],new[]{"A1:V"+secondLast});
            SaveReport(sheetNames[2],grids[sheetNames[2]],16,columns(16),mergeLists[sheetNames[2]],new[]{"A1:P"+(kindTarget-1),"A"+kindTarget+":F"+thirdLast});
            var wc=new List<XElement>{Col(1,1,10),Col(2,2,26),Col(3,2+starts.Count*2,12)};SaveReport(sheetNames[3],work,2+starts.Count*2,wc,mergeLists[sheetNames[3]],new[]{"A1:"+MergeEngine.ExcelColumn(2+starts.Count*2)+matrixLast,"A"+altTitle+":H"+totalLast});
        }
    }
}
