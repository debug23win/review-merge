using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ReviewMerge {
    public sealed partial class XlsxWriter {
        List<string> sharedStrings;
        object Value(XElement cell) {
            if(cell==null)return null;
            if(sharedStrings==null)sharedStrings=entries.ContainsKey("xl/sharedStrings.xml")?Xml("xl/sharedStrings.xml").Root.Elements(N+"si").Select(e=>string.Concat(e.Descendants(N+"t").Select(t=>t.Value))).ToList():new List<string>();
            string type=(string)cell.Attribute("t"),raw=(string)cell.Element(N+"v");
            if(type=="inlineStr")return string.Concat(cell.Descendants(N+"t").Select(t=>t.Value));
            int id;if(type=="s"&&int.TryParse(raw,out id)&&id>=0&&id<sharedStrings.Count)return sharedStrings[id];
            double n;if(type!="str"&&type!="e"&&double.TryParse(raw,NumberStyles.Float,Inv,out n))return n;
            return raw;
        }
        static XElement At(XDocument sheet,int row,int col) {return sheet==null?null:sheet.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==row).Elements(N+"c").FirstOrDefault(e=>Column((string)e.Attribute("r"))==col);}
        int StyleAt(XDocument sheet,int row,int col,int fallback) {var c=At(sheet,row,col);return c==null?fallback:(int?)c.Attribute("s")??fallback;}
        void RemoveLegacyHelpers(XDocument main) {
            main.Root.Element(N+"sheetData").Elements(N+"row").Elements(N+"c").Where(e=>Column((string)e.Attribute("r"))>=21&&Column((string)e.Attribute("r"))<=27).Remove();
            var cols=main.Root.Element(N+"cols");if(cols==null)return;
            foreach(var c in cols.Elements().Where(e=>(int)e.Attribute("max")>=21&&(int)e.Attribute("min")<=27).ToList()) {
                int min=(int)c.Attribute("min"),max=(int)c.Attribute("max");
                if(max>27){var tail=new XElement(c);tail.SetAttributeValue("min",28);c.AddAfterSelf(tail);}
                if(min<21)c.SetAttributeValue("max",20);else c.Remove();
            }
        }
        int SummaryStyle(int font,int fill,int num,int border,bool left) {
            int id=AddStyle(font,fill,num,true);var xf=styles.Root.Element(N+"cellXfs").Elements().Last();
            xf.SetAttributeValue("borderId",border);xf.SetAttributeValue("applyBorder",1);xf.Element(N+"alignment").SetAttributeValue("horizontal",left?"left":"center");return id;
        }
        int[] SummaryStyles() {
            var fonts=styles.Root.Element(N+"fonts");int font=fonts.Elements().Count();
            fonts.Add(new XElement(N+"font",new XElement(N+"b"),new XElement(N+"sz",new XAttribute("val",10)),new XElement(N+"name",new XAttribute("val","Times New Roman")),new XElement(N+"color",new XAttribute("rgb","FF000000"))));fonts.SetAttributeValue("count",font+1);
            var fills=styles.Root.Element(N+"fills");int green=fills.Elements().Count();
            foreach(string rgb in new[]{"FFDAF2D0","FFFBE2D5"})fills.Add(new XElement(N+"fill",new XElement(N+"patternFill",new XAttribute("patternType","solid"),new XElement(N+"fgColor",new XAttribute("rgb",rgb)),new XElement(N+"bgColor",new XAttribute("indexed",64)))));
            fills.SetAttributeValue("count",green+2);var borders=styles.Root.Element(N+"borders");int border=borders.Elements().Count();
            borders.Add(new XElement(N+"border",new[]{"left","right","top","bottom"}.Select(v=>new XElement(N+v,new XAttribute("style","thin"),new XElement(N+"color",new XAttribute("rgb","FF000000")))),new XElement(N+"diagonal")));borders.SetAttributeValue("count",border+1);
            return new[]{SummaryStyle(font,0,0,border,true),SummaryStyle(font,0,1,border,false),SummaryStyle(font,green,0,border,true),SummaryStyle(font,green,1,border,false),SummaryStyle(font,green+1,0,border,true),SummaryStyle(font,green+1,1,border,false),SummaryStyle(font,0,14,0,false),SummaryStyle(font,0,20,0,false),SummaryStyle(font,0,0,0,true),SummaryStyle(font,0,1,0,false)};
        }
        static string LocalRange(int col,int first,int last) {string c=MergeEngine.ExcelColumn(col);return "$"+c+"$"+first+":$"+c+"$"+last;}
        static string MainCell(string name,string col,int row) {return "'"+name.Replace("'","''")+"'!"+col+row;}
        static string GroupCells(int col,IEnumerable<int> indices) {
            var rows=indices.Select(i=>i+7).OrderBy(i=>i).ToList();var ranges=new List<string>();string c="$"+MergeEngine.ExcelColumn(col)+"$";
            for(int n=0;n<rows.Count;n++){int first=rows[n],last=first;while(n+1<rows.Count&&rows[n+1]==last+1){last=rows[++n];}ranges.Add(c+first+(last==first?"":":"+c+last));}return string.Join(",",ranges);
        }
        static string VolumeKey(MergedRow row) {
            string name=XlsxReader.Normal(row.Values[2]),ext=XlsxReader.Normal(row.Values[3]).TrimStart('.');
            if(ext!=""&&name.EndsWith("."+ext,StringComparison.Ordinal))name=name.Substring(0,name.Length-ext.Length-1);
            name=Regex.Replace(name,@"\s*[\(\[ _,\-]*\bФРАГМЕНТ\s*(?:№|N)?\s*\d+(?:\s*ИЗ\s*\d+)?[\)\]]*","");
            return Regex.Replace(name,@"\s+"," ").Trim();
        }
        static void MergeTitle(XDocument doc,int row,int end) {
            var merges=doc.Root.Element(N+"mergeCells");if(merges==null){merges=new XElement(N+"mergeCells");doc.Root.Element(N+"sheetData").AddAfterSelf(merges);}
            merges.Elements().Where(e=>((string)e.Attribute("ref")??"").StartsWith("A"+row+":",StringComparison.Ordinal)).Remove();
            if(end>1)merges.Add(new XElement(N+"mergeCell",new XAttribute("ref","A"+row+":"+MergeEngine.ExcelColumn(end)+row)));
            merges.SetAttributeValue("count",merges.Elements().Count());
        }
        void WriteCurrentSummary(string name,MergeResult result,DateTime asOf,List<int> rowNumbers,List<string> keys,List<double?> dates,List<int> issues,List<int> checked_,Dictionary<string,double> boxes) {
            string oldPath=PathFor("Свод");XDocument previous=oldPath==null?null:Xml(oldPath);
            bool reference=previous!=null&&XlsxReader.Normal(Value(At(previous,7,1))).Contains("ПОЛУЧЕНО КОРОБ")&&XlsxReader.Normal(Value(At(previous,8,1))).Contains("ТОМОВ ПО АКТУ");
            var calendar=new List<double>();
            if(reference)foreach(var cell in previous.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==5).Elements(N+"c").OrderBy(e=>Column((string)e.Attribute("r")))) {
                if(Column((string)cell.Attribute("r"))<2)continue;double? d=XlsxReader.DateSerial(Value(cell));if(d.HasValue)calendar.Add(d.Value);
            }
            if(calendar.Count==0)calendar.AddRange(result.Dates.Where(d=>d.DayOfWeek!=DayOfWeek.Saturday&&d.DayOfWeek!=DayOfWeek.Sunday||d.Date==asOf.Date||dates.Contains(d.ToOADate())).Select(d=>d.ToOADate()));
            if(!calendar.Contains(asOf.ToOADate()))calendar.Add(asOf.ToOADate());calendar=reference?calendar.Distinct().ToList():calendar.Distinct().OrderBy(d=>d).ToList();
            int end=calendar.Count+1,h=Math.Max(14,end+2),last=result.Rows.Count+6,vh=h+12;
            var g=new Grid();var defaults=SummaryStyles();
            int label=reference?StyleAt(previous,7,1,defaults[0]):defaults[0],number=reference?StyleAt(previous,7,2,defaults[1]):defaults[1];
            int greenLabel=reference?StyleAt(previous,9,1,defaults[2]):defaults[2],greenNumber=reference?StyleAt(previous,9,2,defaults[3]):defaults[3];
            int orangeLabel=reference?StyleAt(previous,11,1,defaults[4]):defaults[4],orangeNumber=reference?StyleAt(previous,11,2,defaults[5]):defaults[5];
            int oldHelper=0;if(previous!=null){var marker=previous.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==1).Elements(N+"c").FirstOrDefault(e=>Text(Value(e))=="ReviewMerge.FirstDates.v2");if(marker!=null)oldHelper=Column((string)marker.Attribute("r"));}
            int oldReportEnd=oldHelper>0?Convert.ToInt32(Value(At(previous,1,oldHelper+2))??49,Inv):49;
            if(reference)foreach(var row in previous.Root.Element(N+"sheetData").Elements(N+"row")) {
                int rn=(int)row.Attribute("r");double height;if(double.TryParse((string)row.Attribute("ht"),NumberStyles.Float,Inv,out height))g.Heights[rn]=height;
                g.Templates[rn]=new XElement(row);
                foreach(var c in row.Elements(N+"c")) {int col=Column((string)c.Attribute("r"));if(oldHelper>0&&col>=oldHelper&&col<=oldHelper+36)continue;if(rn>=20&&rn<=49&&col<=end)continue;
                    if(oldHelper>0&&rn>=50&&rn<=oldReportEnd&&col<=5)continue;
                    if(!g.Rows.ContainsKey(rn))g.Rows[rn]=new SortedDictionary<int,XElement>();g.Rows[rn][col]=new XElement(c);
                }
            }
            if(!reference) {g.Set(1,1,"СВОДНАЯ ТАБЛИЦА ПО РЕЗУЛЬТАТАМ ПРОВЕРКИ ДОКУМЕНТАЦИИ",defaults[8]);g.Heights[1]=30;g.Set(4,1,"Кол-во дней проверки:",defaults[8]);g.Set(5,1,"По состоянию на:",defaults[8]);}
            g.Set(1,h,"ReviewMerge.FirstDates.v2",bodyStyle);g.Set(1,h+1,asOf.ToOADate(),dateStyle);
            string[] helperNames={"Ключ короба","Первый файл короба","Первая проверка короба","Файл проверен","Дата для статистики","Есть замечания","Первая дата из источников","Ключ исходной записи","Вид документации","Проверяющий","Строка основного листа","Имя файла"};
            for(int k=0;k<helperNames.Length;k++)g.Set(6,h+k,helperNames[k],bodyStyle);
            var tokens=Enumerable.Range(1,999).Select(n=>"\n"+n+". ").Concat(Enumerable.Range(1,999).Select(n=>"\n"+n+") ")).Concat(new[]{"\n- ","\n* ","\n• ","\n● ","\n▪ "}).ToList();
            for(int n=0;n<tokens.Count;n++)g.Set(n+1,h+27,tokens[n],bodyStyle);
            var firstBoxes=new HashSet<string>();var usedBoxDate=new HashSet<string>();string ur=LocalRange(h,7,last),yr=LocalRange(h+4,7,last);
            var seenRemarks=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for(int i=0;i<result.Rows.Count;i++) {
                int rn=i+7,mr=rowNumbers[i];string u=MergeEngine.ExcelColumn(h)+rn,x=MergeEngine.ExcelColumn(h+3)+rn,y=MergeEngine.ExcelColumn(h+4)+rn,prior=MergeEngine.ExcelColumn(h+6)+rn;
                string k=keys[i];int boxFirst=k!=""&&firstBoxes.Add(k)?1:0;double? boxDate=k!=""&&dates[i].HasValue&&boxes[k]==dates[i].Value&&usedBoxDate.Add(k)?dates[i]:null;
                string kind=MainCell(name,"K",mr),box=MainCell(name,"L",mr),who=MainCell(name,"I",mr),file=MainCell(name,"C",mr),date=MainCell(name,"J",mr);
                g.Set(rn,h,k,bodyStyle,"IF(AND(OR("+kind+"=\"ИИ\","+kind+"=\"ПД\","+kind+"=\"ДПТ\"),ISNUMBER("+box+"),"+box+">0),"+kind+"&\"|\"&TEXT("+box+",\"0\"),\"\")");
                g.Set(rn,h+1,boxFirst,intStyle,"IF("+u+"=\"\",0,IF(COUNTIF("+LocalRange(h,7,rn)+","+u+")=1,1,0))");
                g.Set(rn,h+2,boxDate.HasValue?(object)boxDate.Value:"",dateStyle,"IF(OR("+u+"=\"\","+y+"=\"\"),\"\",IF(COUNTIFS("+ur+","+u+","+yr+",\">0\","+yr+",\"<\"&"+y+")=0,IF(COUNTIFS("+LocalRange(h,7,rn)+","+u+","+LocalRange(h+4,7,rn)+","+y+")=1,"+y+",\"\"),\"\"))");
                g.Set(rn,h+3,checked_[i],intStyle,"IF(AND("+file+"<>\"\","+who+"<>\"\"),1,0)");
                g.Set(rn,h+4,dates[i].HasValue?(object)dates[i].Value:"",dateStyle,"IF(AND("+x+"=1,ISNUMBER("+date+"),"+date+">0),IF(ISNUMBER("+prior+"),MIN(INT("+date+"),"+prior+"),INT("+date+")),\"\")");
                g.Set(rn,h+5,issues[i],intStyle,"IF(OR(COUNTIF("+MainCell(name,"M",mr)+":R"+mr+",1)>0,COUNTIF("+MainCell(name,"M",mr)+":R"+mr+",\"?*\")>0,"+MainCell(name,"S",mr)+"<>\"\","+MainCell(name,"T",mr)+"<>\"\"),1,0)");
                g.Set(rn,h+6,result.Rows[i].FirstReviewDate.HasValue?(object)result.Rows[i].FirstReviewDate.Value:null,dateStyle);
                g.Set(rn,h+7,Convert.ToBase64String(Encoding.UTF8.GetBytes(XlsxReader.Key(new SourceRow{Values=result.Rows[i].Values}))),bodyStyle);
                string volumeKey=VolumeKey(result.Rows[i]);g.Set(rn,h+28,volumeKey,bodyStyle);
                for(int t=0;t<2;t++) {string source=MainCell(name,t==0?"S":"T",mr),clean=MergeEngine.ExcelColumn(h+25+t)+rn,countRef=MergeEngine.ExcelColumn(h+23+t)+rn,volumeRef=MergeEngine.ExcelColumn(h+28)+rn,text=RemarkCounter.Normalize(Text(result.Rows[i].Values[18+t]));int count=RemarkCounter.Count(text);
                    g.Set(rn,h+25+t,text,bodyStyle,RemarkCounter.NormalizeFormula(source));g.Set(rn,h+23+t,count,intStyle,RemarkCounter.Formula(source,clean,LocalRange(h+27,1,tokens.Count)));
                    int unique=seenRemarks.Add(volumeKey+"\u001f"+t+"\u001f"+text)?count:0;
                    // Literal comparisons support comments longer than the COUNTIFS criteria limit.
                    g.Set(rn,h+29+t,unique,intStyle,"IF("+countRef+"=0,0,IF(SUMPRODUCT(("+LocalRange(h+28,7,rn)+"="+volumeRef+")*("+LocalRange(h+25+t,7,rn)+"="+clean+"))=1,"+countRef+",0))");
                }
                for(int t=0;t<6;t++)g.Set(rn,h+31+t,Text(result.Rows[i].Values[12+t])=="1"?1:0,intStyle,"IF(COUNTIF("+MainCell(name,MergeEngine.ExcelColumn(13+t),mr)+",1)>0,1,0)");
                g.Set(rn,h+8,Text(result.Rows[i].Values[10]),bodyStyle,"IF("+kind+"=\"\",\"\","+kind+")");g.Set(rn,h+9,Text(result.Rows[i].Values[8]),bodyStyle,"IF("+who+"=\"\",\"\","+who+")");g.Set(rn,h+10,mr,intStyle);g.Set(rn,h+11,Text(result.Rows[i].Values[2]),bodyStyle,"IF("+file+"=\"\",\"\","+file+")");
            }
            var volumes=result.Rows.Select((r,i)=>new{Key=VolumeKey(r),Index=i}).GroupBy(x=>x.Key,StringComparer.Ordinal).ToList();
            var volumeDates=new List<double?>();var volumeIssues=new List<int>();var typeCounts=new List<int[]>();
            g.Set(6,vh,"Том (фрагменты объединены)",bodyStyle);g.Set(6,vh+1,"Первая проверка тома",bodyStyle);g.Set(6,vh+2,"Есть замечания в томе",bodyStyle);
            string[] typeNames={"Несоответствие наименования тома","Несоответствие шифра тома","Несоответствие количества страниц","Отсутствуют подписи или печати на титуле","Несоответствие контрольной суммы в ИУЛ","Отсутствуют подписи в таблице ИУЛ","Текстовые замечания в описи","Текстовые дополнительные замечания"};
            for(int t=0;t<8;t++)g.Set(6,vh+3+t,typeNames[t],bodyStyle);
            for(int j=0;j<volumes.Count;j++) {
                int rn=j+7;var indices=volumes[j].Select(v=>v.Index).ToList();var vd=indices.Where(i=>dates[i].HasValue).Select(i=>dates[i].Value).ToList();double? d=vd.Count>0?(double?)vd.Min():null;int issue=indices.Any(i=>issues[i]==1)?1:0;
                volumeDates.Add(d);volumeIssues.Add(issue);var counts=new int[8];typeCounts.Add(counts);
                string dateCells=GroupCells(h+4,indices),issueCells=GroupCells(h+5,indices);
                g.Set(rn,vh,volumes[j].Key,bodyStyle);g.Set(rn,vh+1,d.HasValue?(object)d.Value:"",dateStyle,"IF(COUNT("+dateCells+")=0,\"\",MIN("+dateCells+"))");g.Set(rn,vh+2,issue,intStyle,"MAX("+issueCells+")");
                for(int t=0;t<8;t++) {
                    int col=t<6?12+t:18+t-6;counts[t]=indices.Any(i=>t<6?Text(result.Rows[i].Values[col])=="1":Text(result.Rows[i].Values[col])!="")?1:0;
                    string f="MAX("+GroupCells(h+31+t,indices)+")";
                    if(t>=6) {
                        var seen=new HashSet<string>(StringComparer.OrdinalIgnoreCase);counts[t]=indices.Select(i=>RemarkCounter.Normalize(Text(result.Rows[i].Values[col]))).Where(v=>seen.Add(v)).Sum(v=>RemarkCounter.Count(v));
                        f="SUM("+GroupCells(h+29+t-6,indices)+")";
                    }
                    if(f.Length>8192)throw new System.IO.InvalidDataException("Слишком много фрагментов в одном томе для формулы Excel: "+volumes[j].Key);
                    g.Set(rn,vh+3+t,counts[t],intStyle,f);
                }
            }
            string xr=LocalRange(h+3,7,last),zr=LocalRange(h+5,7,last),wr=LocalRange(h+2,7,last),vr=LocalRange(h+1,7,last),cr=LocalRange(h+11,7,last),vdr=LocalRange(vh+1,7,volumes.Count+6),vir=LocalRange(vh+2,7,volumes.Count+6);
            g.Set(20,1,"СТАТИСТИКА ПО ОСНОВНОМУ ЛИСТУ",defaults[8]);g.Heights[20]=24;
            g.Set(21,1,"Показатель",greenLabel);g.Set(22,1,"Файлов в реестре",label);g.Set(23,1,"Проверено файлов нарастающим итогом",greenLabel);g.Set(24,1,"Проверено томов (фрагменты объединены)",greenLabel);g.Set(25,1,"Проверенных файлов с замечаниями",label);g.Set(26,1,"Проверенных томов с замечаниями",label);g.Set(27,1,"Впервые проверено файлов за сутки",label);g.Set(28,1,"Впервые проверено томов за сутки",label);
            g.Set(30,1,"КОЛИЧЕСТВО ЗАМЕЧАНИЙ ПО ТИПАМ",defaults[8]);g.Heights[30]=24;g.Set(31,1,"Тип замечания",greenLabel);
            for(int t=0;t<8;t++)g.Set(32+t,1,typeNames[t],label);g.Set(40,1,"Всего отметок и текстовых замечаний",greenLabel);g.Set(41,1,"Томов с любыми замечаниями",orangeLabel);
            g.Set(42,1,"Отметки M:R считаются по томам. В S:T считаются пункты списка или отдельные абзацы; одинаковый текст в нескольких фрагментах учитывается один раз.",noteStyle);g.Heights[42]=32;
            WritePhysicalTotals(g,previous,reference,calendar,asOf,defaults,label,number,greenLabel,greenNumber,orangeLabel,orangeNumber,vr,wr,cr,xr,yr,result,dates,boxes);
            for(int i=0;i<calendar.Count;i++) {
                int c=i+2;double serial=calendar[i];string cs=MergeEngine.ExcelColumn(c),d=cs+"$5",criteria=",\">0\","+yr+",\"<=\"&"+d,vc=",\">0\","+vdr+",\"<=\"&"+d;bool future=serial>asOf.ToOADate();
                g.Set(21,c,serial,StyleAt(previous,5,c,defaults[6]));g.Set(31,c,serial,StyleAt(previous,5,c,defaults[6]));
                int reviewed=dates.Count(v=>v.HasValue&&v.Value<=serial),vcount=volumeDates.Count(v=>v.HasValue&&v.Value<=serial),fileIssue=dates.Select((v,j)=>new{v,j}).Count(v=>v.v.HasValue&&v.v.Value<=serial&&issues[v.j]==1),volIssue=volumeDates.Select((v,j)=>new{v,j}).Count(v=>v.v.HasValue&&v.v.Value<=serial&&volumeIssues[v.j]==1);
                string guard="IF("+d+">$"+MergeEngine.ExcelColumn(h+1)+"$1,\"\",";
                Action<int,int,string,int> set=(r,value,formula,style)=>g.Set(r,c,future?(object)"":value,style,guard+formula+")");
                set(22,result.Documents,"COUNTA("+cr+")",number);set(23,reviewed,"COUNTIFS("+xr+",1,"+yr+criteria+")",greenNumber);set(24,vcount,"COUNTIFS("+vdr+vc+")",greenNumber);set(25,fileIssue,"COUNTIFS("+xr+",1,"+yr+criteria+","+zr+",1)",number);set(26,volIssue,"COUNTIFS("+vdr+vc+","+vir+",1)",number);set(27,dates.Count(v=>v==serial),"COUNTIFS("+xr+",1,"+yr+","+d+")",number);set(28,volumeDates.Count(v=>v==serial),"COUNTIFS("+vdr+","+d+")",number);
                int total=0;for(int t=0;t<8;t++) {int count=volumeDates.Select((v,j)=>new{v,j}).Where(v=>v.v.HasValue&&v.v.Value<=serial).Sum(v=>typeCounts[v.j][t]);total+=count;string range=LocalRange(vh+3+t,7,volumes.Count+6);set(32+t,count,t<6?"COUNTIFS("+vdr+vc+","+range+",1)":"SUMIFS("+range+","+vdr+vc+")",number);}
                set(40,total,"SUM("+cs+"32:"+cs+"39)",greenNumber);set(41,volIssue,cs+"26",orangeNumber);
            }
            g.Set(44,1,"Вид документации",greenLabel);g.Set(44,2,"Файлов в реестре",greenLabel);g.Set(44,3,"Проверено файлов",greenLabel);g.Set(44,4,"С замечаниями",greenLabel);g.Set(44,5,"Осталось файлов",greenLabel);g.Heights[44]=36;
            string[] kinds={"ИИ","ПД","ДПТ"};string kindRange=LocalRange(h+8,7,last),lastDate="$"+MergeEngine.ExcelColumn(h+1)+"$1";
            for(int k=0;k<3;k++) {int r=45+k;string kind=kinds[k];var indices=Enumerable.Range(0,result.Rows.Count).Where(i=>Text(result.Rows[i].Values[10])==kind).ToList();int count=indices.Count(i=>dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()),issue=indices.Count(i=>dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()&&issues[i]==1);g.Set(r,1,kind,label);g.Set(r,2,indices.Count,number,"COUNTIF("+kindRange+",A"+r+")");g.Set(r,3,count,number,"COUNTIFS("+kindRange+",A"+r+","+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+lastDate+")");g.Set(r,4,issue,number,"COUNTIFS("+kindRange+",A"+r+","+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+lastDate+","+zr+",1)");g.Set(r,5,indices.Count-count,number,"B"+r+"-C"+r);}
            var other=Enumerable.Range(0,result.Rows.Count).Where(i=>!kinds.Contains(Text(result.Rows[i].Values[10]))).ToList();int otherChecked=other.Count(i=>dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()),otherIssues=other.Count(i=>dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()&&issues[i]==1);
            g.Set(48,1,"Не указан / иной",label);g.Set(48,2,other.Count,number,"COUNTA("+cr+")-SUM(B45:B47)");g.Set(48,3,otherChecked,number,"COUNTIFS("+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+lastDate+")-SUM(C45:C47)");g.Set(48,4,otherIssues,number,"COUNTIFS("+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+lastDate+","+zr+",1)-SUM(D45:D47)");g.Set(48,5,other.Count-otherChecked,number,"B48-C48");
            for(int r=22;r<=41;r++)if(r!=29&&r!=30&&r!=31)g.Heights[r]=r==35?45:32;
            string[] reviewerHeads={"Последний проверяющий","Проверено файлов","С замечаниями","Без даты"};for(int c=0;c<4;c++)g.Set(50,c+1,reviewerHeads[c],greenLabel);g.Heights[50]=36;
            var reviewers=result.Rows.Select(r=>Text(r.Values[8])).Where(v=>v!="").Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v=>v,StringComparer.CurrentCulture).ToList();string ir=LocalRange(h+9,7,last);
            for(int n=0;n<reviewers.Count;n++) {
                int r=51+n;string who=reviewers[n];var indices=Enumerable.Range(0,result.Rows.Count).Where(i=>string.Equals(Text(result.Rows[i].Values[8]),who,StringComparison.OrdinalIgnoreCase)).ToList();
                string crit="SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($A"+r+",\"~\",\"~~\"),\"*\",\"~*\"),\"?\",\"~?\")";
                g.Set(r,1,who,label);g.Set(r,2,indices.Count(i=>dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()),number,"COUNTIFS("+ir+","+crit+","+yr+",\">0\","+yr+",\"<=\"&"+lastDate+")");g.Set(r,3,indices.Count(i=>dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()&&issues[i]==1),number,"COUNTIFS("+ir+","+crit+","+yr+",\">0\","+yr+",\"<=\"&"+lastDate+","+zr+",1)");g.Set(r,4,indices.Count(i=>!dates[i].HasValue),number,"SUMPRODUCT(("+ir+"=$A"+r+")*("+xr+"=1)*("+yr+"=\"\"))");g.Heights[r]=30;
            }
            g.Set(1,h+2,50+reviewers.Count,intStyle);
            XDocument doc=reference?new XDocument(previous):Sheet(g,new[]{Col(1,1,35),Col(2,end,11)},false,end);
            if(reference)doc.Root.Element(N+"sheetData").ReplaceWith(g.Data());
            var dimension=doc.Root.Element(N+"dimension");dimension.SetAttributeValue("ref","A1:"+MergeEngine.ExcelColumn(Math.Max(h+36,g.Rows.Values.SelectMany(r=>r.Keys).DefaultIfEmpty(h+36).Max()))+g.Rows.Keys.Max());
            var columns=doc.Root.Element(N+"cols");if(columns==null){columns=new XElement(N+"cols");doc.Root.Element(N+"sheetData").AddBeforeSelf(columns);}
            var visible=reference?columns.Elements().Where(e=>(int)e.Attribute("min")<=end).Select(e=>new XElement(e)).ToList():new List<XElement>{Col(1,1,35),Col(2,end,11)};
            foreach(var col in visible)if((int)col.Attribute("max")>end)col.SetAttributeValue("max",end);
            columns.RemoveNodes();columns.Add(visible);for(int c=2;c<=end;c++)if(!visible.Any(e=>(int)e.Attribute("min")<=c&&(int)e.Attribute("max")>=c))columns.Add(Col(c,c,11));if(end+1<=h-1)columns.Add(Col(end+1,h-1,2));columns.Add(Col(h,h+36,18,true));
            var sortedCols=columns.Elements().OrderBy(e=>(int)e.Attribute("min")).ToList();sortedCols.Remove();columns.Add(sortedCols);
            MergeTitle(doc,1,end);MergeTitle(doc,20,end);MergeTitle(doc,30,end);MergeTitle(doc,42,end);PutSheet("Свод",doc);
        }
        void WritePhysicalTotals(Grid g,XDocument previous,bool reference,List<double> calendar,DateTime asOf,int[] defaults,int label,int number,int greenLabel,int greenNumber,int orangeLabel,int orangeNumber,string vr,string wr,string cr,string xr,string yr,MergeResult result,List<double?> dates,Dictionary<string,double> boxes) {
            string physicalPath=PathFor("Справка по томам");var physical=physicalPath==null?null:Xml(physicalPath);bool usePhysical=physical!=null&&XlsxReader.Normal(Value(At(physical,1,1))).Contains("КОРОБ");
            if(!reference) {string[] labels=usePhysical?new[]{"Получено коробов","Кол-во томов по акту","Проверено коробов нарастающим итогом","Проверено томов нарастающим итогом","Осталось проверить коробов","Осталось проверить томов","Дельта за сутки проверенных коробов","Дельта за сутки проверенных томов"}:new[]{"Коробов указано в реестре","Файлов в реестре","Коробов с проверенными файлами","Проверено файлов нарастающим итогом","Коробов без проверенных файлов","Осталось проверить файлов","Впервые отмечено коробов за сутки","Впервые проверено файлов за сутки"};for(int r=7;r<=14;r++){g.Set(r,1,labels[r-7],r==9||r==10?greenLabel:r==11||r==12?orangeLabel:label);g.Heights[r]=32;}}
            var daily=new List<KeyValuePair<double,int>>();if(usePhysical)foreach(var c in physical.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==8).Elements(N+"c")) {double? d=XlsxReader.DateSerial(Value(c));if(d.HasValue)daily.Add(new KeyValuePair<double,int>(d.Value,Column((string)c.Attribute("r"))+5));}
            for(int i=0;i<calendar.Count;i++) {
                int c=i+2;string cs=MergeEngine.ExcelColumn(c),prev=MergeEngine.ExcelColumn(c-1),d=cs+"$5";double serial=calendar[i];
                if(!reference||At(previous,5,c)==null) {g.Set(4,c,c-1,StyleAt(previous,4,2,defaults[9]));g.Set(5,c,serial,StyleAt(previous,5,2,defaults[6]));g.Set(6,c,17.0/24,StyleAt(previous,6,2,defaults[7]));}
                if(usePhysical) {
                    for(int r=7;r<=8;r++)if(!reference||At(previous,r,c)==null)g.Set(r,c,Value(At(physical,r-6,2)),number,"'Справка по томам'!B"+(r-6));
                    if(serial>asOf.ToOADate())continue;
                    var cols=daily.Where(v=>v.Key<=serial).Select(v=>v.Value).ToList();
                    for(int r=9;r<=10;r++)if(!reference||At(previous,r,c)==null||Value(At(previous,r,c))==null) {int sourceRow=r==9?42:43;double sum=cols.Select(n=>Value(At(physical,sourceRow,n))).OfType<double>().Sum();string f=cols.Count==0?"0":"SUM("+string.Join(",",cols.Select(n=>"'Справка по томам'!"+MergeEngine.ExcelColumn(n)+sourceRow))+")";g.Set(r,c,sum,greenNumber,f);}
                    for(int r=11;r<=14;r++)if(!reference||At(previous,r,c)==null||Value(At(previous,r,c))==null) {string f=r==11?cs+"7-"+cs+"9":r==12?cs+"8-"+cs+"10":r==13?(i==0?cs+"9":cs+"9-"+prev+"9"):(i==0?cs+"10":cs+"10-"+prev+"10");g.Set(r,c,null,r==11||r==12?orangeNumber:number,f);}
                } else if(!reference) {
                    int checkedCount=dates.Count(v=>v.HasValue&&v.Value<=serial),totalBoxes=result.Rows.Select(r=>Text(r.Values[10])+"|"+Text(r.Values[11])).Where(k=>!k.EndsWith("|",StringComparison.Ordinal)).Distinct().Count(),boxCount=boxes.Count(v=>v.Value<=serial);
                    g.Set(7,c,totalBoxes,number,"SUM("+vr+")");g.Set(8,c,result.Documents,number,"COUNTA("+cr+")");g.Set(9,c,boxCount,greenNumber,"COUNTIFS("+wr+",\">0\","+wr+",\"<=\"&"+d+")");g.Set(10,c,checkedCount,greenNumber,"COUNTIFS("+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+d+")");g.Set(11,c,totalBoxes-boxCount,orangeNumber,cs+"7-"+cs+"9");g.Set(12,c,result.Documents-checkedCount,orangeNumber,cs+"8-"+cs+"10");g.Set(13,c,boxes.Count(v=>v.Value==serial),number,"COUNTIF("+wr+","+d+")");g.Set(14,c,dates.Count(v=>v==serial),number,"COUNTIF("+yr+","+d+")");
                }
            }
        }
    }
}
