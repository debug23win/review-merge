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
        static string VolumeKey(MergedRow row) {return Rules.VolumeKey(row.Values[2],row.Values[3]);}
        // Excel limits a text constant in a formula to 255 characters.
        static string Literal(string s) {
            s=s??"";var parts=new List<string>();
            for(int i=0;i<s.Length;i+=200)parts.Add("\""+s.Substring(i,Math.Min(200,s.Length-i)).Replace("\"","\"\"")+"\"");
            return parts.Count==0?"\"\"":string.Join("&",parts);
        }
        // ";ПД|45;" from K and L of the main sheet. A list such as "45 и 46" keeps the boxes read at the build; another value is read as one number.
        static string BoxListFormula(string kind,string box,object value,List<string> keys) {
            string kindExpr="IF(TRIM("+kind+")=\"\",\"Не указан\",UPPER(TRIM("+kind+")))";
            string single="IF(AND(IFERROR(VALUE("+box+"),0)>0,IFERROR(VALUE("+box+"),0)=INT(IFERROR(VALUE("+box+"),0))),\";\"&"+kindExpr+"&\"|\"&TEXT(VALUE("+box+"),\"0\")&\";\",\"\")";
            if(keys.Count==0||Rules.PlainBoxNumber(value))return single;
            string template=";"+string.Join(";",keys.Select(k=>"@"+k.Substring(k.IndexOf('|'))))+";";
            return "IF(TRIM("+box+")="+Literal(Regex.Replace(Text(value)," +"," "))+",SUBSTITUTE("+Literal(template)+",\"@\","+kindExpr+"),"+single+")";
        }
        // The same remark text in several parts of one volume is counted once, in its first record.
        static string UniqueRemarkFormula(int h,int t,int rn,List<int> before,string countRef,string clean) {
            if(before.Count==0)return countRef;
            string cleanCol=MergeEngine.ExcelColumn(h+25+t);
            if(before.Count<=200)return "IF("+countRef+"=0,0,IF(OR("+string.Join(",",before.Select(j=>clean+"="+cleanCol+(j+7)))+"),0,"+countRef+"))";
            return "IF("+countRef+"=0,0,IF(SUMPRODUCT(("+LocalRange(h+28,7,rn)+"="+MergeEngine.ExcelColumn(h+28)+rn+")*("+LocalRange(h+25+t,7,rn)+"="+clean+"))=1,"+countRef+",0))";
        }
        static void MergeTitle(XDocument doc,int row,int end) {
            var merges=doc.Root.Element(N+"mergeCells");if(merges==null){merges=new XElement(N+"mergeCells");doc.Root.Element(N+"sheetData").AddAfterSelf(merges);}
            merges.Elements().Where(e=>((string)e.Attribute("ref")??"").StartsWith("A"+row+":",StringComparison.Ordinal)).Remove();
            if(end>1)merges.Add(new XElement(N+"mergeCell",new XAttribute("ref","A"+row+":"+MergeEngine.ExcelColumn(end)+row)));
            merges.SetAttributeValue("count",merges.Elements().Count());
        }
        void WriteCurrentSummary(string name,MergeResult result,DateTime asOf,List<int> rowNumbers,List<List<string>> keys,List<double?> dates,List<int> issues,List<int> checked_) {
            string oldPath=PathFor("Свод");XDocument previous=oldPath==null?null:Xml(oldPath);
            bool reference=previous!=null&&XlsxReader.Normal(Value(At(previous,7,1))).Contains("ПОЛУЧЕНО КОРОБ")&&XlsxReader.Normal(Value(At(previous,8,1))).Contains("ТОМОВ ПО АКТУ");
            var calendar=new List<double>();
            if(reference)foreach(var cell in previous.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==5).Elements(N+"c").OrderBy(e=>Column((string)e.Attribute("r")))) {
                if(Column((string)cell.Attribute("r"))<2)continue;double? d=XlsxReader.DateSerial(Value(cell));if(d.HasValue)calendar.Add(d.Value);
            }
            if(calendar.Count==0)calendar.AddRange(result.Dates.Where(d=>d.DayOfWeek!=DayOfWeek.Saturday&&d.DayOfWeek!=DayOfWeek.Sunday||d.Date==asOf.Date||dates.Contains(d.ToOADate())).Select(d=>d.ToOADate()));
            if(!calendar.Contains(asOf.ToOADate()))calendar.Add(asOf.ToOADate());calendar=reference?calendar.Distinct().ToList():calendar.Distinct().OrderBy(d=>d).ToList();
            int end=calendar.Count+1,oldHelper=HelperStart(previous);
            int templateEnd=previous==null?0:previous.Root.Element(N+"sheetData").Elements(N+"row").Elements(N+"c").Where(c=>c.Element(N+"v")!=null||c.Element(N+"f")!=null||c.Element(N+"is")!=null)
                .Select(c=>Column((string)c.Attribute("r"))).Where(c=>oldHelper==0||c<oldHelper).DefaultIfEmpty(0).Max();
            // Calculation columns start to the right of every date column and every filled cell of the summary.
            referenceLayout=PrepareReference(result,dates,asOf,Math.Max(end,templateEnd)+3);
            int h=referenceLayout.Helper,last=result.Rows.Count+6,vh=h+12;
            var g=new Grid();var defaults=SummaryStyles();
            int label=reference?StyleAt(previous,7,1,defaults[0]):defaults[0],number=reference?StyleAt(previous,7,2,defaults[1]):defaults[1];
            int greenLabel=reference?StyleAt(previous,9,1,defaults[2]):defaults[2],greenNumber=reference?StyleAt(previous,9,2,defaults[3]):defaults[3];
            int orangeLabel=reference?StyleAt(previous,11,1,defaults[4]):defaults[4],orangeNumber=reference?StyleAt(previous,11,2,defaults[5]):defaults[5];
            int oldReportEnd=oldHelper>0?Convert.ToInt32(Value(At(previous,1,oldHelper+2))??49,Inv):49;
            int generatedTo=Math.Max(end,7);
            if(reference)foreach(var row in previous.Root.Element(N+"sheetData").Elements(N+"row")) {
                int rn=(int)row.Attribute("r");double height;if(double.TryParse((string)row.Attribute("ht"),NumberStyles.Float,Inv,out height))g.Heights[rn]=height;
                g.Templates[rn]=new XElement(row);
                foreach(var c in row.Elements(N+"c")) {int col=Column((string)c.Attribute("r"));if(oldHelper>0&&col>=oldHelper&&col<=oldHelper+36)continue;if(rn>=20&&rn<=49&&col<=generatedTo)continue;
                    if((oldHelper>0||referenceLayout.OldHelper>0)&&rn>=50&&rn<=Math.Max(oldReportEnd,referenceLayout.OldHelper>0?Convert.ToInt32(Value(At(referenceLayout.Previous,1,referenceLayout.OldHelper+2))??49,Inv):49)&&col<=7)continue;
                    if(!g.Rows.ContainsKey(rn))g.Rows[rn]=new SortedDictionary<int,XElement>();g.Rows[rn][col]=new XElement(c);
                }
            }
            if(!reference) {g.Set(1,1,"СВОДНАЯ ТАБЛИЦА ПО РЕЗУЛЬТАТАМ ПРОВЕРКИ ДОКУМЕНТАЦИИ",defaults[8]);g.Heights[1]=30;g.Set(4,1,"Кол-во дней проверки:",defaults[8]);g.Set(5,1,"По состоянию на:",defaults[8]);}
            g.Set(1,h,CalculationMarker,bodyStyle);g.Set(1,h+1,asOf.ToOADate(),dateStyle);
            string[] helperNames={"Ключ короба","Проверяющий для статистики","Считается томом","Файл проверен","Дата для статистики","Есть замечания","Первая дата из источников","Ключ исходной записи","Вид документации","Проверяющий","Строка основного листа","Имя файла"};
            for(int k=0;k<helperNames.Length;k++)g.Set(6,h+k,helperNames[k],bodyStyle);
            // List item markers; RemarkCounter replaces the digits of an item number by # before counting.
            string[] tokens={"\n#. ","\n#) ","\n- ","\n* ","\n• ","\n● ","\n▪ "};
            for(int n=0;n<tokens.Length;n++)g.Set(n+1,h+27,tokens[n],bodyStyle);
            string yr=LocalRange(h+4,7,last),nameTable=NamesRange(h),tokenRange=LocalRange(h+27,1,tokens.Length);
            var volumeKeys=result.Rows.Select(VolumeKey).ToList();var earlier=new Dictionary<string,List<int>>(StringComparer.Ordinal);
            var seenRemarks=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for(int i=0;i<result.Rows.Count;i++) {
                int rn=i+7,mr=rowNumbers[i];var row=result.Rows[i];string x=MergeEngine.ExcelColumn(h+3)+rn,prior=MergeEngine.ExcelColumn(h+6)+rn,rawRef=MergeEngine.ExcelColumn(h+9)+rn;
                string kind=MainCell(name,"K",mr),box=MainCell(name,"L",mr),who=MainCell(name,"I",mr),file=MainCell(name,"C",mr),date=MainCell(name,"J",mr);
                g.Set(rn,h,Rules.BoxList(keys[i]),bodyStyle,BoxListFormula(kind,box,row.Values[11],keys[i]));
                g.Set(rn,h+1,StatKeyOfRaw(row.Values[8]),bodyStyle,"IF("+rawRef+"=\"\",\"\",IFERROR(VLOOKUP("+rawRef+","+nameTable+",3,FALSE),"+rawRef+"))");
                g.Set(rn,h+3,checked_[i],intStyle,"IF(AND("+file+"<>\"\","+who+"<>\"\"),1,0)");
                g.Set(rn,h+4,dates[i].HasValue?(object)dates[i].Value:"",dateStyle,"IF(AND("+x+"=1,ISNUMBER("+date+"),"+date+">0),IF(ISNUMBER("+prior+"),MIN(INT("+date+"),"+prior+"),INT("+date+")),\"\")");
                g.Set(rn,h+5,issues[i],intStyle,"IF(OR(COUNTIF("+MainCell(name,"M",mr)+":R"+mr+",1)>0,COUNTIF("+MainCell(name,"M",mr)+":R"+mr+",\"?*\")>0,"+MainCell(name,"S",mr)+"<>\"\","+MainCell(name,"T",mr)+"<>\"\"),1,0)");
                g.Set(rn,h+6,row.FirstReviewDate.HasValue?(object)row.FirstReviewDate.Value:null,dateStyle);
                g.Set(rn,h+7,Convert.ToBase64String(Encoding.UTF8.GetBytes(XlsxReader.Key(new SourceRow{Values=row.Values}))),bodyStyle);
                string volumeKey=volumeKeys[i];g.Set(rn,h+28,volumeKey,bodyStyle);
                List<int> before;if(!earlier.TryGetValue(volumeKey,out before)){before=new List<int>();earlier[volumeKey]=before;}
                for(int t=0;t<2;t++) {
                    string source=MainCell(name,t==0?"S":"T",mr),clean=MergeEngine.ExcelColumn(h+25+t)+rn,countRef=MergeEngine.ExcelColumn(h+23+t)+rn,text=RemarkCounter.Normalize(Text(row.Values[18+t]));int count=RemarkCounter.Count(text);
                    g.Set(rn,h+25+t,text,bodyStyle,RemarkCounter.NormalizeFormula(source));g.Set(rn,h+23+t,count,intStyle,RemarkCounter.Formula(source,clean,tokenRange));
                    int unique=seenRemarks.Add(volumeKey+"\u001f"+t+"\u001f"+text)?count:0;
                    g.Set(rn,h+29+t,unique,intStyle,UniqueRemarkFormula(h,t,rn,before,countRef,clean));
                }
                before.Add(i);
                for(int t=0;t<6;t++)g.Set(rn,h+31+t,Text(row.Values[12+t])=="1"?1:0,intStyle,"IF(COUNTIF("+MainCell(name,MergeEngine.ExcelColumn(13+t),mr)+",1)>0,1,0)");
                g.Set(rn,h+8,Text(row.Values[10]),bodyStyle,"IF("+kind+"=\"\",\"\","+kind+")");g.Set(rn,h+9,Rules.CleanName(row.Values[8]),bodyStyle,"IF(TRIM("+who+")=\"\",\"\",TRIM("+who+"))");g.Set(rn,h+10,mr,intStyle);g.Set(rn,h+11,Text(row.Values[2]),bodyStyle,"IF("+file+"=\"\",\"\","+file+")");
            }
            var volumes=result.Rows.Select((r,i)=>new{Key=volumeKeys[i],Index=i}).GroupBy(x=>x.Key,StringComparer.Ordinal).ToList();
            var volumeDates=new List<double?>();var volumeIssues=new List<int>();var counted=new List<bool>();var typeCounts=new List<int[]>();
            g.Set(6,vh,"Том (фрагменты и ИУЛ объединены)",bodyStyle);g.Set(6,vh+1,"Первая проверка тома",bodyStyle);g.Set(6,vh+2,"Есть замечания в томе",bodyStyle);
            string[] typeNames={"Несоответствие наименования тома","Несоответствие шифра тома","Несоответствие количества страниц","Отсутствуют подписи или печати на титуле","Несоответствие контрольной суммы в ИУЛ","Отсутствуют подписи в таблице ИУЛ","Текстовые замечания в описи","Текстовые дополнительные замечания"};
            for(int t=0;t<8;t++)g.Set(6,vh+3+t,typeNames[t],bodyStyle);
            for(int j=0;j<volumes.Count;j++) {
                int rn=j+7;var indices=volumes[j].Select(v=>v.Index).ToList();var vd=indices.Where(i=>dates[i].HasValue).Select(i=>dates[i].Value).ToList();double? d=vd.Count==indices.Count?(double?)vd.Max():null;int issue=indices.Any(i=>issues[i]==1)?1:0;
                // A group made only of ИУЛ is not a volume: its files and remarks are counted, the volume is not.
                bool isVolume=indices.Any(i=>!Rules.IsUl(result.Rows[i].Values[2],result.Rows[i].Values[3]));
                volumeDates.Add(d);volumeIssues.Add(issue);counted.Add(isVolume);var counts=new int[8];typeCounts.Add(counts);
                string dateCells=GroupCells(h+4,indices),issueCells=GroupCells(h+5,indices);
                g.Set(rn,vh,volumes[j].Key,bodyStyle);g.Set(rn,vh+1,d.HasValue?(object)d.Value:"",dateStyle,"IF(COUNT("+dateCells+")="+indices.Count+",MAX("+dateCells+"),\"\")");g.Set(rn,vh+2,issue,intStyle,"MAX("+issueCells+")");g.Set(rn,h+2,isVolume?1:0,intStyle);
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
                ReferenceVolume(g,h,rn,indices,result,keys,dates,issue,isVolume);
            }
            int vlast=volumes.Count+6;
            string xr=LocalRange(h+3,7,last),zr=LocalRange(h+5,7,last),cr=LocalRange(h+11,7,last),vdr=LocalRange(vh+1,7,vlast),vir=LocalRange(vh+2,7,vlast),isv=","+LocalRange(h+2,7,vlast)+",1";
            g.Set(20,1,"СТАТИСТИКА ПО ОСНОВНОМУ ЛИСТУ",defaults[8]);g.Heights[20]=24;
            g.Set(21,1,"Показатель",greenLabel);g.Set(22,1,"Файлов в реестре",label);g.Set(23,1,"Проверено файлов нарастающим итогом",greenLabel);g.Set(24,1,"Проверено томов (фрагменты и ИУЛ объединены)",greenLabel);g.Set(25,1,"Проверенных файлов с замечаниями",label);g.Set(26,1,"Проверенных томов с замечаниями",label);g.Set(27,1,"Впервые проверено файлов за сутки",label);g.Set(28,1,"Впервые проверено томов за сутки",label);
            g.Set(30,1,"КОЛИЧЕСТВО ЗАМЕЧАНИЙ ПО ТИПАМ",defaults[8]);g.Heights[30]=24;g.Set(31,1,"Тип замечания",greenLabel);
            for(int t=0;t<8;t++)g.Set(32+t,1,typeNames[t],label);g.Set(40,1,"Всего отметок и текстовых замечаний",greenLabel);g.Set(41,1,"Томов с любыми замечаниями",orangeLabel);
            g.Set(42,1,"Отметки M:R считаются по томам. В S:T считаются пункты списка или отдельные абзацы; одинаковый текст в нескольких частях тома учитывается один раз. ИУЛ без основного файла томом не считается.",noteStyle);g.Heights[42]=32;
            for(int i=0;i<calendar.Count;i++) {
                int c=i+2;double serial=calendar[i];string cs=MergeEngine.ExcelColumn(c),d=cs+"$5",criteria=",\">0\","+yr+",\"<=\"&"+d,vc=",\">0\","+vdr+",\"<=\"&"+d;bool future=serial>asOf.ToOADate();
                g.Set(21,c,serial,StyleAt(previous,5,c,defaults[6]));g.Set(31,c,serial,StyleAt(previous,5,c,defaults[6]));
                Func<int,bool> done=j=>counted[j]&&volumeDates[j].HasValue&&volumeDates[j].Value<=serial;
                int reviewed=dates.Count(v=>v.HasValue&&v.Value<=serial),vcount=Enumerable.Range(0,volumes.Count).Count(done),fileIssue=dates.Select((v,j)=>new{v,j}).Count(v=>v.v.HasValue&&v.v.Value<=serial&&issues[v.j]==1),volIssue=Enumerable.Range(0,volumes.Count).Count(j=>done(j)&&volumeIssues[j]==1);
                string guard="IF("+d+">$"+MergeEngine.ExcelColumn(h+1)+"$1,\"\",";
                Action<int,int,string,int> set=(r,value,formula,style)=>g.Set(r,c,future?(object)"":value,style,guard+formula+")");
                set(22,result.Documents,"COUNTA("+cr+")",number);set(23,reviewed,"COUNTIFS("+xr+",1,"+yr+criteria+")",greenNumber);set(24,vcount,"COUNTIFS("+vdr+vc+isv+")",greenNumber);set(25,fileIssue,"COUNTIFS("+xr+",1,"+yr+criteria+","+zr+",1)",number);set(26,volIssue,"COUNTIFS("+vdr+vc+","+vir+",1"+isv+")",number);set(27,dates.Count(v=>v==serial),"COUNTIFS("+xr+",1,"+yr+","+d+")",number);set(28,Enumerable.Range(0,volumes.Count).Count(j=>counted[j]&&volumeDates[j]==serial),"COUNTIFS("+vdr+","+d+isv+")",number);
                int total=0;for(int t=0;t<8;t++) {int count=volumeDates.Select((v,j)=>new{v,j}).Where(v=>v.v.HasValue&&v.v.Value<=serial).Sum(v=>typeCounts[v.j][t]);total+=count;string range=LocalRange(vh+3+t,7,vlast);set(32+t,count,t<6?"COUNTIFS("+vdr+vc+","+range+",1)":"SUMIFS("+range+","+vdr+vc+")",number);}
                set(40,total,"SUM("+cs+"32:"+cs+"39)",greenNumber);set(41,volIssue,cs+"26",orangeNumber);
            }
            g.Set(44,1,"Вид документации",greenLabel);g.Set(44,2,"Файлов в реестре",greenLabel);g.Set(44,3,"Проверено файлов",greenLabel);g.Set(44,4,"С замечаниями",greenLabel);g.Heights[44]=36;
            string[] kinds={"ИИ","ПД","ДПТ"};string kindRange=LocalRange(h+8,7,last),lastDate="$"+MergeEngine.ExcelColumn(h+1)+"$1";
            for(int k=0;k<3;k++) {int r=45+k;string kind=kinds[k];var indices=Enumerable.Range(0,result.Rows.Count).Where(i=>Text(result.Rows[i].Values[10])==kind).ToList();int count=indices.Count(i=>dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()),issue=indices.Count(i=>dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()&&issues[i]==1);g.Set(r,1,kind,label);g.Set(r,2,indices.Count,number,"COUNTIF("+kindRange+",A"+r+")");g.Set(r,3,count,number,"COUNTIFS("+kindRange+",A"+r+","+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+lastDate+")");g.Set(r,4,issue,number,"COUNTIFS("+kindRange+",A"+r+","+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+lastDate+","+zr+",1)");}
            int notChecked=result.Documents-dates.Count(v=>v.HasValue&&v.Value<=asOf.ToOADate());
            g.Set(48,1,"Не проверено",label);g.Set(48,2,notChecked,number,"COUNTA("+cr+")-COUNTIFS("+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+lastDate+")");g.Set(48,3,null,number);g.Set(48,4,null,number);
            for(int r=22;r<=41;r++)if(r!=29&&r!=30&&r!=31)g.Heights[r]=r==35?45:32;
            string[] reviewerHeads={"Фамилия проверяющего","Проверено томов","Томов с замечаниями","Томов за дату свода","Проверено файлов","Файлов без даты","Проверено коробов"};for(int c=0;c<reviewerHeads.Length;c++)g.Set(50,c+1,reviewerHeads[c],greenLabel);g.Heights[50]=48;
            var reviewers=result.Rows.Select(r=>StatKeyOfRaw(r.Values[8])).Where(v=>v!="").Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v=>v,StringComparer.CurrentCulture).ToList();string fileFamily=LocalRange(h+1,7,last),family=LocalRange(h+43,7,vlast);
            WriteReference(g,h,50+reviewers.Count);
            WritePhysicalTotals(g,previous,reference,calendar,asOf,defaults,label,number,greenLabel,greenNumber,orangeLabel,orangeNumber,cr,xr,yr,result,dates);
            for(int n=0;n<reviewers.Count;n++) {
                int r=51+n;string who=reviewers[n];var indices=Enumerable.Range(0,result.Rows.Count).Where(i=>string.Equals(StatKeyOfRaw(result.Rows[i].Values[8]),who,StringComparison.OrdinalIgnoreCase)).ToList();
                var members=referenceLayout.Volumes.Where(v=>v.IsVolume&&string.Equals(StatKeyOf(v.Who),who,StringComparison.OrdinalIgnoreCase)&&v.Date.HasValue&&v.Date.Value<=asOf.ToOADate()).ToList();
                string crit="SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($A"+r+",\"~\",\"~~\"),\"*\",\"~*\"),\"?\",\"~?\")";
                g.Set(r,1,who,label);g.Set(r,2,members.Count,number,"COUNTIFS("+family+","+crit+","+vdr+",\">0\","+vdr+",\"<=\"&"+lastDate+isv+")");g.Set(r,3,members.Count(v=>v.Issue==1),number,"COUNTIFS("+family+","+crit+","+vdr+",\">0\","+vdr+",\"<=\"&"+lastDate+","+vir+",1"+isv+")");g.Set(r,4,members.Count(v=>v.Date==asOf.ToOADate()),number,"COUNTIFS("+family+","+crit+","+vdr+","+lastDate+isv+")");
                g.Set(r,5,indices.Count(i=>dates[i].HasValue&&dates[i].Value<=asOf.ToOADate()),number,"COUNTIFS("+fileFamily+","+crit+","+yr+",\">0\","+yr+",\"<=\"&"+lastDate+")");
                g.Set(r,6,indices.Count(i=>!dates[i].HasValue),number,"COUNTIFS("+fileFamily+","+crit+","+xr+",1)-COUNTIFS("+fileFamily+","+crit+","+yr+",\">0\")");
                var people=referenceLayout.Names.Values.Where(v=>string.Equals(StatKeyOf(v),who,StringComparison.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();if(people.Count==0)people.Add(who);
                string done="'Справка по томам'!"+LocalRange(6,referenceLayout.ManifestStart,referenceLayout.ManifestEnd),completedBy="'Справка по томам'!"+LocalRange(8,referenceLayout.ManifestStart,referenceLayout.ManifestEnd);
                int completedCount=referenceLayout.Completed.Count(v=>v.Value.HasValue&&v.Value.Value<=asOf.ToOADate()&&people.Contains(referenceLayout.CompletedBy[v.Key],StringComparer.OrdinalIgnoreCase));
                string list="{"+string.Join(",",people.Select(v=>"\""+v.Replace("\"","\"\"")+"\""))+"}";
                g.Set(r,7,completedCount,number,"SUMPRODUCT(ISNUMBER(MATCH("+completedBy+","+list+",0))*("+done+">0)*("+done+"<="+lastDate+"))");g.Heights[r]=30;
            }
            foreach(var row in g.Rows.Values)foreach(var c in row.Values.ToList())if(Column((string)c.Attribute("r"))<h){var f=c.Element(N+"f");if(f!=null)f.Value=ReferenceFormula(f.Value,h);}
            foreach(var row in g.Rows.Values)foreach(var key in row.Keys.Where(c=>c>=h&&c<h+HelperWidth).ToList())row.Remove(key);
            foreach(int rn in g.Rows.Where(r=>r.Value.Count==0).Select(r=>r.Key).ToList())g.Rows.Remove(rn);
            MoveSummaryStatistics(g,end);
            XDocument doc=reference?new XDocument(previous):Sheet(g,new[]{Col(1,1,35),Col(2,end,11)},false,end);
            if(reference)doc.Root.Element(N+"sheetData").ReplaceWith(g.Data());
            var dimension=doc.Root.Element(N+"dimension");dimension.SetAttributeValue("ref","A1:"+MergeEngine.ExcelColumn(Math.Max(end,g.Rows.Values.SelectMany(r=>r.Keys).DefaultIfEmpty(end).Max()))+g.Rows.Keys.Max());
            var columns=doc.Root.Element(N+"cols");if(columns==null){columns=new XElement(N+"cols");doc.Root.Element(N+"sheetData").AddBeforeSelf(columns);}
            var visible=reference?columns.Elements().Where(e=>(int)e.Attribute("min")<=end).Select(e=>new XElement(e)).ToList():new List<XElement>{Col(1,1,35),Col(2,end,11)};
            foreach(var col in visible)if((int)col.Attribute("max")>end)col.SetAttributeValue("max",end);
            columns.RemoveNodes();columns.Add(visible);for(int c=2;c<=end;c++)if(!visible.Any(e=>(int)e.Attribute("min")<=c&&(int)e.Attribute("max")>=c))columns.Add(Col(c,c,11));
            var sortedCols=columns.Elements().OrderBy(e=>(int)e.Attribute("min")).ToList();sortedCols.Remove();columns.Add(sortedCols);
            var oldMerges=doc.Root.Element(N+"mergeCells");if(oldMerges!=null)oldMerges.Elements().Where(e=>Regex.IsMatch((string)e.Attribute("ref")??"",@"^A(?:20|30|32|34|42):")).Remove();
            MergeTitle(doc,1,end);MergeTitle(doc,20,end);MergeTitle(doc,34,end);MergeTitle(doc,32,end);PutSheet("Свод",doc);
            int sheetIndex=book.Root.Element(N+"sheets").Elements().Select((e,i)=>new{e,i}).First(x=>(string)x.e.Attribute("name")=="Свод").i;
            var names=book.Root.Element(N+"definedNames");if(names==null){names=new XElement(N+"definedNames");var calc=book.Root.Element(N+"calcPr");if(calc!=null)calc.AddBeforeSelf(names);else book.Root.Add(names);}
            names.Elements().Where(e=>(string)e.Attribute("name")=="_xlnm.Print_Area"&&(int?)e.Attribute("localSheetId")==sheetIndex).Remove();
            names.Add(new XElement(N+"definedName",new XAttribute("name","_xlnm.Print_Area"),new XAttribute("localSheetId",sheetIndex),"'Свод'!$A$1:$"+MergeEngine.ExcelColumn(g.Rows.Values.SelectMany(r=>r.Keys).DefaultIfEmpty(end).Max())+"$"+g.Rows.Keys.Max()));
        }
        static void MoveSummaryStatistics(Grid g,int end) {
            var map=new Dictionary<int,int>{{20,34},{21,35},{30,20},{31,21},{42,32}};for(int i=22;i<=28;i++)map[i]=i+14;for(int i=32;i<=41;i++)map[i]=i-10;
            var moved=g.Rows.Where(r=>map.ContainsKey(r.Key)).Select(r=>new{Row=r.Key,Cells=r.Value.Values.Select(e=>new XElement(e)).ToList()}).ToList();
            foreach(var item in moved)g.Rows.Remove(item.Row);
            foreach(var item in moved){int row=map[item.Row];var cells=new SortedDictionary<int,XElement>();foreach(var c in item.Cells){int col=Column((string)c.Attribute("r"));c.SetAttributeValue("r",MergeEngine.ExcelColumn(col)+row);cells[col]=c;}g.Rows[row]=cells;}
            var heights=g.Heights.Where(r=>map.ContainsKey(r.Key)).ToList();foreach(var item in heights)g.Heights.Remove(item.Key);foreach(var item in heights)g.Heights[map[item.Key]]=item.Value;
            foreach(var c in g.Rows.Values.SelectMany(r=>r.Values)){var f=c.Element(N+"f");if(f==null)continue;f.Value=Regex.Replace(f.Value,@"(?<sheet>'[^']+'!)?\$?(?<col>[A-Z]{1,3})\$?(?<row>\d+)",m=>{if(m.Groups["sheet"].Success||Column(m.Groups["col"].Value)>Math.Max(7,end))return m.Value;int row=int.Parse(m.Groups["row"].Value);return map.ContainsKey(row)?m.Value.Substring(0,m.Value.Length-m.Groups["row"].Length)+map[row]:m.Value;});}
        }
        void WritePhysicalTotals(Grid g,XDocument previous,bool reference,List<double> calendar,DateTime asOf,int[] defaults,int label,int number,int greenLabel,int greenNumber,int orangeLabel,int orangeNumber,string cr,string xr,string yr,MergeResult result,List<double?> dates) {
            string physicalPath=PathFor("Справка по томам");var physical=physicalPath==null?null:Xml(physicalPath);bool usePhysical=physical!=null&&XlsxReader.Normal(Value(At(physical,1,1))).Contains("КОРОБ");
            if(!reference) {string[] labels=usePhysical?new[]{"Получено коробов","Кол-во томов по акту","Проверено коробов нарастающим итогом","Проверено томов нарастающим итогом","Осталось проверить коробов","Осталось проверить томов","Дельта за сутки проверенных коробов","Дельта за сутки проверенных томов"}:new[]{"Коробов указано в реестре","Файлов в реестре","Коробов с проверенными файлами","Проверено файлов нарастающим итогом","Коробов без проверенных файлов","Осталось проверить файлов","Впервые отмечено коробов за сутки","Впервые проверено файлов за сутки"};for(int r=7;r<=14;r++){g.Set(r,1,labels[r-7],r==9||r==10?greenLabel:r==11||r==12?orangeLabel:label);g.Heights[r]=32;}}
            var daily=new List<KeyValuePair<double,int>>();if(usePhysical)foreach(var c in physical.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==8).Elements(N+"c")) {double? d=XlsxReader.DateSerial(Value(c));if(d.HasValue)daily.Add(new KeyValuePair<double,int>(d.Value,Column((string)c.Attribute("r"))+5));}
            var l=referenceLayout;string boxDates="'Справка по томам'!"+LocalRange(6,l.ManifestStart,l.ManifestEnd),boxNumbers="'Справка по томам'!"+LocalRange(2,l.ManifestStart,l.ManifestEnd);
            var previousSums=new double[2];
            for(int i=0;i<calendar.Count;i++) {
                int c=i+2;string cs=MergeEngine.ExcelColumn(c),prev=MergeEngine.ExcelColumn(c-1),d=cs+"$5";double serial=calendar[i];
                if(!reference||At(previous,5,c)==null) {g.Set(4,c,c-1,StyleAt(previous,4,2,defaults[9]));g.Set(5,c,serial,StyleAt(previous,5,2,defaults[6]));g.Set(6,c,17.0/24,StyleAt(previous,6,2,defaults[7]));}
                if(usePhysical) {
                    // Cached values let viewers that do not recalculate show the same numbers as Excel.
                    var received=new double[2];for(int r=7;r<=8;r++){object v=Value(At(physical,r-6,2));received[r-7]=v is double?(double)v:0;g.Set(r,c,v is double?v:0.0,number,"'Справка по томам'!B"+(r-6));}
                    if(serial>asOf.ToOADate())continue;
                    var cols=daily.Where(v=>v.Key<=serial).Select(v=>v.Value).ToList();var sums=new double[2];
                    for(int r=9;r<=10;r++) {int sourceRow=r==9?42:43;double sum=cols.Select(n=>Value(At(physical,sourceRow,n))).OfType<double>().Sum();sums[r-9]=sum;string f=cols.Count==0?"0":"SUM("+string.Join(",",cols.Select(n=>"'Справка по томам'!"+MergeEngine.ExcelColumn(n)+sourceRow))+")";g.Set(r,c,sum,greenNumber,f);}
                    for(int r=11;r<=14;r++) {
                        string f=r==11?cs+"7-"+cs+"9":r==12?cs+"8-"+cs+"10":r==13?(i==0?cs+"9":cs+"9-"+prev+"9"):(i==0?cs+"10":cs+"10-"+prev+"10");
                        double value=r==11?received[0]-sums[0]:r==12?received[1]-sums[1]:sums[r-13]-(i==0?0:previousSums[r-13]);
                        g.Set(r,c,value,r==11||r==12?orangeNumber:number,f);
                    }
                    previousSums=sums;
                } else if(!reference) {
                    // Boxes come from the box table of the reference sheet, where a list cell such as "45 и 46" gives two boxes.
                    int checkedCount=dates.Count(v=>v.HasValue&&v.Value<=serial),boxCount=l.Completed.Count(v=>v.Value.HasValue&&v.Value.Value<=serial);
                    g.Set(7,c,l.BoxCount,number,"COUNT("+boxNumbers+")");g.Set(8,c,result.Documents,number,"COUNTA("+cr+")");g.Set(9,c,boxCount,greenNumber,"COUNTIFS("+boxDates+",\">0\","+boxDates+",\"<=\"&"+d+")");g.Set(10,c,checkedCount,greenNumber,"COUNTIFS("+xr+",1,"+yr+",\">0\","+yr+",\"<=\"&"+d+")");g.Set(11,c,l.BoxCount-boxCount,orangeNumber,cs+"7-"+cs+"9");g.Set(12,c,result.Documents-checkedCount,orangeNumber,cs+"8-"+cs+"10");g.Set(13,c,l.Completed.Count(v=>v.Value==serial),number,"COUNTIF("+boxDates+","+d+")");g.Set(14,c,dates.Count(v=>v==serial),number,"COUNTIF("+yr+","+d+")");
                }
            }
        }
    }
}
