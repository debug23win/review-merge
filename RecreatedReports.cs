using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;

namespace ReviewMerge {
    public sealed partial class XlsxWriter {
        [StructLayout(LayoutKind.Sequential)]struct ReportTextSize {public int Width,Height;}
        [DllImport("gdi32.dll")]static extern IntPtr CreateCompatibleDC(IntPtr device);
        [DllImport("gdi32.dll",CharSet=CharSet.Unicode,EntryPoint="CreateFontW")]static extern IntPtr CreateReportFont(int height,int width,int escapement,int orientation,int weight,uint italic,uint underline,uint strikeout,uint charset,uint outputPrecision,uint clipPrecision,uint quality,uint pitchAndFamily,string face);
        [DllImport("gdi32.dll")]static extern IntPtr SelectObject(IntPtr device,IntPtr item);
        [DllImport("gdi32.dll",CharSet=CharSet.Unicode,EntryPoint="GetTextExtentPoint32W")]static extern bool MeasureReportText(IntPtr device,string text,int length,out ReportTextSize size);
        [DllImport("gdi32.dll",CharSet=CharSet.Unicode,EntryPoint="GetTextFaceW")]static extern int ReportTextFace(IntPtr device,int count,StringBuilder face);
        [DllImport("gdi32.dll")]static extern bool DeleteObject(IntPtr item);
        [DllImport("gdi32.dll")]static extern bool DeleteDC(IntPtr device);
        internal static int ReportDigitWidth(string name,double size,bool bold,bool italic,int charset) {
            string actual;return MeasureReportFont(name,size,bold,italic,charset,out actual);
        }
        static int MeasureReportFont(string name,double size,bool bold,bool italic,int charset,out string actual) {
            // Realize a pixel-sized font, so Windows display scaling cannot change XLSX widths.
            // Negative GDI height is the character/em height; 96 dpi is the workbook baseline.
            actual=name;IntPtr device=CreateCompatibleDC(IntPtr.Zero),font=IntPtr.Zero,previous=IntPtr.Zero;
            if(device==IntPtr.Zero)throw new InvalidOperationException("Не удалось измерить шрифт отчёта.");
            try {
                font=CreateReportFont(-(int)Math.Round(Math.Max(1,size)*96/72,MidpointRounding.AwayFromZero),0,0,0,bold?700:400,italic?1u:0u,0,0,(uint)charset,0,0,0,0,string.IsNullOrEmpty(name)?"Calibri":name);
                if(font==IntPtr.Zero)throw new InvalidOperationException("Не удалось создать шрифт отчёта.");
                previous=SelectObject(device,font);if(previous==IntPtr.Zero||previous==new IntPtr(-1))throw new InvalidOperationException("Не удалось выбрать шрифт отчёта.");
                var measuredFace=new StringBuilder(256);if(ReportTextFace(device,measuredFace.Capacity,measuredFace)>0)actual=measuredFace.ToString();int width=1;
                for(char digit='0';digit<='9';digit++){ReportTextSize extent;if(!MeasureReportText(device,digit.ToString(),1,out extent))throw new InvalidOperationException("Не удалось определить ширину цифры.");width=Math.Max(width,extent.Width);}
                return width;
            } finally {if(previous!=IntPtr.Zero&&previous!=new IntPtr(-1))SelectObject(device,previous);if(font!=IntPtr.Zero)DeleteObject(font);DeleteDC(device);}
        }
        // Excel's dialog uses character units; OOXML stores pixel width / maximum digit width.
        // Font-dependent padding and whole-pixel rounding matter (raw 9 can display as 8.22).
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column
        // https://github.com/ClosedXML/ClosedXML/wiki/Cell-Dimensions
        internal static double ReportColumnWidth(double characters,int maximumDigitWidth) {
            if(characters<0||characters>255||double.IsNaN(characters)||maximumDigitWidth<1)throw new ArgumentOutOfRangeException();
            int padding=2*(int)Math.Ceiling(maximumDigitWidth/4.0)+1;
            double pixels=Math.Floor((characters<1?characters*(maximumDigitWidth+padding):characters*maximumDigitWidth+padding)+0.5);
            return Math.Floor(pixels/maximumDigitWidth*256)/256;
        }
        int ReportMaximumDigitWidth(MergeResult result) {
            var named=styles.Root.Element(N+"cellStyles");
            var normal=named==null?null:named.Elements(N+"cellStyle").FirstOrDefault(e=>(int?)e.Attribute("builtinId")==0);
            if(normal==null&&named!=null)normal=named.Elements(N+"cellStyle").FirstOrDefault(e=>(string)e.Attribute("name")=="Normal"||(string)e.Attribute("name")=="Обычный");
            var styleXfs=styles.Root.Element(N+"cellStyleXfs");int xfId=normal==null?0:(int?)normal.Attribute("xfId")??0;
            var xf=styleXfs==null?null:styleXfs.Elements().ElementAtOrDefault(xfId);
            int fontId=xf==null?0:(int?)xf.Attribute("fontId")??0;var fontXml=styles.Root.Element(N+"fonts").Elements().ElementAtOrDefault(fontId);
            var nameElement=fontXml==null?null:fontXml.Element(N+"name");var sizeElement=fontXml==null?null:fontXml.Element(N+"sz");
            string name=nameElement==null?"Calibri":(string)nameElement.Attribute("val");double size=sizeElement==null?11:(double?)sizeElement.Attribute("val")??11;
            var schemeElement=fontXml==null?null:fontXml.Element(N+"scheme");string scheme=schemeElement==null?null:(string)schemeElement.Attribute("val");
            var themeRel=rels.Root.Elements().FirstOrDefault(e=>(string)e.Attribute("TargetMode")!="External"&&((string)e.Attribute("Type")??"").EndsWith("/theme",StringComparison.Ordinal));
            if(themeRel!=null&&(scheme=="major"||scheme=="minor")){
                string target=(string)themeRel.Attribute("Target"),path=Uri.UnescapeDataString(new Uri(new Uri("http://xlsx/xl/workbook.xml"),target).AbsolutePath.TrimStart('/'));
                if(entries.ContainsKey(path)){XNamespace a="http://schemas.openxmlformats.org/drawingml/2006/main";var latin=Xml(path).Descendants(a+(scheme+"Font")).Elements(a+"latin").FirstOrDefault();string themed=latin==null?null:(string)latin.Attribute("typeface");if(!string.IsNullOrEmpty(themed))name=themed;}
            }
            Func<string,bool> enabled=tag=>{var e=fontXml==null?null:fontXml.Element(N+tag);return e!=null&&(string)e.Attribute("val")!="0"&&(string)e.Attribute("val")!="false";};
            var charsetElement=fontXml==null?null:fontXml.Element(N+"charset");int charset=charsetElement==null?1:(int?)charsetElement.Attribute("val")??1;
            string actual;int width=MeasureReportFont(name,size,enabled("b"),enabled("i"),charset,out actual);
            if(!string.Equals(name,actual,StringComparison.OrdinalIgnoreCase))result.Problems.Add(new Problem{File="Итоги работы",Detail="Шрифт обычного стиля «"+name+"» заменён Windows на «"+actual+"». Ширина столбцов рассчитана по доступному шрифту; в Excel с другим шрифтом возможны отличия."});
            return width;
        }
        // Reports are outputs: their layout is rebuilt independently of the previous workbook.
        // The register and the reference sheet remain the sources of every calculation.
        int RecreatedStyle(string color,bool bold,int format=0,bool gray=false,int size=12,bool table=true) {
            int style=ReportStyle(SummaryStylesBase,table?null:"none",table?null:"none",table?null:"none",table?null:"none",color:color,bold:bold,number:format,size:size);
            string key="recreated|"+style+"|"+gray+"|"+table;int id;if(reportStyleCache.TryGetValue(key,out id))return id;
            var xfs=styles.Root.Element(N+"cellXfs");var xf=new XElement(xfs.Elements().ElementAt(style));
            xf.SetAttributeValue("fillId",gray?RecreatedGrayFill:0);xf.SetAttributeValue("applyFill",1);
            var a=xf.Element(N+"alignment");a.SetAttributeValue("horizontal",table||size>=16?"center":"left");a.SetAttributeValue("wrapText",table?1:0);
            id=xfs.Elements().Count();xfs.Add(xf);xfs.SetAttributeValue("count",id+1);reportStyleCache[key]=id;return id;
        }
        // A differential style of conditional formatting: bold, or explicitly not bold; an equal one is reused.
        int Dxf(bool bold) {
            var root=styles.Root;var dxfs=root.Element(N+"dxfs");
            if(dxfs==null){dxfs=new XElement(N+"dxfs",new XAttribute("count",0));(root.Element(N+"cellStyles")??root.Element(N+"cellXfs")).AddAfterSelf(dxfs);}
            var dxf=new XElement(N+"dxf",new XElement(N+"font",bold?new XElement(N+"b"):new XElement(N+"b",new XAttribute("val",0))));
            int found=dxfs.Elements().TakeWhile(e=>e.ToString()!=dxf.ToString()).Count();if(found<dxfs.Elements().Count())return found;
            dxfs.Add(dxf);dxfs.SetAttributeValue("count",found+1);return found;
        }
        int LeftAligned(int style) {
            string key="left|"+style;int id;if(reportStyleCache.TryGetValue(key,out id))return id;
            var xfs=styles.Root.Element(N+"cellXfs");var xf=new XElement(xfs.Elements().ElementAt(style));xf.Element(N+"alignment").SetAttributeValue("horizontal","left");
            id=xfs.Elements().Count();xfs.Add(xf);xfs.SetAttributeValue("count",id+1);reportStyleCache[key]=id;return id;
        }
        // Height of a wrapped name (the 12-point report font, about 6 points a character) in a cell of the given width in pixels;
        // 0 when it fits in one line and the default height is enough.
        static double NameHeight(string text,double pixels,bool bold) {
            double capacity=Math.Max(1,(pixels*0.75-4)/(bold?6.2:5.8)),line=0;int lines=0;
            foreach(string word in (text??"").Split(new[]{' '},StringSplitOptions.RemoveEmptyEntries)) {
                if(line>0&&line+1+word.Length<=capacity){line+=1+word.Length;continue;}
                int parts=Math.Max(1,(int)Math.Ceiling(word.Length/capacity));lines+=parts;line=word.Length-(parts-1)*capacity;
            }
            return lines<=1?0:lines*15+4;
        }
        int SummaryStylesBase,RecreatedGrayFill,RecreatedDecimal;
        void WriteRecreatedReports(MergeResult result,List<double> calendar,DateTime asOf,List<double?> dates,List<int> issues,List<int[]> typeCounts,XDocument previous) {
            var l=referenceLayout;int h=l.Helper,fl=Math.Max(7,result.Rows.Count+6),vl=Math.Max(7,l.Volumes.Count+6);
            // Every day of the reports is calculated from the current data at every build.
            SummaryStylesBase=SummaryStyles()[0];var fills=styles.Root.Element(N+"fills");RecreatedGrayFill=fills.Elements().Count();
            fills.Add(new XElement(N+"fill",new XElement(N+"patternFill",new XAttribute("patternType","solid"),new XElement(N+"fgColor",new XAttribute("rgb","FFF2F2F2")),new XElement(N+"bgColor",new XAttribute("indexed",64)))));fills.SetAttributeValue("count",RecreatedGrayFill+1);
            var fmts=styles.Root.Element(N+"numFmts");RecreatedDecimal=Math.Max(164,fmts.Elements().Select(e=>(int)e.Attribute("numFmtId")).DefaultIfEmpty(163).Max()+1);
            fmts.Add(new XElement(N+"numFmt",new XAttribute("numFmtId",RecreatedDecimal),new XAttribute("formatCode","0.0")));fmts.SetAttributeValue("count",fmts.Elements().Count());
            int df=(int)styles.Root.Element(N+"cellXfs").Elements().ElementAt(dateStyle).Attribute("numFmtId");
            Func<int,int,string> range=(offset,last)=>"'Справка по томам'!"+LocalRange(h+offset,7,last);
            string fd=range(4,fl),vd=range(13,vl),fv=range(3,fl),vi=range(14,vl),fi=range(5,fl),isv=range(2,vl),ff=range(1,fl),vf=range(43,vl);
            string bd="'Справка по томам'!"+LocalRange(6,l.ManifestStart,l.ManifestEnd),bf="'Справка по томам'!"+LocalRange(8,l.ManifestStart,l.ManifestEnd);
            string cutoff="'Справка по томам'!$"+MergeEngine.ExcelColumn(h+1)+"$1";
            Func<string,string> exact=s=>Literal(s.Replace("~","~~").Replace("*","~*").Replace("?","~?"));
            object rb=Value(At(l.Previous,1,2)),rv=Value(At(l.Previous,2,2));double totalBoxes=rb is double?(double)rb:l.BoxCount,totalVolumes=rv is double?(double)rv:l.Volumes.Count(v=>v.IsVolume);
            string tb="IF(ISNUMBER('Справка по томам'!$B$1),'Справка по томам'!$B$1,"+RegisterBoxes()+")";
            string tv="IF(ISNUMBER('Справка по томам'!$B$2),'Справка по томам'!$B$2,COUNTIF("+isv+",1))";
            // Membership is based on a fully checked volume anywhere in the history, not
            // activity on the selected day. Standalone IULs and incomplete volumes do not qualify.
            var people=l.Volumes.Where(v=>v.IsVolume&&v.Date.HasValue&&v.Date.Value>0).Select(v=>StatKeyOf(v.Who)).Where(n=>n!="").Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(n=>n,StringComparer.CurrentCulture).ToList();
            var days=calendar.Concat(l.Blocks.Where(b=>b.Date.HasValue).Select(b=>b.Date.Value)).Distinct().OrderBy(d=>d).ToList();if(days.Count==0)days.Add(asOf.ToOADate());
            Func<int,double,string,string,bool,Tuple<double,string>> metric=(kind,serial,day,who,daily)=> {
                Func<string,string> dateCriteria=d=>daily?d+","+day:d+",\">0\","+d+",\"<=\"&"+day;
                string fwho=who==""?"":","+ff+","+exact(who),vwho=who==""?"":","+vf+","+exact(who);
                Func<double?,bool> inDay=d=>d.HasValue&&(daily?d.Value==serial:d.Value>0&&d.Value<=serial);
                Func<int,bool> file=i=>inDay(dates[i])&&(who==""||string.Equals(StatKeyOfRaw(result.Rows[i].Values[8]),who,StringComparison.OrdinalIgnoreCase));
                Func<int,bool> dated=i=>inDay(l.Volumes[i].Date)&&(who==""||string.Equals(StatKeyOf(l.Volumes[i].Who),who,StringComparison.OrdinalIgnoreCase));
                Func<int,bool> volume=i=>l.Volumes[i].IsVolume&&dated(i);
                if(kind==0)return Tuple.Create(totalBoxes,tb);if(kind==1)return Tuple.Create(totalVolumes,tv);
                if(kind==2){
                    var raw=l.Names.Values.Concat(l.People.Values).Where(n=>string.Equals(StatKeyOf(n),who,StringComparison.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();if(who!=""&&raw.Count==0)raw.Add(who);
                    string formula=who==""?"COUNTIFS("+dateCriteria(bd)+")":"SUM("+string.Join(",",raw.Select(n=>"COUNTIFS("+dateCriteria(bd)+","+bf+","+exact(n)+")"))+")";
                    double count=l.Completed.Count(p=>inDay(p.Value)&&(who==""||string.Equals(StatKeyOf(l.CompletedBy[p.Key]),who,StringComparison.OrdinalIgnoreCase)));return Tuple.Create(count,formula);
                }
                if(kind==3||kind==7||kind==9){string extra=kind==7?","+vi+",1":"";return Tuple.Create((double)Enumerable.Range(0,l.Volumes.Count).Count(i=>volume(i)&&(kind!=7||l.Volumes[i].Issue==1)),"COUNTIFS("+dateCriteria(vd)+","+isv+",1"+vwho+extra+")");}
                if(kind==4)return Tuple.Create((double)result.Documents,"COUNTA("+range(11,fl)+")");
                if(kind==5||kind==6||kind==8){string extra=kind==6?","+fi+",1":"";return Tuple.Create((double)Enumerable.Range(0,result.Rows.Count).Count(i=>file(i)&&(kind!=6||issues[i]==1)),"COUNTIFS("+fv+",1,"+dateCriteria(fd)+fwho+extra+")");}
                // Remark types also count a standalone ИУЛ, which is a checked file but not a volume.
                if(kind>=10&&kind<=17){int t=kind-10;return Tuple.Create((double)Enumerable.Range(0,l.Volumes.Count).Count(i=>dated(i)&&typeCounts[i][t]==1),"COUNTIFS("+dateCriteria(vd)+vwho+","+range(15+t,vl)+",1)");}
                if(kind==20)return Tuple.Create((double)Enumerable.Range(0,result.Rows.Count).Count(i=>(who==""||string.Equals(StatKeyOfRaw(result.Rows[i].Values[8]),who,StringComparison.OrdinalIgnoreCase))&&!dates[i].HasValue&&Text(result.Rows[i].Values[8])!=""),"COUNTIFS("+fv+",1"+fwho+")-COUNTIFS("+fv+",1,"+fd+",\">0\""+fwho+")");
                throw new InvalidOperationException("Неизвестный показатель свода: "+kind);
            };
            string black="FF000000",gold="FF996600",green="FF00B050",red="FFFF0000",blue="FF0070C0";
            int plain=RecreatedStyle(black,false),header=RecreatedStyle(black,true),heading=RecreatedStyle(black,false,table:false),title=RecreatedStyle(black,true,size:16,table:false);
            Action<Grid,List<string>,string,int> start=(g,merges,text,width)=>{g.Set(1,1,text,title);g.Heights[1]=32;g.Set(2,1,"Этап 1 ВСМ-1",heading);merges.Add("A1:"+MergeEngine.ExcelColumn(width)+"1");};
            Action<Grid,int,int,Tuple<double,string>,string,bool,bool> set=(g,r,c,value,color,bold,gray)=>g.Set(r,c,value.Item1,RecreatedStyle(color,bold,1,gray),value.Item2);
            Action<Grid,int,int,int,double,string,string,bool,bool,bool,double,string> pair=(g,r,c,kind,serial,day,who,daily,bold,gray,denominator,denomFormula)=> {
                var v=metric(kind,serial,day,who,daily);bool future=serial>asOf.ToOADate();string color=black;
                g.Set(r,c,future?(object)"":v.Item1,RecreatedStyle(color,bold,1,gray),"IF("+day+">"+cutoff+",\"\","+v.Item2+")");
                g.Set(r,c+1,future?(object)"":denominator>0?v.Item1/denominator*100:0,RecreatedStyle(color,bold,RecreatedDecimal,gray,size:10),"IF("+day+">"+cutoff+",\"\",IFERROR("+MergeEngine.ExcelColumn(c)+r+"/("+denomFormula+")*100,0))");
            };
            Action<Grid,int,int,int,string> colorCells=(g,r,from,to,color)=>{for(int c=from;c<=to;c++)if(g.Rows[r].ContainsKey(c))g.Rows[r][c].SetAttributeValue("s",ReportStyle((int?)g.Rows[r][c].Attribute("s")??plain,color:color));};
            Action<Grid,int,int,string,bool> group=(g,first,last,who,hidden)=>{if(last<first)return;for(int r=first;r<=last;r++)g.Templates[r]=new XElement(N+"row",new XAttribute("r",r),new XAttribute("outlineLevel",1),new XAttribute("hidden",hidden?1:0));};
            Func<double,List<string>> owners=day=>people.Where(w=>Enumerable.Range(0,result.Rows.Count).Any(i=>dates[i]==day&&string.Equals(StatKeyOfRaw(result.Rows[i].Values[8]),w,StringComparison.OrdinalIgnoreCase))||l.Volumes.Any(v=>v.Date==day&&string.Equals(StatKeyOf(v.Who),w,StringComparison.OrdinalIgnoreCase))||l.Completed.Any(b=>b.Value==day&&string.Equals(StatKeyOf(l.CompletedBy[b.Key]),w,StringComparison.OrdinalIgnoreCase))).ToList();
            // Every date includes every reviewer: checked is cumulative, the daily block is not.
            var main=new Grid();var mm=new List<string>();start(main,mm,"СВОДНАЯ ТАБЛИЦА ПО РЕЗУЛЬТАТАМ ПРОВЕРКИ ДОКУМЕНТАЦИИ",19);
            main.Set(3,1,"Проверено — накопительно по датам и проверяющим; проверено за сутки — за указанный день.",heading);mm.Add("A3:S3");
            // Row 4 holds the volumes with remarks; the table header takes rows 5-7, the dates start at row 8.
            const int top=5;
            for(int c=1;c<=3;c++){main.Set(top,c,new[]{"Этап","Дата / проверяющий","Время"}[c-1],header);mm.Add(MergeEngine.ExcelColumn(c)+top+":"+MergeEngine.ExcelColumn(c)+(top+2));}
            for(int b=0;b<4;b++){int c=4+4*b;string color=new[]{gold,green,red,blue}[b];int st=RecreatedStyle(color,true);main.Set(top,c,new[]{"Всего","Проверено","Осталось проверить","Проверено за сутки"}[b],st);mm.Add(MergeEngine.ExcelColumn(c)+top+":"+MergeEngine.ExcelColumn(c+3)+top);for(int j=0;j<2;j++){int p=c+2*j;main.Set(top+1,p,j==0?"коробов":"томов",st);mm.Add(MergeEngine.ExcelColumn(p)+(top+1)+":"+MergeEngine.ExcelColumn(p+1)+(top+1));main.Set(top+2,p,"Шт.",st);main.Set(top+2,p+1,"%",st);}}
            main.Heights[top]=32;main.Heights[top+1]=30;main.Heights[top+2]=26;int mr=top+3,di=0;
            // Reviewers of the last date of the report are expanded; earlier and future dates stay collapsed.
            double open=days.Where(d=>d<=asOf.ToOADate()).DefaultIfEmpty(days[days.Count-1]).Max();
            foreach(double serial in days){int summary=mr;bool gray=di++%2==1;var names=people;
                for(int n=-1;n<names.Count;n++){int r=mr++;bool summaryRow=n<0;string who=summaryRow?"":names[n],day=summaryRow?"$B"+r:"$B$"+summary;main.Set(r,2,summaryRow?(object)serial:who,RecreatedStyle(black,summaryRow,summaryRow?df:0,gray));if(NameHeight(who,201,false)>0)main.Heights[r]=NameHeight(who,201,false);main.Set(r,3,17.0/24,RecreatedStyle(black,summaryRow,20,gray));
                    for(int j=0;j<2;j++){double denominator=j==0?totalBoxes:totalVolumes;string denom=j==0?tb:tv;
                        pair(main,r,8+j*2,2+j,serial,day,who,false,summaryRow,gray,denominator,denom);pair(main,r,16+j*2,2+j,serial,day,who,true,summaryRow,gray,denominator,denom);
                        if(summaryRow){
                            pair(main,r,4+j*2,j,serial,day,who,false,true,gray,denominator,denom);
                            int c=12+j*2;main.Set(r,c,serial>asOf.ToOADate()?(object)"":denominator-metric(2+j,serial,day,who,false).Item1,RecreatedStyle(red,true,1,gray),"IF("+day+">"+cutoff+",\"\","+MergeEngine.ExcelColumn(4+j*2)+r+"-"+MergeEngine.ExcelColumn(8+j*2)+r+")");main.Set(r,c+1,serial>asOf.ToOADate()?(object)"":denominator>0?(denominator-metric(2+j,serial,day,who,false).Item1)/denominator*100:0,RecreatedStyle(red,true,RecreatedDecimal,gray,size:10),"IF("+day+">"+cutoff+",\"\",IFERROR("+MergeEngine.ExcelColumn(c)+r+"/("+denom+")*100,0))");
                        }else foreach(int c in new[]{4+j*2,5+j*2,12+j*2,13+j*2})main.Set(r,c,null,RecreatedStyle(black,false,0,gray));
                    }for(int b=0;b<4;b++)colorCells(main,r,4+b*4,7+b*4,new[]{gold,green,red,blue}[b]);
                }group(main,summary+1,mr-1,"",serial!=open);if(names.Count>0&&serial!=open)main.Templates[summary]=new XElement(N+"row",new XAttribute("r",summary),new XAttribute("collapsed",1));
            }
            // Percentages of the summary are italic, size 9; their colors, borders, fill and format stay.
            for(int r=top+3;r<mr;r++)for(int c=5;c<=19;c+=2){XElement cell;if(main.Rows.ContainsKey(r)&&main.Rows[r].TryGetValue(c,out cell))cell.SetAttributeValue("s",ReportStyle(CellStyle(cell,plain),size:9,italic:true));}
            main.Set(top+3,1,"Этап 1",plain);if(mr>top+4)mm.Add("A"+(top+3)+":A"+(mr-1));TableEdges(main,top,mr-1,19,Enumerable.Range(0,8).Select(i=>5+2*i),new[]{2,3,4,8,12,16},top+2);
            // Volumes with remarks right above the table: checked up to the report date, resolved in «Замечания устранены», still being resolved.
            // Each is two lines of plain text over its block (Всего, Проверено, Осталось проверить), framed when the option says so.
            {double cut=asOf.ToOADate();var remarked=l.Volumes.Where(v=>v.IsVolume&&v.Issue==1&&v.Date.HasValue&&v.Date.Value>0&&v.Date.Value<=cut).ToList();int resolved=remarked.Count(v=>v.Resolved==1);
                string all="COUNTIFS("+vd+",\">0\","+vd+",\"<=\"&"+cutoff+","+isv+",1,"+vi+",1)",done="COUNTIFS("+vd+",\">0\","+vd+",\"<=\"&"+cutoff+","+isv+",1,"+vi+",1,"+range(50,vl)+",1)";
                var boxes=new[]{Tuple.Create(4,"Всего\nзамечаний: ",gold,remarked.Count,all),Tuple.Create(8,"Устранено\nзамечаний: ",green,resolved,done),Tuple.Create(12,"Устраняется\nзамечаний: ",red,remarked.Count-resolved,all+"-"+done)};
                int line=top-1;
                foreach(var b in boxes){int st=RecreatedStyle(b.Item3,true);if(!options.RemarksWithBorders)st=ReportStyle(st,"none","none","none","none");
                    string[] parts=b.Item2.Split('\n');
                    for(int c=b.Item1;c<b.Item1+4;c++)main.Set(line,c,null,st);main.Set(line,b.Item1,b.Item2+b.Item4.ToString(Inv),st,Literal(parts[0])+"&CHAR(10)&"+Literal(parts[1])+"&("+b.Item5+")");mm.Add(MergeEngine.ExcelColumn(b.Item1)+line+":"+MergeEngine.ExcelColumn(b.Item1+3)+line);}
                main.Heights[line]=34;}
            SaveReport("СВОД",main,19,new[]{Col(1,1,10),Col(2,2,28),Col(3,3,10)}.Concat(Enumerable.Range(4,16).Select(c=>Col(c,c,c%2==1?7:11))),mm,new[]{"A1:S"+(mr-1)});
            string[] typeNames={"Несоответствие наименования тома","Несоответствие шифра тома","Несоответствие количества страниц","Отсутствуют подписи или печати на титуле","Несоответствие контрольной суммы в ИУЛ","Отсутствуют подписи в таблице ИУЛ","Текстовые замечания в описи","Дополнительные текстовые замечания","Всего замечаний","Томов с любыми замечаниями"};
            string[] statNames={"Файлов в реестре","Проверено файлов","Проверено томов (фрагменты и ИУЛ объединены)","Проверенных файлов с замечаниями","Проверенных томов с замечаниями","Впервые проверено файлов за сутки","Впервые проверено томов за сутки"};
            foreach(bool types in new[]{true,false}){var g=new Grid();var merges=new List<string>();int width=types?22:16;string tab=types?"Стат. по типам":"Стат. по зам.";start(g,merges,types?"КОЛИЧЕСТВО ЗАМЕЧАНИЙ ПО ТИПАМ":"СТАТИСТИКА ПО ЗАМЕЧАНИЯМ",width);
                g.Set(4,1,"Этап",header);g.Set(4,2,"Дата / проверяющий",header);merges.Add("A4:A5");merges.Add("B4:B5");string[] labels=types?typeNames:statNames;
                for(int i=0;i<labels.Length;i++){int c=3+2*i;g.Set(4,c,labels[i],header);g.Set(5,c,"Шт.",header);g.Set(5,c+1,"%",header);merges.Add(MergeEngine.ExcelColumn(c)+"4:"+MergeEngine.ExcelColumn(c+1)+"4");}g.Heights[4]=92;g.Heights[5]=26;
                int row=6,index=0;foreach(double serial in days){int summary=row;bool gray=index++%2==1;var names=people;for(int n=-1;n<names.Count;n++){int r=row++;bool sum=n<0;string who=sum?"":names[n],day=sum?"$B"+r:"$B$"+summary;g.Set(r,2,sum?(object)serial:who,RecreatedStyle(black,sum,sum?df:0,gray));if(NameHeight(who,201,false)>0)g.Heights[r]=NameHeight(who,201,false);
                    for(int i=0;i<labels.Length;i++){int c=3+2*i;bool daily=!sum||!types&&i>=5;int kind=types?(i<8?10+i:i==8?18:7):new[]{4,5,3,6,7,8,9}[i];double denominator=types?totalVolumes:i==2||i==4||i==6?totalVolumes:result.Documents;string denom=types||i==2||i==4||i==6?tv:"COUNTA("+range(11,fl)+")";
                        if(kind==18){double count=Enumerable.Range(10,8).Sum(k=>metric(k,serial,day,who,daily).Item1);g.Set(r,c,serial>asOf.ToOADate()?(object)"":count,RecreatedStyle(black,sum,1,gray),"IF("+day+">"+cutoff+",\"\","+string.Join("+",Enumerable.Range(0,8).Select(k=>MergeEngine.ExcelColumn(3+2*k)+r))+")");g.Set(r,c+1,serial>asOf.ToOADate()?(object)"":denominator>0?count/denominator*100:0,RecreatedStyle(black,sum,RecreatedDecimal,gray,size:10),"IF("+day+">"+cutoff+",\"\",IFERROR("+MergeEngine.ExcelColumn(c)+r+"/("+denom+")*100,0))");}
                        else pair(g,r,c,kind,serial,day,who,daily,sum,gray,denominator,denom);
                        if(!types)colorCells(g,r,c,c+1,i==0?gold:i<=2?green:i<=4?red:blue);
                    }
                }group(g,summary+1,row-1,"",true);if(names.Count>0)g.Templates[summary]=new XElement(N+"row",new XAttribute("r",summary),new XAttribute("collapsed",1));}
                g.Set(6,1,"Этап 1",plain);if(row>7)merges.Add("A6:A"+(row-1));TableEdges(g,4,row-1,width,Enumerable.Range(0,labels.Length).Select(i=>4+2*i),null,5);
                if(!types)for(int c=3;c<=16;c++){colorCells(g,4,c,c,c<=4?gold:c<=8?green:c<=12?red:blue);colorCells(g,5,c,c,c<=4?gold:c<=8?green:c<=12?red:blue);}
                var areas=new List<string>{"A1:"+MergeEngine.ExcelColumn(width)+(row-1)};
                if(!types){int r=row+2;g.Set(r,1,"Этап",header);g.Set(r,2,"Вид документации",header);g.Set(r,3,"Файлов в реестре",header);g.Set(r,4,"Проверено файлов",header);g.Set(r,5,"С замечаниями",header);g.Heights[r]=60;g.Set(r+1,1,"Этап 1",plain);merges.Add("A"+(r+1)+":A"+(r+4));
                    string kr=range(8,fl);int j=0;foreach(string kind in new[]{"ИИ","ПД","ДПТ"}){int q=r+1+j++;g.Set(q,2,kind,plain);var ids=Enumerable.Range(0,result.Rows.Count).Where(i=>Text(result.Rows[i].Values[10])==kind).ToList();g.Set(q,3,ids.Count,plain,"COUNTIF("+kr+",$B"+q+")");g.Set(q,4,ids.Count(i=>dates[i].HasValue&&dates[i]<=asOf.ToOADate()),plain,"COUNTIFS("+kr+",$B"+q+","+fd+",\">0\","+fd+",\"<=\"&"+cutoff+")");g.Set(q,5,ids.Count(i=>dates[i].HasValue&&dates[i]<=asOf.ToOADate()&&issues[i]==1),plain,"COUNTIFS("+kr+",$B"+q+","+fd+",\">0\","+fd+",\"<=\"&"+cutoff+","+fi+",1)");}
                    g.Set(r+4,2,"Не проверено",plain);g.Set(r+4,3,result.Documents-dates.Count(d=>d.HasValue&&d<=asOf.ToOADate()),plain,"COUNTA("+range(11,fl)+")-COUNTIFS("+fd+",\">0\","+fd+",\"<=\"&"+cutoff+")");TableEdges(g,r,r+4,5);areas.Add("A"+r+":E"+(r+4));
                }
                SaveReport(tab,g,width,new[]{Col(1,1,10),Col(2,2,28)}.Concat(Enumerable.Range(3,width-2).Select(c=>Col(c,c,c%2==0?7:14))),merges,areas.ToArray());
            }
            // Reviewer matrix and its alternative grouped representation.
            var work=new Grid();var wm=new List<string>();int ww=4+people.Count*2;start(work,wm,"СВОДНАЯ ТАБЛИЦА ПО ИТОГАМ РАБОТЫ ПРОВЕРЯЮЩИХ",ww);
            // The title of the work totals is not merged: it starts in A1 and runs over the empty cells to the right.
            wm.Remove("A1:"+MergeEngine.ExcelColumn(ww)+"1");work.Rows[1][1].SetAttributeValue("s",LeftAligned(title));work.Set(4,1,"Этап",header);work.Set(4,2,"Дата",header);wm.Add("A4:A5");wm.Add("B4:B5");work.Set(6,1,"Этап 1",plain);work.Set(6,2,"Всего",header);
            // The final pair counts unique boxes/volumes directly, just like the dated summary.
            // It must not add box and volume units together or duplicate a shared box.
            var workPeople=people.Concat(new[]{""}).ToList();
            for(int i=0;i<workPeople.Count;i++){int c=3+2*i;string who=workPeople[i];work.Set(4,c,who==""?"ВСЕГО":who,header);wm.Add(MergeEngine.ExcelColumn(c)+"4:"+MergeEngine.ExcelColumn(c+1)+"4");work.Set(5,c,"коробов",header);work.Set(5,c+1,"томов",header);for(int j=0;j<2;j++){var total=metric(2+j,asOf.ToOADate(),cutoff,who,false);set(work,6,c+j,total,green,total.Item1!=0,false);}}
            for(int d=0;d<days.Count;d++){int r=7+d;double serial=days[d];bool gray=d%2==1;work.Set(r,2,serial,RecreatedStyle(black,false,df,gray));for(int i=0;i<workPeople.Count;i++)for(int j=0;j<2;j++){bool total=workPeople[i]=="",future=serial>asOf.ToOADate();var v=metric(2+j,serial,"$B"+r,workPeople[i],true);work.Set(r,3+2*i+j,future?(object)"":v.Item1,RecreatedStyle(total?green:black,!future&&v.Item1!=0,1,gray),"IF($B"+r+">"+cutoff+",\"\","+v.Item2+")");}}
            // A long name (a full name of namesakes) wraps in its pair of columns: the header row grows to show it whole.
            double headerHeight=people.Select(w=>NameHeight(w,2*(7*8.9+5),true)).DefaultIfEmpty(0).Max();if(headerHeight>0)work.Heights[4]=headerHeight;
            int matrixLast=6+days.Count;wm.Add("A6:A"+matrixLast);TableEdges(work,4,matrixLast,ww,null,Enumerable.Range(0,workPeople.Count).Select(i=>3+2*i));
            // A merged header carries its whole group's right edge, not the thin inner divider.
            for(int i=0;i<workPeople.Count;i++){var cell=work.Rows[4][3+2*i];cell.SetAttributeValue("s",ReportStyle((int)cell.Attribute("s"),right:i==workPeople.Count-1?"thick":"medium"));}
            var stage=work.Rows[6][1];stage.SetAttributeValue("s",ReportStyle((int)stage.Attribute("s"),bottom:"thick"));
            int alt=matrixLast+3;work.Set(alt,1,"Альтернативное представление по проверяющим",RecreatedStyle(black,true,table:false));wm.Add("A"+alt+":H"+alt);int ar=alt+1;
            string[] ah={"Этап","Проверяющий / дата","Проверено томов","Томов с замечаниями","Томов за день","Проверено файлов","Файлов без даты","Проверено коробов"};for(int c=1;c<=8;c++)work.Set(ar,c,ah[c-1],header);work.Heights[ar]=60;ar++;
            foreach(string who in people){int summary=ar++;work.Set(summary,2,who,header);if(NameHeight(who,7*16+5,true)>0)work.Heights[summary]=NameHeight(who,7*16+5,true);int[] kinds={3,7,9,5,20,2};for(int k=0;k<6;k++)set(work,summary,k+3,metric(kinds[k],asOf.ToOADate(),cutoff,who,k==2),green,true,false);
                foreach(double serial in days.Where(d=>owners(d).Contains(who))){int r=ar++;work.Set(r,2,serial,RecreatedStyle(black,false,df));for(int k=0;k<6;k++){var v=metric(kinds[k],serial,"$B"+r,who,true);work.Set(r,k+3,serial>asOf.ToOADate()?(object)"":v.Item1,RecreatedStyle(k==2?blue:black,false,1),"IF($B"+r+">"+cutoff+",\"\","+v.Item2+")");}}group(work,summary+1,ar-1,"",false);
            }work.Set(alt+2,1,"Этап 1",plain);if(ar>alt+3)wm.Add("A"+(alt+2)+":A"+(ar-1));TableEdges(work,alt+1,ar-1,8);
            // Keep the user's free-form notes from the old report, without copying its old tables.
            var notes=new List<string>();if(previous!=null)foreach(var cell in previous.Root.Element(N+"sheetData").Elements(N+"row").Elements(N+"c")){if(cell.Element(N+"f")!=null)continue;string text=Text(Value(cell));if(text.StartsWith("Примечание",StringComparison.OrdinalIgnoreCase)||text.StartsWith("Примечания",StringComparison.OrdinalIgnoreCase))notes.Add(text);}
            foreach(string text in notes.Distinct()){int r=mr+1;main.Set(r,1,text,heading);mm.Add("A"+r+":S"+r);mr+=2;}
            if(notes.Count>0)SaveReport("СВОД",main,19,new[]{Col(1,1,10),Col(2,2,28),Col(3,3,10)}.Concat(Enumerable.Range(4,16).Select(c=>Col(c,c,c%2==1?7:11))),mm,new[]{"A1:S"+mr});
            // Non-zero boxes and volumes of the first table are bold; the rules keep it so when Excel recalculates.
            string values="C6:"+MergeEngine.ExcelColumn(ww)+matrixLast;
            var boldRules=new XElement(N+"conditionalFormatting",new XAttribute("sqref",values),
                new XElement(N+"cfRule",new XAttribute("type","cellIs"),new XAttribute("dxfId",Dxf(true)),new XAttribute("priority",1),new XAttribute("operator","notEqual"),new XElement(N+"formula","0")),
                new XElement(N+"cfRule",new XAttribute("type","cellIs"),new XAttribute("dxfId",Dxf(false)),new XAttribute("priority",2),new XAttribute("operator","equal"),new XElement(N+"formula","0")));
            SaveReport("Итоги работы",work,Math.Max(8,ww),new[]{Col(1,1,10),Col(2,2,16),Col(3,Math.Max(8,ww),ReportColumnWidth(8.9,ReportMaximumDigitWidth(result)))},wm,new[]{"A1:"+MergeEngine.ExcelColumn(ww)+matrixLast,"A"+alt+":H"+(ar-1)},boldRules);
        }
    }
}
