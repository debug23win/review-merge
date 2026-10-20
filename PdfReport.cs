using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ReviewMerge {
    // Paper of the PDF; A4 portrait is the default.
    public enum PdfPaper { A4Portrait, A4Landscape, A3Portrait, A3Landscape }
    public static class PdfPapers {
        public static readonly string[] Names={"А4, книжная","А4, альбомная","А3, книжная","А3, альбомная"};
        // Keys of the settings file and of the --pdf-paper switch.
        public static readonly string[] Keys={"A4","A4-landscape","A3","A3-landscape"};
        public static PdfPaper Parse(string key) {
            for(int i=0;i<Keys.Length;i++)if(string.Equals(key,Keys[i],StringComparison.OrdinalIgnoreCase)||string.Equals(key,Names[i],StringComparison.OrdinalIgnoreCase))return (PdfPaper)i;
            throw new ArgumentException("Неизвестный формат PDF «"+key+"». Допустимо: "+string.Join(", ",Keys)+".");
        }
        public static string Key(PdfPaper paper){return Keys[(int)paper];}
        // Width and height in points.
        public static double[] Size(PdfPaper paper) {
            bool a3=paper==PdfPaper.A3Portrait||paper==PdfPaper.A3Landscape,landscape=paper==PdfPaper.A4Landscape||paper==PdfPaper.A3Landscape;
            double shorter=a3?842:595,longer=a3?1191:842;return landscape?new[]{longer,shorter}:new[]{shorter,longer};
        }
    }

    // What the build writes as PDF: the paper, separate files (the default) or one file, and the file names.
    // Names may contain {книга} (the workbook name) and {дата} (the report date, yyyy.MM.dd).
    public sealed class PdfSettings {
        public const string SummaryDefault="{книга} СВОД {дата}",WorkDefault="{книга} Итоги работы {дата}",CombinedDefault="{книга} {дата}";
        public PdfPaper Paper;
        public bool Combined;
        public string SummaryName=SummaryDefault,WorkName=WorkDefault,CombinedName=CombinedDefault;
        // The sheet of the second table: «Итоги работы» of PD, «Итоги проверки» of the working documentation.
        public string WorkSheet=PdWork;
        public const string PdWork="Итоги работы";
        // --pdf: the date must be the report date the book was built for, since the saved tables show that date.
        public bool CheckReportDate;
    }
    // A written PDF; Path differs from Requested when the requested file was open in another program.
    public sealed class PdfFile { public string Path,Requested; }

    // PDF of the finished workbook: the «СВОД» table and the first table of «Итоги работы», each on its own page
    // and scaled so that all its columns fit the page of the chosen paper; two files or one file of two pages.
    // The tables are drawn from the saved cells (values, number formats, fonts, fills, borders, merges),
    // so the PDF shows what Excel shows. Excel and a printer are not needed.
    public static class PdfReport {
        // The file of a name pattern next to the workbook: placeholders filled, characters Windows forbids replaced.
        public static string PathFor(string workbook,DateTime asOf,string pattern,string fallback) {
            string name=Regex.Replace(string.IsNullOrWhiteSpace(pattern)?fallback:pattern,@"\{(книга|дата)\}",m=>m.Groups[1].Value.ToLowerInvariant()=="книга"?Path.GetFileNameWithoutExtension(workbook):asOf.ToString("yyyy.MM.dd",CultureInfo.InvariantCulture),RegexOptions.IgnoreCase);
            name=Regex.Replace(name,@"[\\/:*?""<>|\x00-\x1f]","_").Trim().TrimEnd('.');
            if(name.EndsWith(".pdf",StringComparison.OrdinalIgnoreCase))name=name.Substring(0,name.Length-4).TrimEnd();
            return Path.Combine(Path.GetDirectoryName(workbook),(name==""?"PDF":name)+".pdf");
        }
        public static List<PdfFile> Write(string workbook,DateTime asOf,PdfSettings settings) {
            settings=settings??new PdfSettings();
            var book=new PdfWorkbook(workbook);
            // A book of the working documentation has «Итоги проверки» in place of «Итоги работы» (also for --pdf).
            string work=settings.WorkSheet??PdfSettings.PdWork;if(!book.Has(work)&&book.Has(RdEngine.RdWork))work=RdEngine.RdWork;
            string workDefault=work==PdfSettings.PdWork?PdfSettings.WorkDefault:"{книга} "+work+" {дата}",workName=settings.WorkName==PdfSettings.WorkDefault?workDefault:settings.WorkName;
            var parts=settings.Combined?new[]{Tuple.Create(new[]{"СВОД",work},PathFor(workbook,asOf,settings.CombinedName,PdfSettings.CombinedDefault))}
                :new[]{Tuple.Create(new[]{"СВОД"},PathFor(workbook,asOf,settings.SummaryName,PdfSettings.SummaryDefault)),Tuple.Create(new[]{work},PathFor(workbook,asOf,workName,workDefault))};
            if(parts.Length==2&&string.Equals(parts[0].Item2,parts[1].Item2,StringComparison.OrdinalIgnoreCase))
                throw new IOException("у PDF «СВОД» и «"+work+"» одинаковое имя «"+System.IO.Path.GetFileName(parts[0].Item2)+"»: задайте разные имена.");
            var size=PdfPapers.Size(settings.Paper);string day=asOf.ToString("dd.MM.yyyy",CultureInfo.InvariantCulture);
            if(settings.CheckReportDate){var built=book.ReportDate();if(built.HasValue&&built.Value!=asOf.Date)
                throw new InvalidDataException("книга рассчитана на дату свода "+built.Value.ToString("dd.MM.yyyy",CultureInfo.InvariantCulture)+", а для PDF указана "+day+": укажите дату "+built.Value.ToString("yyyy-MM-dd",CultureInfo.InvariantCulture)+" или соберите свод на нужную дату.");}
            // Every part is drawn and every name checked before the first file is written, so a failure leaves no partial PDF.
            var data=new List<byte[]>();
            foreach(var part in parts){
                CheckName(part.Item2);
                var pdf=new PdfDocument("Свод проверки "+day);foreach(string sheet in part.Item1)book.Render(pdf,sheet,0,size[0],size[1]);data.Add(pdf.Save());
            }
            var files=new List<PdfFile>();
            try{for(int i=0;i<parts.Length;i++)files.Add(new PdfFile{Requested=parts[i].Item2,Path=Save(data[i],parts[i].Item2)});}
            catch{foreach(var file in files)TryDelete(file.Path);throw;}
            return files;
        }
        // Windows allows 255 characters in a name and 259 in a path; the indexed name «имя (99).pdf» must fit too.
        // Other systems report a long name when the file is saved.
        static void CheckName(string path) {
            string name=System.IO.Path.GetFileName(path),full;
            try{full=System.IO.Path.GetFullPath(path);}catch(PathTooLongException){full=null;}catch(NotSupportedException){full=null;}catch(ArgumentException){full=null;}
            bool windows=System.IO.Path.DirectorySeparatorChar=='\\';
            if(full==null||name.Length+5>255||windows&&full.Length+5>259)throw new IOException("имя PDF «"+name+"» слишком длинное или недопустимое: задайте имя короче.");
        }
        // A file open in another program cannot be replaced: the PDF goes to the first free «имя (1).pdf», «имя (2).pdf»…
        static string Save(byte[] data,string path) {
            string folder=System.IO.Path.GetDirectoryName(path),stem=System.IO.Path.GetFileNameWithoutExtension(path);
            for(int i=0;i<100;i++) {
                // A short temporary name: «имя.pdf.tmp» would be longer than the name itself.
                string candidate=i==0?path:System.IO.Path.Combine(folder,stem+" ("+i+").pdf"),temp=System.IO.Path.Combine(folder,"~pdf-"+Guid.NewGuid().ToString("N").Substring(0,12)+".tmp");
                try {File.WriteAllBytes(temp,data);}
                catch(UnauthorizedAccessException){throw new IOException("нет доступа к папке «"+folder+"».");}
                catch(PathTooLongException){throw new IOException("путь к папке PDF слишком длинный: «"+folder+"».");}
                try {if(File.Exists(candidate))File.Delete(candidate);File.Move(temp,candidate);return candidate;}
                catch(PathTooLongException){TryDelete(temp);throw new IOException("имя PDF «"+System.IO.Path.GetFileName(candidate)+"» слишком длинное: задайте имя короче.");}
                catch(IOException){TryDelete(temp);}
                catch(UnauthorizedAccessException){TryDelete(temp);}
            }
            throw new IOException("«"+System.IO.Path.GetFileName(path)+"» и файлы с индексами до 99 открыты или недоступны.");
        }
        static void TryDelete(string path){try{if(File.Exists(path))File.Delete(path);}catch(IOException){}catch(UnauthorizedAccessException){}}
    }

    sealed class PdfStyle {
        public string Font="Calibri",Color="FF000000",Fill,Format="General",Horizontal="general",Vertical="bottom";
        public double Size=11;public bool Bold,Italic,Wrap;
        public readonly string[] Border=new string[4],BorderColor=new string[4];
    }

    // Reads the saved workbook part by part; only what the drawing needs.
    sealed class PdfWorkbook {
        static readonly XNamespace N="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly XNamespace R="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static readonly CultureInfo Inv=CultureInfo.InvariantCulture;
        readonly Dictionary<string,byte[]> parts=new Dictionary<string,byte[]>(StringComparer.Ordinal);
        readonly List<string> strings=new List<string>();
        readonly List<PdfStyle> styles=new List<PdfStyle>();
        readonly XDocument book;
        readonly List<Tuple<string,string>> sheets=new List<Tuple<string,string>>();
        string normalFont="Calibri";double normalSize=11;
        public PdfWorkbook(string path) {
            using(var stream=File.OpenRead(path))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))
                foreach(var e in zip.Entries){using(var m=new MemoryStream())using(var s=e.Open()){s.CopyTo(m);parts[e.FullName]=m.ToArray();}}
            book=Load("xl/workbook.xml");var rels=Load("xl/_rels/workbook.xml.rels");
            foreach(var sheet in book.Root.Element(N+"sheets").Elements(N+"sheet")){
                string id=(string)sheet.Attribute(R+"id");var rel=rels.Root.Elements().FirstOrDefault(e=>(string)e.Attribute("Id")==id);if(rel==null)continue;
                string target=((string)rel.Attribute("Target")).Replace('\\','/');sheets.Add(Tuple.Create((string)sheet.Attribute("name"),target.StartsWith("/")?target.TrimStart('/'):"xl/"+target));
            }
            if(parts.ContainsKey("xl/sharedStrings.xml"))foreach(var si in Load("xl/sharedStrings.xml").Root.Elements(N+"si"))strings.Add(string.Concat(si.Descendants(N+"t").Where(t=>t.Parent.Name!=N+"rPh").Select(t=>t.Value)));
            ReadStyles();
        }
        XDocument Load(string part){using(var m=new MemoryStream(parts[part]))return XDocument.Load(m);}
        public bool Has(string sheet){return sheets.Any(s=>string.Equals(s.Item1,sheet,StringComparison.OrdinalIgnoreCase));}
        // The report date the book was built for: RD «СВОД»!U2 under «Дата отчёта», PD the cell after the marker in row 1 of «Справка по томам»; null when unknown.
        public DateTime? ReportDate() {
            Func<object,DateTime?> date=v=>v is double&&(double)v>=1&&(double)v<2958466?(DateTime?)DateTime.FromOADate((double)v).Date:null;
            var rd=Row("СВОД",1);object head;if(rd.TryGetValue(21,out head)&&head as string=="Дата отчёта"&&Has(RdEngine.RdWork)){object v;return Row("СВОД",2).TryGetValue(21,out v)?date(v):null;}
            var pd=Row("Справка по томам",1);
            foreach(var cell in pd){var text=cell.Value as string;object v;if(text!=null&&text.StartsWith("ReviewMerge.FirstDates.v",StringComparison.Ordinal))return pd.TryGetValue(cell.Key+1,out v)?date(v):null;}
            return null;
        }
        // The values of one row by column, read up to that row only.
        Dictionary<int,object> Row(string sheetName,int number) {
            var result=new Dictionary<int,object>();var sheet=sheets.FirstOrDefault(s=>string.Equals(s.Item1,sheetName,StringComparison.OrdinalIgnoreCase));
            if(sheet==null||!parts.ContainsKey(sheet.Item2))return result;
            using(var m=new MemoryStream(parts[sheet.Item2]))using(var reader=System.Xml.XmlReader.Create(m))
                while(reader.ReadToFollowing("row",N.NamespaceName)){
                    int r;if(!int.TryParse(reader.GetAttribute("r"),NumberStyles.Integer,Inv,out r)||r>number)break;if(r<number)continue;
                    foreach(var c in ((XElement)XNode.ReadFrom(reader)).Elements(N+"c")){string address=(string)c.Attribute("r");if(address!=null)result[Column(address)]=CellValue(c);}
                    break;
                }
            return result;
        }
        static string Rgb(XElement color,string fallback){var rgb=color==null?null:(string)color.Attribute("rgb");return string.IsNullOrEmpty(rgb)?fallback:rgb.Length==6?"FF"+rgb:rgb;}
        void ReadStyles() {
            if(!parts.ContainsKey("xl/styles.xml")){styles.Add(new PdfStyle());return;}
            var doc=Load("xl/styles.xml");var root=doc.Root;
            var formats=new Dictionary<int,string>();var numFmts=root.Element(N+"numFmts");if(numFmts!=null)foreach(var f in numFmts.Elements(N+"numFmt"))formats[(int)f.Attribute("numFmtId")]=(string)f.Attribute("formatCode");
            var fonts=root.Element(N+"fonts").Elements(N+"font").ToList();var fills=root.Element(N+"fills").Elements(N+"fill").ToList();var borders=root.Element(N+"borders").Elements(N+"border").ToList();
            Func<XElement,string,bool> flag=(font,tag)=>{var e=font.Element(N+tag);return e!=null&&(string)e.Attribute("val")!="0"&&(string)e.Attribute("val")!="false";};
            var cellStyleXfs=root.Element(N+"cellStyleXfs");var normal=cellStyleXfs==null?null:cellStyleXfs.Elements(N+"xf").FirstOrDefault();
            var normalXml=fonts.ElementAtOrDefault(normal==null?0:(int?)normal.Attribute("fontId")??0);
            if(normalXml!=null){var name=normalXml.Element(N+"name");var sz=normalXml.Element(N+"sz");if(name!=null)normalFont=(string)name.Attribute("val");if(sz!=null)normalSize=(double)sz.Attribute("val");}
            foreach(var xf in root.Element(N+"cellXfs").Elements(N+"xf")) {
                var st=new PdfStyle();var font=fonts.ElementAtOrDefault((int?)xf.Attribute("fontId")??0);
                if(font!=null){var name=font.Element(N+"name");var sz=font.Element(N+"sz");if(name!=null)st.Font=(string)name.Attribute("val");if(sz!=null)st.Size=(double)sz.Attribute("val");st.Bold=flag(font,"b");st.Italic=flag(font,"i");st.Color=Rgb(font.Element(N+"color"),"FF000000");}
                var fill=fills.ElementAtOrDefault((int?)xf.Attribute("fillId")??0);var pattern=fill==null?null:fill.Element(N+"patternFill");
                if(pattern!=null&&(string)pattern.Attribute("patternType")=="solid")st.Fill=Rgb(pattern.Element(N+"fgColor"),null);
                var border=borders.ElementAtOrDefault((int?)xf.Attribute("borderId")??0);string[] edges={"left","right","top","bottom"};
                for(int i=0;i<4;i++){var e=border==null?null:border.Element(N+edges[i]);string style=e==null?null:(string)e.Attribute("style");st.Border[i]=style=="none"?null:style;st.BorderColor[i]=Rgb(e==null?null:e.Element(N+"color"),"FF000000");}
                int id=(int?)xf.Attribute("numFmtId")??0;string code;st.Format=formats.TryGetValue(id,out code)?code:PdfFormat.BuiltIn(id);
                var a=xf.Element(N+"alignment");if(a!=null){st.Horizontal=(string)a.Attribute("horizontal")??"general";st.Vertical=(string)a.Attribute("vertical")??"bottom";st.Wrap=(string)a.Attribute("wrapText")=="1"||(string)a.Attribute("wrapText")=="true";}
                styles.Add(st);
            }
            if(styles.Count==0)styles.Add(new PdfStyle());
        }
        static int Column(string address){int n=0;foreach(char c in address){if(c<'A'||c>'Z')break;n=n*26+c-'A'+1;}return n;}
        static int RowOf(string address){return int.Parse(Regex.Match(address,@"\d+$").Value,Inv);}
        object CellValue(XElement c) {
            string type=(string)c.Attribute("t"),raw=(string)c.Element(N+"v");
            if(type=="inlineStr")return string.Concat(c.Descendants(N+"t").Select(t=>t.Value));
            if(raw==null)return null;
            if(type=="s"){int i;return int.TryParse(raw,out i)&&i>=0&&i<strings.Count?strings[i]:"";}
            if(type=="str"||type=="e")return raw;
            if(type=="b")return raw=="1";
            double d;return double.TryParse(raw,NumberStyles.Float,Inv,out d)?(object)d:raw;
        }
        public void Render(PdfDocument pdf,string sheetName,int area,double pw,double ph) {
            int index=sheets.FindIndex(s=>string.Equals(s.Item1,sheetName,StringComparison.OrdinalIgnoreCase));
            if(index<0)throw new InvalidDataException("в книге нет листа «"+sheetName+"».");
            var doc=Load(sheets[index].Item2);var root=doc.Root;
            // The print area of the sheet; the first range of «Итоги работы» is its first table.
            int c1=1,r1=1,c2=0,r2=0;var names=book.Root.Element(N+"definedNames");
            var print=names==null?null:names.Elements(N+"definedName").FirstOrDefault(e=>(string)e.Attribute("name")=="_xlnm.Print_Area"&&(int?)e.Attribute("localSheetId")==index);
            var ranges=print==null?new string[0]:print.Value.Split(',');
            var m=ranges.Length>area?Regex.Match(ranges[area],@"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)"):Match.Empty;
            var cells=new Dictionary<long,Tuple<object,PdfStyle>>();var rowInfo=new Dictionary<int,Tuple<double?,bool>>();
            foreach(var row in root.Element(N+"sheetData").Elements(N+"row")) {
                int rn=(int)row.Attribute("r");double ht;bool hidden=(string)row.Attribute("hidden")=="1"||(string)row.Attribute("hidden")=="true";
                rowInfo[rn]=Tuple.Create(double.TryParse((string)row.Attribute("ht"),NumberStyles.Float,Inv,out ht)?(double?)ht:null,hidden);
                foreach(var c in row.Elements(N+"c")){string address=(string)c.Attribute("r");int s=(int?)c.Attribute("s")??0;cells[(long)RowOf(address)<<16|(uint)Column(address)]=Tuple.Create(CellValue(c),styles[s<styles.Count?s:0]);}
            }
            if(m.Success){c1=Column(m.Groups[1].Value);r1=int.Parse(m.Groups[2].Value,Inv);c2=Column(m.Groups[3].Value);r2=int.Parse(m.Groups[4].Value,Inv);}
            else {foreach(long k in cells.Keys){c2=Math.Max(c2,(int)(k&0xFFFF));r2=Math.Max(r2,(int)(k>>16));}}
            // Column widths in Excel units become pixels with the maximum digit width of the normal font (96 dpi), then points.
            var format=root.Element(N+"sheetFormatPr");double defaultHeight=15,defaultWidth=8.43,value;
            if(format!=null){if(double.TryParse((string)format.Attribute("defaultRowHeight"),NumberStyles.Float,Inv,out value))defaultHeight=value;if(double.TryParse((string)format.Attribute("defaultColWidth"),NumberStyles.Float,Inv,out value))defaultWidth=value;}
            var normal=pdf.Font(normalFont,false,false);int digit=7;
            if(normal!=null){digit=1;for(char ch='0';ch<='9';ch++)digit=Math.Max(digit,(int)Math.Round(normal.Width(ch.ToString())*normalSize*96/72,MidpointRounding.AwayFromZero));}
            var cols=root.Element(N+"cols");var widths=new Dictionary<int,double>();var hiddenCols=new HashSet<int>();
            if(cols!=null)foreach(var col in cols.Elements(N+"col")){int min=(int)col.Attribute("min"),max=(int)col.Attribute("max");double w;bool hide=(string)col.Attribute("hidden")=="1"||(string)col.Attribute("hidden")=="true";
                for(int c=min;c<=Math.Min(max,16384);c++){if(double.TryParse((string)col.Attribute("width"),NumberStyles.Float,Inv,out w))widths[c]=w;if(hide)hiddenCols.Add(c);}}
            var x=new Dictionary<int,double>();var y=new Dictionary<int,double>();double width=0,height=0;
            for(int c=c1;c<=c2;c++){if(hiddenCols.Contains(c))continue;double w;if(!widths.TryGetValue(c,out w))w=defaultWidth;x[c]=width;width+=Math.Truncate((256*w+Math.Truncate(128.0/digit))/256*digit)*72/96;}
            for(int r=r1;r<=r2;r++){Tuple<double?,bool> info;rowInfo.TryGetValue(r,out info);if(info!=null&&info.Item2)continue;y[r]=height;height+=info!=null&&info.Item1.HasValue?info.Item1.Value:defaultHeight;}
            Func<int,double> right=c=>{for(int k=c+1;k<=c2+1;k++){double v;if(x.TryGetValue(k,out v))return v;}return width;};
            Func<int,double> bottom=r=>{for(int k=r+1;k<=r2+1;k++){double v;if(y.TryGetValue(k,out v))return v;}return height;};
            // Merged areas: the top-left cell carries the value; inner edges are not drawn.
            var merged=new Dictionary<long,int[]>();var mergedCells=new Dictionary<long,int[]>();var merges=root.Element(N+"mergeCells");
            if(merges!=null)foreach(var mc in merges.Elements(N+"mergeCell")){var p=Regex.Match((string)mc.Attribute("ref")??"",@"^([A-Z]+)(\d+):([A-Z]+)(\d+)$");if(!p.Success)continue;
                int[] rect={Column(p.Groups[1].Value),int.Parse(p.Groups[2].Value,Inv),Column(p.Groups[3].Value),int.Parse(p.Groups[4].Value,Inv)};merged[(long)rect[1]<<16|(uint)rect[0]]=rect;
                for(int r=rect[1];r<=rect[3];r++)for(int c=rect[0];c<=rect[2];c++)mergedCells[(long)r<<16|(uint)c]=rect;}
            // The page of the chosen paper, margins of the sheet in inches. Text and rows are scaled down to fit both ways;
            // the columns always take the whole width of the page (sx is never smaller than scale).
            double ml=18,mr=18,mt=25.2,mb=25.2;var margins=root.Element(N+"pageMargins");
            if(margins!=null){Func<string,double,double> inch=(name,fallback)=>{double v;return double.TryParse((string)margins.Attribute(name),NumberStyles.Float,Inv,out v)?v*72:fallback;};ml=inch("left",ml);mr=inch("right",mr);mt=inch("top",mt);mb=inch("bottom",mb);}
            double scale=Math.Min(1,Math.Min((pw-ml-mr)/Math.Max(1,width),(ph-mt-mb)/Math.Max(1,height)));
            double sx=(pw-ml-mr)/Math.Max(1,width);
            var page=pdf.AddPage(pw,ph);
            Func<double,double> X=v=>ml+v*sx;Func<double,double> Y=v=>ph-mt-v*scale;
            Func<int,int,PdfStyle> styleAt=(r,c)=>{Tuple<object,PdfStyle> t;return cells.TryGetValue((long)r<<16|(uint)c,out t)?t.Item2:null;};
            // Fills first, then text, then borders on top.
            foreach(int r in y.Keys)foreach(int c in x.Keys){
                long key=(long)r<<16|(uint)c;int[] rect;if(mergedCells.TryGetValue(key,out rect)&&(rect[0]!=c||rect[1]!=r))continue;
                var st=styleAt(r,c);if(st==null||st.Fill==null)continue;int er=rect==null?r:rect[3],ec=rect==null?c:rect[2];
                page.Rect(X(x[c]),Y(bottom(er)),(right(ec)-x[c])*sx,(bottom(er)-y[r])*scale,st.Fill);
            }
            foreach(int r in y.Keys)foreach(int c in x.Keys){
                long key=(long)r<<16|(uint)c;Tuple<object,PdfStyle> cell;if(!cells.TryGetValue(key,out cell)||cell.Item1==null)continue;
                int[] rect;if(mergedCells.TryGetValue(key,out rect)&&(rect[0]!=c||rect[1]!=r))continue;int er=rect==null?r:Math.Min(rect[3],r2),ec=rect==null?c:Math.Min(rect[2],c2);
                string text=PdfFormat.Text(cell.Item1,cell.Item2.Format);if(text=="")continue;
                var st=cell.Item2;var font=pdf.Font(st.Font,st.Bold,st.Italic);if(font==null)continue;
                double left=x[c],top=y[r],w=right(ec)-left,h=bottom(er)-top;
                string align=st.Horizontal;if(align=="general")align=cell.Item1 is double?"right":cell.Item1 is bool?"center":"left";if(align=="centerContinuous"||align=="distributed"||align=="justify")align="center";if(align=="fill")align="left";
                // Like Excel, left-aligned text that is not wrapped continues over empty cells to the right.
                if(rect==null&&!st.Wrap&&align=="left"){double need=(font.Width(text)*st.Size+3)*scale/sx;int next=c;
                    while(w<need){int k=next+1;while(k<=c2&&!x.ContainsKey(k))k++;if(k>c2)break;long nk=(long)r<<16|(uint)k;Tuple<object,PdfStyle> other;
                        if(mergedCells.ContainsKey(nk)||cells.TryGetValue(nk,out other)&&other.Item1!=null&&PdfFormat.Text(other.Item1,other.Item2.Format)!="")break;next=k;w=right(k)-left;}}
                page.Text(font,text,st.Size,st.Color,X(left),Y(top+h),w*sx,h*scale,align,st.Vertical,st.Wrap,scale);
            }
            var lines=new Dictionary<string,Tuple<double,double,double,double,string,string>>();
            Action<double,double,double,double,string,string> edge=(ax,ay,bx,by,style,color)=>{
                if(style==null)return;string id=Math.Round(ax,2)+"|"+Math.Round(ay,2)+"|"+Math.Round(bx,2)+"|"+Math.Round(by,2);Tuple<double,double,double,double,string,string> old;
                if(!lines.TryGetValue(id,out old)||PdfPage.Weight(style)>PdfPage.Weight(old.Item5))lines[id]=Tuple.Create(ax,ay,bx,by,style,color);
            };
            foreach(int r in y.Keys)foreach(int c in x.Keys){
                var st=styleAt(r,c);if(st==null)continue;int[] rect;mergedCells.TryGetValue((long)r<<16|(uint)c,out rect);
                double l=X(x[c]),rr=X(right(c)),t=Y(y[r]),b=Y(bottom(r));
                // A merged area is framed by its outer cells; when its last rows or columns are hidden, its last visible ones carry the frame.
                int firstC=c,lastC=c,firstR=r,lastR=r;if(rect!=null){firstC=x.Keys.Where(k=>k>=rect[0]&&k<=rect[2]).Min();lastC=x.Keys.Where(k=>k>=rect[0]&&k<=rect[2]).Max();firstR=y.Keys.Where(k=>k>=rect[1]&&k<=rect[3]).Min();lastR=y.Keys.Where(k=>k>=rect[1]&&k<=rect[3]).Max();}
                var rightStyle=rect==null?st:styleAt(r,rect[2])??st;var bottomStyle=rect==null?st:styleAt(rect[3],c)??st;
                if(c==firstC)edge(l,b,l,t,st.Border[0],st.BorderColor[0]);
                if(c==lastC)edge(rr,b,rr,t,rightStyle.Border[1],rightStyle.BorderColor[1]);
                if(r==firstR)edge(l,t,rr,t,st.Border[2],st.BorderColor[2]);
                if(r==lastR)edge(l,b,rr,b,bottomStyle.Border[3],bottomStyle.BorderColor[3]);
            }
            foreach(var line in lines.Values)page.Line(line.Item1,line.Item2,line.Item3,line.Item4,line.Item5,line.Item6,scale);
        }
    }

    // Excel number formats as a Russian Excel shows them: decimal comma, group space, dd.mm.yyyy dates.
    static class PdfFormat {
        static readonly CultureInfo Ru=CultureInfo.GetCultureInfo("ru-RU");
        public static string BuiltIn(int id) {
            switch(id){
                case 1:return "0";case 2:return "0.00";case 3:return "#,##0";case 4:return "#,##0.00";case 9:return "0%";case 10:return "0.00%";case 11:return "0.00E+00";
                case 12:return "# ?/?";case 13:return "# ??/??";case 14:return "dd.mm.yyyy";case 15:return "d-mmm-yy";case 16:return "d-mmm";case 17:return "mmm-yy";
                case 18:return "h:mm AM/PM";case 19:return "h:mm:ss AM/PM";case 20:return "h:mm";case 21:return "h:mm:ss";case 22:return "dd.mm.yyyy h:mm";
                case 37:return "#,##0 ;(#,##0)";case 38:return "#,##0 ;[Red](#,##0)";case 39:return "#,##0.00;(#,##0.00)";case 40:return "#,##0.00;[Red](#,##0.00)";
                case 45:return "mm:ss";case 46:return "[h]:mm:ss";case 47:return "mm:ss.0";case 48:return "##0.0E+0";case 49:return "@";default:return "General";
            }
        }
        static List<string> Sections(string code) {
            var list=new List<string>();var sb=new StringBuilder();bool quote=false;
            for(int i=0;i<code.Length;i++){char ch=code[i];if(ch=='"')quote=!quote;if(ch=='\\'&&i+1<code.Length&&!quote){sb.Append(ch).Append(code[++i]);continue;}if(ch==';'&&!quote){list.Add(sb.ToString());sb.Clear();continue;}sb.Append(ch);}
            list.Add(sb.ToString());return list;
        }
        // Literal text of a format: quoted strings, escaped characters and plain signs; brackets ([Red], [>=10]) are dropped.
        static string Plain(string code) {
            var sb=new StringBuilder();bool quote=false;
            for(int i=0;i<code.Length;i++){char ch=code[i];if(ch=='"'){quote=!quote;continue;}if(quote){sb.Append(ch);continue;}if(ch=='\\'&&i+1<code.Length){sb.Append(code[++i]);continue;}if(ch=='_'&&i+1<code.Length){i++;sb.Append(' ');continue;}if(ch=='*'&&i+1<code.Length){i++;continue;}if(ch=='['){int end=code.IndexOf(']',i);if(end>i){i=end;continue;}}sb.Append(ch);}
            return sb.ToString();
        }
        static bool IsDate(string code) {
            string plain=Regex.Replace(Regex.Replace(code,"\"[^\"]*\"|\\\\.",""),@"\[(?![hms]+\])[^\]]*\]","");
            return Regex.IsMatch(plain,"[dmyhsDMYHS]")&&!Regex.IsMatch(plain,"[0#?]");
        }
        public static string Text(object value,string code) {
            if(value is string)return (string)value;
            if(value is bool)return (bool)value?"ИСТИНА":"ЛОЖЬ";
            if(!(value is double))return Convert.ToString(value,CultureInfo.InvariantCulture);
            double d=(double)value;if(string.IsNullOrEmpty(code)||code.Equals("General",StringComparison.OrdinalIgnoreCase)||code=="@")return General(d);
            var sections=Sections(code);string section=sections[0];bool sign=d<0;
            if(d<0&&sections.Count>1){section=sections[1];sign=false;}else if(d==0&&sections.Count>2)section=sections[2];
            if(section.Equals("General",StringComparison.OrdinalIgnoreCase))return (sign?"-":"")+General(Math.Abs(d));
            if(IsDate(section))return Date(d,section);
            return (sign?"-":"")+Number(Math.Abs(d),section);
        }
        static string General(double d) {
            if(d==Math.Floor(d)&&Math.Abs(d)<1e11)return ((long)d).ToString(CultureInfo.InvariantCulture);
            string s=d.ToString("0.#########",Ru);if(s.Replace("-","").Replace(",","").Length>11)s=d.ToString("G10",Ru);return s;
        }
        static string Number(double d,string section) {
            string plain=Plain(section);int percents=plain.Count(ch=>ch=='%');for(int i=0;i<percents;i++)d*=100;
            var m=Regex.Match(plain,@"[0#?][0#?,]*(?:\.[0#?]*)?|\.[0#?]+");
            if(!m.Success)return plain;
            string pattern=m.Value;int dot=pattern.IndexOf('.');string integer=dot<0?pattern:pattern.Substring(0,dot),fraction=dot<0?"":pattern.Substring(dot+1);
            int max=fraction.Length,min=fraction.Count(ch=>ch=='0');bool group=integer.Contains(",");
            string digits=Math.Round(d,max,MidpointRounding.AwayFromZero).ToString((group?"#,0":"0")+(max>0?"."+new string('0',max):""),Ru);
            if(max>min){int comma=digits.LastIndexOf(Ru.NumberFormat.NumberDecimalSeparator,StringComparison.Ordinal);int keep=digits.Length-(max-min);while(digits.Length>keep&&digits[digits.Length-1]=='0')digits=digits.Substring(0,digits.Length-1);if(digits.EndsWith(Ru.NumberFormat.NumberDecimalSeparator,StringComparison.Ordinal))digits=digits.Substring(0,comma);}
            if(!integer.Contains("0")&&digits.StartsWith("0"+Ru.NumberFormat.NumberDecimalSeparator,StringComparison.Ordinal))digits=digits.Substring(1);
            return plain.Substring(0,m.Index)+digits+plain.Substring(m.Index+m.Length);
        }
        static string Date(double d,string section) {
            DateTime t;try{t=DateTime.FromOADate(d);}catch(ArgumentException){return General(d);}
            var sb=new StringBuilder();bool quote=false,hourSeen=false;var lower=section.ToLowerInvariant();
            for(int i=0;i<section.Length;) {
                char ch=section[i],lc=char.ToLowerInvariant(ch);
                if(ch=='"'){quote=!quote;i++;continue;}if(quote){sb.Append(ch);i++;continue;}
                if(ch=='\\'&&i+1<section.Length){sb.Append(section[i+1]);i+=2;continue;}
                if(ch=='['){int end=section.IndexOf(']',i);string inner=end>i?lower.Substring(i+1,end-i-1):"";
                    if(inner=="h"||inner=="hh"){sb.Append(((long)Math.Floor(d*24)).ToString(CultureInfo.InvariantCulture));hourSeen=true;}else if(inner=="mm"||inner=="m")sb.Append(((long)Math.Floor(d*1440)).ToString(CultureInfo.InvariantCulture));
                    i=end>i?end+1:i+1;continue;}
                if(lower.Substring(i).StartsWith("am/pm",StringComparison.Ordinal)){sb.Append(t.Hour<12?"AM":"PM");i+=5;continue;}
                int n=1;while(i+n<section.Length&&char.ToLowerInvariant(section[i+n])==lc)n++;
                if(lc=='y')sb.Append(n<=2?t.ToString("yy",CultureInfo.InvariantCulture):t.Year.ToString("0000",CultureInfo.InvariantCulture));
                else if(lc=='d')sb.Append(n==1?t.Day.ToString(CultureInfo.InvariantCulture):n==2?t.Day.ToString("00",CultureInfo.InvariantCulture):n==3?Ru.DateTimeFormat.GetAbbreviatedDayName(t.DayOfWeek):Ru.DateTimeFormat.GetDayName(t.DayOfWeek));
                else if(lc=='h'){int hour=lower.Contains("am/pm")?(t.Hour%12==0?12:t.Hour%12):t.Hour;sb.Append(n==1?hour.ToString(CultureInfo.InvariantCulture):hour.ToString("00",CultureInfo.InvariantCulture));hourSeen=true;}
                else if(lc=='s')sb.Append(n==1?t.Second.ToString(CultureInfo.InvariantCulture):t.Second.ToString("00",CultureInfo.InvariantCulture));
                else if(lc=='m') {
                    // m is minutes right after hours or before seconds, otherwise the month.
                    string rest=Regex.Replace(lower.Substring(i+n),"\"[^\"]*\"","");bool minutes=n<=2&&(hourSeen||Regex.IsMatch(rest,@"^[^dmyh]*s"));
                    if(minutes)sb.Append(n==1?t.Minute.ToString(CultureInfo.InvariantCulture):t.Minute.ToString("00",CultureInfo.InvariantCulture));
                    else sb.Append(n==1?t.Month.ToString(CultureInfo.InvariantCulture):n==2?t.Month.ToString("00",CultureInfo.InvariantCulture):n==3?Ru.DateTimeFormat.GetAbbreviatedMonthName(t.Month):Ru.DateTimeFormat.GetMonthName(t.Month));
                    hourSeen=false;
                }
                else {sb.Append(section,i,n);i+=n;continue;}
                i+=n;
            }
            return sb.ToString();
        }
    }

    // TrueType font with glyph widths and a subset for embedding (glyph numbers stay; unused glyphs are emptied).
    sealed class PdfFont {
        readonly byte[] data;readonly Dictionary<string,int[]> tables=new Dictionary<string,int[]>(StringComparer.Ordinal);
        readonly int unitsPerEm,numGlyphs,metrics,locaFormat;readonly Dictionary<char,int> glyphs=new Dictionary<char,int>();
        public readonly SortedDictionary<int,string> Used=new SortedDictionary<int,string>();
        public readonly string Name,Key;public readonly int Ascent,Descent,CapHeight;public readonly int[] Box;public readonly bool Italic;
        public string Resource;
        int U16(int o){return data[o]<<8|data[o+1];}
        int S16(int o){return (short)(data[o]<<8|data[o+1]);}
        long U32(int o){return (long)data[o]<<24|(long)data[o+1]<<16|(long)data[o+2]<<8|data[o+3];}
        public PdfFont(string path,string key) {
            data=File.ReadAllBytes(path);Key=key;
            if(U32(0)!=0x00010000&&U32(0)!=0x74727565)throw new InvalidDataException("не TrueType: "+path);
            int count=U16(4);for(int i=0;i<count;i++){int o=12+16*i;tables[Encoding.ASCII.GetString(data,o,4)]=new[]{(int)U32(o+8),(int)U32(o+12)};}
            foreach(string t in new[]{"head","hhea","hmtx","maxp","cmap","loca","glyf"})if(!tables.ContainsKey(t))throw new InvalidDataException("в шрифте нет таблицы "+t+": "+path);
            int head=tables["head"][0];unitsPerEm=U16(head+18);locaFormat=S16(head+50);Box=new[]{S16(head+36),S16(head+38),S16(head+40),S16(head+42)};Italic=(U16(head+44)&2)!=0;
            int hhea=tables["hhea"][0];Ascent=S16(hhea+4);Descent=S16(hhea+6);metrics=U16(hhea+34);numGlyphs=U16(tables["maxp"][0]+4);
            CapHeight=(int)(Ascent*0.7);int[] os2;if(tables.TryGetValue("OS/2",out os2)&&U16(os2[0])>=2&&os2[1]>=90)CapHeight=S16(os2[0]+88);
            Name=Path.GetFileNameWithoutExtension(path);int[] names;
            if(tables.TryGetValue("name",out names)){int o=names[0],n=U16(o+2),strings=o+U16(o+4);
                for(int i=0;i<n;i++){int r=o+6+12*i;if(U16(r+6)!=6)continue;int platform=U16(r),length=U16(r+8),offset=strings+U16(r+10);
                    string s=platform==3||platform==0?Encoding.BigEndianUnicode.GetString(data,offset,length):Encoding.ASCII.GetString(data,offset,length);s=Regex.Replace(s,@"[^A-Za-z0-9\-]","");if(s!=""){Name=s;break;}}}
        }
        // Unicode BMP map (format 4 of platform 3 or 0).
        int Map(char ch) {
            int cmap=tables["cmap"][0],n=U16(cmap+2),best=-1;
            for(int i=0;i<n;i++){int r=cmap+4+8*i,platform=U16(r),encoding=U16(r+2),o=cmap+(int)U32(r+4);if(U16(o)!=4)continue;if(platform==3&&encoding==1||platform==0){best=o;if(platform==3)break;}}
            if(best<0)return 0;int seg=U16(best+6)/2,ends=best+14,starts=ends+2*seg+2,deltas=starts+2*seg,ranges=deltas+2*seg;
            for(int i=0;i<seg;i++){int end=U16(ends+2*i);if(ch>end)continue;int start=U16(starts+2*i);if(ch<start)return 0;int delta=S16(deltas+2*i),range=U16(ranges+2*i);
                if(range==0)return (ch+delta)&0xFFFF;int g=U16(ranges+2*i+range+2*(ch-start));return g==0?0:(g+delta)&0xFFFF;}
            return 0;
        }
        public int Glyph(char ch){int g;if(!glyphs.TryGetValue(ch,out g)){g=Map(ch);if(g>=numGlyphs)g=0;glyphs[ch]=g;}return g;}
        int Advance(int glyph){int hmtx=tables["hmtx"][0];return U16(hmtx+4*Math.Min(glyph,metrics-1));}
        public int Units(int glyph){return (int)Math.Round(Advance(glyph)*1000.0/unitsPerEm);}
        public double Width(string text){double w=0;foreach(char ch in text)w+=Advance(Glyph(ch));return w/unitsPerEm;}
        public double Scaled(int v){return v*1000.0/unitsPerEm;}
        public string Encode(string text) {
            var sb=new StringBuilder();foreach(char ch in text){int g=Glyph(ch);if(!Used.ContainsKey(g))Used[g]=g==0?"\uFFFD":ch.ToString();sb.Append(g.ToString("X4",CultureInfo.InvariantCulture));}return sb.ToString();
        }
        int GlyphOffset(int g){int loca=tables["loca"][0];return locaFormat==0?U16(loca+2*g)*2:(int)U32(loca+4*g);}
        // Keeps the used glyphs, the glyphs they are built from and glyph 0; other glyphs become empty.
        public byte[] Subset() {
            var keep=new HashSet<int>(Used.Keys);keep.Add(0);var queue=new Queue<int>(keep);int glyf=tables["glyf"][0];
            while(queue.Count>0){int g=queue.Dequeue();int start=GlyphOffset(g),end=GlyphOffset(g+1);if(end<=start||S16(glyf+start)>=0)continue;int p=glyf+start+10;
                while(true){int flags=U16(p),component=U16(p+2);if(keep.Add(component))queue.Enqueue(component);p+=4+((flags&1)!=0?4:2);if((flags&8)!=0)p+=2;else if((flags&0x40)!=0)p+=4;else if((flags&0x80)!=0)p+=8;if((flags&0x20)==0)break;}}
            var glyfData=new MemoryStream();var loca=new List<long>();
            for(int g=0;g<numGlyphs;g++){loca.Add(glyfData.Length);if(!keep.Contains(g))continue;int start=GlyphOffset(g),end=GlyphOffset(g+1);if(end>start){glyfData.Write(data,glyf+start,end-start);while(glyfData.Length%4!=0)glyfData.WriteByte(0);}}
            loca.Add(glyfData.Length);
            var locaData=new byte[loca.Count*4];for(int i=0;i<loca.Count;i++){long v=loca[i];locaData[4*i]=(byte)(v>>24);locaData[4*i+1]=(byte)(v>>16);locaData[4*i+2]=(byte)(v>>8);locaData[4*i+3]=(byte)v;}
            var output=new SortedDictionary<string,byte[]>(StringComparer.Ordinal);
            foreach(string t in new[]{"OS/2","cmap","cvt ","fpgm","name","post","prep","hhea","hmtx","maxp"}){int[] e;if(tables.TryGetValue(t,out e)){var b=new byte[e[1]];Array.Copy(data,e[0],b,0,e[1]);output[t]=b;}}
            var headData=new byte[tables["head"][1]];Array.Copy(data,tables["head"][0],headData,0,headData.Length);for(int i=8;i<12;i++)headData[i]=0;headData[50]=0;headData[51]=1;
            output["head"]=headData;output["glyf"]=glyfData.ToArray();output["loca"]=locaData;
            var ms=new MemoryStream();Action<long,int> put=(v,bytes)=>{for(int i=bytes-1;i>=0;i--)ms.WriteByte((byte)(v>>(8*i)));};
            int tablesCount=output.Count,power=1,log=0;while(power*2<=tablesCount){power*=2;log++;}
            put(0x00010000,4);put(tablesCount,2);put(power*16,2);put(log,2);put(tablesCount*16-power*16,2);
            long offset=12+16*tablesCount;var order=output.ToList();
            foreach(var t in order){long sum=0;for(int i=0;i<t.Value.Length;i+=4){long w=0;for(int k=0;k<4;k++)w=w<<8|(i+k<t.Value.Length?t.Value[i+k]:(byte)0);sum=(sum+w)&0xFFFFFFFF;}
                ms.Write(Encoding.ASCII.GetBytes(t.Key),0,4);put(sum,4);put(offset,4);put(t.Value.Length,4);offset+=(t.Value.Length+3)/4*4;}
            foreach(var t in order){ms.Write(t.Value,0,t.Value.Length);while(ms.Length%4!=0)ms.WriteByte(0);}
            return ms.ToArray();
        }
    }

    sealed class PdfPage {
        static readonly CultureInfo Inv=CultureInfo.InvariantCulture;
        public readonly double Width,Height;public readonly StringBuilder Content=new StringBuilder();public readonly HashSet<PdfFont> Fonts=new HashSet<PdfFont>();
        public PdfPage(double width,double height){Width=width;Height=height;}
        static string F(double v){return Math.Round(v,3).ToString("0.###",Inv);}
        static string Color(string argb){if(argb==null||argb.Length<8)return "0 0 0";Func<int,string> part=i=>F(int.Parse(argb.Substring(i,2),NumberStyles.HexNumber,Inv)/255.0);return part(2)+" "+part(4)+" "+part(6);}
        public static double Weight(string style){switch(style){case "hair":return 0.25;case "thin":case "dashed":case "dotted":case "dashDot":case "dashDotDot":return 0.5;case "medium":case "mediumDashed":case "mediumDashDot":case "mediumDashDotDot":case "slantDashDot":return 1;case "thick":case "double":return 1.5;default:return 0.5;}}
        public void Rect(double x,double y,double w,double h,string fill){Content.Append("q ").Append(Color(fill)).Append(" rg ").Append(F(x)).Append(' ').Append(F(y)).Append(' ').Append(F(w)).Append(' ').Append(F(h)).Append(" re f Q\n");}
        public void Line(double ax,double ay,double bx,double by,string style,string color,double scale) {
            string dash=style=="dashed"||style=="mediumDashed"?"[2 1.5] 0 d":style=="dotted"?"[0.6 1] 0 d":style!=null&&style.Contains("DashDot")?"[2 1 0.6 1] 0 d":"[] 0 d";
            Content.Append("q ").Append(Color(color)).Append(" RG ").Append(F(Math.Max(0.25,Weight(style)*Math.Max(scale,0.5)))).Append(" w ").Append(dash).Append(' ').Append(F(ax)).Append(' ').Append(F(ay)).Append(" m ").Append(F(bx)).Append(' ').Append(F(by)).Append(" l S Q\n");
        }
        // Text inside a cell rectangle (x, y is the bottom-left corner), clipped to it; wrapped text breaks at spaces.
        public void Text(PdfFont font,string text,double size,string color,double x,double y,double w,double h,string align,string vertical,bool wrap,double scale) {
            Fonts.Add(font);double s=size*scale,pad=2.25*scale,lineHeight=s*1.15,ascent=font.Scaled(font.Ascent)/1000*s,descent=-font.Scaled(font.Descent)/1000*s;
            var lines=new List<string>();
            foreach(string paragraph in text.Replace("\r\n","\n").Split('\n')) {
                if(!wrap){lines.Add(paragraph);continue;}
                string line="";foreach(string word in paragraph.Split(' ')){string next=line==""?word:line+" "+word;if(line!=""&&font.Width(next)*s>w-2*pad){lines.Add(line);line=word;}else line=next;}lines.Add(line);
            }
            double block=lines.Count*lineHeight;double top=vertical=="top"?y+h-pad:vertical=="center"||vertical=="justify"||vertical=="distributed"?y+(h+block)/2:y+block+pad*0.5;
            Content.Append("q ").Append(F(x)).Append(' ').Append(F(y)).Append(' ').Append(F(w)).Append(' ').Append(F(h)).Append(" re W n BT ").Append(Color(color)).Append(" rg\n");
            for(int i=0;i<lines.Count;i++){
                string line=lines[i];double width=font.Width(line)*s,left=align=="right"?x+w-pad-width:align=="center"?x+(w-width)/2:x+pad;
                double baseline=top-i*lineHeight-(lineHeight-(ascent+descent))/2-ascent;
                Content.Append('/').Append(font.Resource).Append(' ').Append(F(s)).Append(" Tf 1 0 0 1 ").Append(F(left)).Append(' ').Append(F(baseline)).Append(" Tm <").Append(font.Encode(line)).Append("> Tj\n");
            }
            Content.Append("ET Q\n");
        }
    }

    sealed class PdfDocument {
        static readonly CultureInfo Inv=CultureInfo.InvariantCulture;
        readonly List<PdfPage> pages=new List<PdfPage>();readonly Dictionary<string,PdfFont> fonts=new Dictionary<string,PdfFont>(StringComparer.OrdinalIgnoreCase);readonly string title;
        public PdfDocument(string title){this.title=title;}
        public PdfPage AddPage(double width,double height){var p=new PdfPage(width,height);pages.Add(p);return p;}
        // Fonts of Windows (registry and the Fonts folder); metric-compatible free fonts where Windows fonts are absent.
        public PdfFont Font(string family,bool bold,bool italic) {
            string key=family+"|"+bold+"|"+italic;PdfFont font;if(fonts.TryGetValue(key,out font))return font;
            string path=PdfFonts.Find(family,bold,italic);
            if(path!=null)try{font=new PdfFont(path,key);font.Resource="F"+(fonts.Values.Count(f=>f!=null)+1);}catch(Exception){font=null;}
            if(font==null&&(bold||italic)){font=Font(family,false,false);}
            if(font==null&&!string.Equals(family,"Arial",StringComparison.OrdinalIgnoreCase))font=Font("Arial",bold,italic);
            fonts[key]=font;return font;
        }
        static byte[] Deflate(byte[] input) {
            var ms=new MemoryStream();ms.WriteByte(0x78);ms.WriteByte(0x9C);
            using(var d=new DeflateStream(ms,CompressionLevel.Optimal,true))d.Write(input,0,input.Length);
            uint a=1,b=0;foreach(byte x in input){a=(a+x)%65521;b=(b+a)%65521;}uint adler=b<<16|a;
            ms.WriteByte((byte)(adler>>24));ms.WriteByte((byte)(adler>>16));ms.WriteByte((byte)(adler>>8));ms.WriteByte((byte)adler);return ms.ToArray();
        }
        static string Hex16(string s){var sb=new StringBuilder("<FEFF");foreach(char ch in s)sb.Append(((int)ch).ToString("X4",Inv));return sb.Append('>').ToString();}
        public byte[] Save() {
            var objects=new List<byte[]>();Func<int> reserve=()=>{objects.Add(null);return objects.Count;};
            Action<int,string> set=(n,text)=>objects[n-1]=Encoding.ASCII.GetBytes(text);
            Action<int,string,byte[]> stream=(n,dict,body)=>{var head=Encoding.ASCII.GetBytes("<< "+dict+" /Length "+body.Length+" >>\nstream\n");var tail=Encoding.ASCII.GetBytes("\nendstream");var all=new byte[head.Length+body.Length+tail.Length];head.CopyTo(all,0);body.CopyTo(all,head.Length);tail.CopyTo(all,head.Length+body.Length);objects[n-1]=all;};
            int catalog=reserve(),tree=reserve(),info=reserve();var fontObjects=new Dictionary<PdfFont,int>();var kids=new List<int>();
            // Page contents first: they decide which glyphs every font needs.
            var contents=pages.Select(p=>Deflate(Encoding.ASCII.GetBytes(p.Content.ToString()))).ToList();
            int tag=0;
            foreach(var font in pages.SelectMany(p=>p.Fonts).Distinct()) {
                int type0=reserve(),cid=reserve(),descriptor=reserve(),file=reserve(),unicode=reserve();fontObjects[font]=type0;
                string name=new string(Enumerable.Range(0,6).Select(i=>(char)('A'+(tag/(int)Math.Pow(26,i))%26)).ToArray())+"+"+font.Name;tag++;
                byte[] subset=font.Subset();
                var widths=new StringBuilder();foreach(int g in font.Used.Keys)widths.Append(g).Append(" [").Append(font.Units(g)).Append("] ");
                var cmap=new StringBuilder("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n");
                var used=font.Used.ToList();for(int i=0;i<used.Count;i+=100){var chunk=used.Skip(i).Take(100).ToList();cmap.Append(chunk.Count).Append(" beginbfchar\n");foreach(var pair in chunk)cmap.Append('<').Append(pair.Key.ToString("X4",Inv)).Append("> <").Append(((int)pair.Value[0]).ToString("X4",Inv)).Append(">\n");cmap.Append("endbfchar\n");}
                cmap.Append("endcmap\nCMapName currentdict /CIDFont defineresource pop\nend\nend\n");
                set(type0,"<< /Type /Font /Subtype /Type0 /BaseFont /"+name+" /Encoding /Identity-H /DescendantFonts ["+cid+" 0 R] /ToUnicode "+unicode+" 0 R >>");
                set(cid,"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /"+name+" /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor "+descriptor+" 0 R /DW 1000 /W ["+widths+"] /CIDToGIDMap /Identity >>");
                Func<int,string> u=v=>Math.Round(font.Scaled(v)).ToString(Inv);
                set(descriptor,"<< /Type /FontDescriptor /FontName /"+name+" /Flags "+(font.Italic?96:32)+" /FontBBox ["+u(font.Box[0])+" "+u(font.Box[1])+" "+u(font.Box[2])+" "+u(font.Box[3])+"] /ItalicAngle "+(font.Italic?-12:0)+" /Ascent "+u(font.Ascent)+" /Descent "+u(font.Descent)+" /CapHeight "+u(font.CapHeight)+" /StemV 80 /FontFile2 "+file+" 0 R >>");
                stream(file,"/Filter /FlateDecode /Length1 "+subset.Length,Deflate(subset));
                stream(unicode,"/Filter /FlateDecode",Deflate(Encoding.ASCII.GetBytes(cmap.ToString())));
            }
            for(int i=0;i<pages.Count;i++){
                var p=pages[i];int page=reserve(),content=reserve();kids.Add(page);
                string resources=string.Join(" ",p.Fonts.Select(f=>"/"+f.Resource+" "+fontObjects[f]+" 0 R"));
                set(page,"<< /Type /Page /Parent "+tree+" 0 R /MediaBox [0 0 "+p.Width.ToString(Inv)+" "+p.Height.ToString(Inv)+"] /Resources << /Font << "+resources+" >> >> /Contents "+content+" 0 R >>");
                stream(content,"/Filter /FlateDecode",contents[i]);
            }
            set(catalog,"<< /Type /Catalog /Pages "+tree+" 0 R >>");
            set(tree,"<< /Type /Pages /Kids ["+string.Join(" ",kids.Select(k=>k+" 0 R"))+"] /Count "+kids.Count+" >>");
            set(info,"<< /Title "+Hex16(title)+" /Producer "+Hex16("Свод проверки "+typeof(PdfReport).Assembly.GetName().Version.ToString(3))+" >>");
            var ms=new MemoryStream();Action<string> write=t=>{var b=Encoding.ASCII.GetBytes(t);ms.Write(b,0,b.Length);};
            write("%PDF-1.7\n");ms.Write(new byte[]{0x25,0xE2,0xE3,0xCF,0xD3,0x0A},0,6);var offsets=new List<long>();
            for(int i=0;i<objects.Count;i++){offsets.Add(ms.Length);write((i+1)+" 0 obj\n");ms.Write(objects[i],0,objects[i].Length);write("\nendobj\n");}
            long xref=ms.Length;write("xref\n0 "+(objects.Count+1)+"\n0000000000 65535 f \n");foreach(long o in offsets)write(o.ToString("0000000000",Inv)+" 00000 n \n");
            write("trailer\n<< /Size "+(objects.Count+1)+" /Root "+catalog+" 0 R /Info "+info+" 0 R >>\nstartxref\n"+xref+"\n%%EOF\n");
            return ms.ToArray();
        }
    }

    static class PdfFonts {
        static readonly Dictionary<string,string[]> Windows=new Dictionary<string,string[]>(StringComparer.OrdinalIgnoreCase){
            {"Times New Roman",new[]{"times.ttf","timesbd.ttf","timesi.ttf","timesbi.ttf"}},{"Arial",new[]{"arial.ttf","arialbd.ttf","ariali.ttf","arialbi.ttf"}},
            {"Calibri",new[]{"calibri.ttf","calibrib.ttf","calibrii.ttf","calibriz.ttf"}},{"Segoe UI",new[]{"segoeui.ttf","segoeuib.ttf","segoeuii.ttf","segoeuiz.ttf"}},
            {"Verdana",new[]{"verdana.ttf","verdanab.ttf","verdanai.ttf","verdanaz.ttf"}},{"Tahoma",new[]{"tahoma.ttf","tahomabd.ttf","tahoma.ttf","tahomabd.ttf"}}};
        // Metric-compatible fonts for Linux test machines.
        static readonly Dictionary<string,string[]> Free=new Dictionary<string,string[]>(StringComparer.OrdinalIgnoreCase){
            {"Times New Roman",new[]{"liberation/LiberationSerif-Regular.ttf","liberation/LiberationSerif-Bold.ttf","liberation/LiberationSerif-Italic.ttf","liberation/LiberationSerif-BoldItalic.ttf"}},
            {"Arial",new[]{"liberation/LiberationSans-Regular.ttf","liberation/LiberationSans-Bold.ttf","liberation/LiberationSans-Italic.ttf","liberation/LiberationSans-BoldItalic.ttf"}},
            {"Calibri",new[]{"crosextra/Carlito-Regular.ttf","crosextra/Carlito-Bold.ttf","crosextra/Carlito-Italic.ttf","crosextra/Carlito-BoldItalic.ttf"}}};
        public static string Find(string family,bool bold,bool italic) {
            int index=(bold?1:0)+(italic?2:0);string style=bold&&italic?" Bold Italic":bold?" Bold":italic?" Italic":"";
            string folder=Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            foreach(var hive in new[]{Microsoft.Win32.Registry.LocalMachine,Microsoft.Win32.Registry.CurrentUser})try {
                using(var key=hive.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")){if(key==null)continue;
                    foreach(string name in key.GetValueNames())if(string.Equals(name,family+style+" (TrueType)",StringComparison.OrdinalIgnoreCase)){string file=key.GetValue(name) as string;if(string.IsNullOrEmpty(file))continue;
                        string path=Path.IsPathRooted(file)?file:Path.Combine(folder,file);if(File.Exists(path)&&path.EndsWith(".ttf",StringComparison.OrdinalIgnoreCase))return path;}}
            } catch(Exception){}
            string[] files;
            if(!string.IsNullOrEmpty(folder)&&Windows.TryGetValue(family,out files)&&File.Exists(Path.Combine(folder,files[index])))return Path.Combine(folder,files[index]);
            if(Free.TryGetValue(family,out files))foreach(string root in new[]{"/usr/share/fonts/truetype","/usr/local/share/fonts"}){string path=Path.Combine(root,files[index]);if(File.Exists(path))return path;}
            if(index==0&&File.Exists("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")&&!Windows.ContainsKey(family))return "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf";
            return null;
        }
    }
}
