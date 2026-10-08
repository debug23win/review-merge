using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ReviewMerge {
    public sealed partial class XlsxWriter {
        const string CalculationMarker="ReviewMerge.FirstDates.v3";
        const int HelperWidth=48;
        sealed class VolumeRecord {public string List="",Who;public List<string> Boxes=new List<string>();public double? Date;public int Issue;public bool IsVolume=true;}
        sealed class ReferenceLayout {
            public XDocument Previous;
            public int Helper,End,OldHelper,OldEnd,ManifestStart,ManifestEnd,BoxCount;
            public readonly SortedDictionary<int,double> Days=new SortedDictionary<int,double>();
            public readonly SortedDictionary<int,string> People=new SortedDictionary<int,string>();
            // Reviewer as written in the files -> person of the reference sheet; person -> row label of the surname statistics.
            public readonly Dictionary<string,string> Names=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            public readonly Dictionary<string,string> StatKeys=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            public readonly Dictionary<string,object> Expected=new Dictionary<string,object>(StringComparer.OrdinalIgnoreCase);
            public readonly List<VolumeRecord> Volumes=new List<VolumeRecord>();
            public readonly Dictionary<string,double?> Completed=new Dictionary<string,double?>();
            public readonly Dictionary<string,string> CompletedBy=new Dictionary<string,string>();
            public readonly List<VolumeRecord> BoxFiles=new List<VolumeRecord>();
        }
        ReferenceLayout referenceLayout;
        int HelperStart(XDocument doc) {
            var marker=doc==null?null:doc.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==1).Elements(N+"c").FirstOrDefault(e=>Text(Value(e)).StartsWith("ReviewMerge.FirstDates.v",StringComparison.Ordinal));
            return marker==null?0:Column((string)marker.Attribute("r"));
        }
        string NamesRange(int h){return "$"+MergeEngine.ExcelColumn(h+45)+"$7:$"+MergeEngine.ExcelColumn(h+47)+"$"+Math.Max(7,referenceLayout.Names.Count+6);}
        string StatKeyOf(string who){string key;return string.IsNullOrEmpty(who)?"":referenceLayout.StatKeys.TryGetValue(who,out key)?key:who;}
        string StatKeyOfRaw(object raw){string name=Rules.CleanName(raw),who;if(name=="")return "";if(!referenceLayout.Names.TryGetValue(name,out who))who=name;return StatKeyOf(who);}
        // A reviewer name from the files becomes a known person: the same spelling, or the same surname when the operator confirms it.
        string ResolvePerson(string raw,List<string> known,MergeResult result) {
            var exact=known.FirstOrDefault(k=>Rules.NameKey(k)==Rules.NameKey(raw));
            if(exact!=null)return exact;
            string surname=Rules.SurnameKey(raw);
            if(options.MatchSurnames) {
                foreach(string candidate in known.Where(k=>surname!=""&&Rules.SurnameKey(k)==surname).ToList())
                    if(options.SamePerson==null||options.SamePerson(raw,candidate)){result.PeopleMerges.Add("«"+raw+"» → «"+candidate+"»");return candidate;}
                return raw;
            }
            var bySurname=known.Where(k=>string.Equals(Rules.Surname(k),raw,StringComparison.OrdinalIgnoreCase)).ToList();
            return bySurname.Count==1?bySurname[0]:raw;
        }
        ReferenceLayout PrepareReference(MergeResult result,List<double?> dates,DateTime asOf,int minHelper) {
            string path=PathFor("Справка по томам");var layout=new ReferenceLayout{Previous=path==null?null:Xml(path)};var old=layout.Previous;layout.OldHelper=HelperStart(old);
            layout.OldEnd=layout.OldHelper>0?Convert.ToInt32(Value(At(old,1,layout.OldHelper+4))??43,Inv):43;
            if(old!=null) {
                foreach(var c in old.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==8).Elements(N+"c")) {
                    int col=Column((string)c.Attribute("r"));double? day=XlsxReader.DateSerial(Value(c));if(day.HasValue&&col>=2&&(layout.OldHelper==0||col<layout.OldHelper))layout.Days[col]=day.Value;
                }
                for(int r=12;r<=40;r+=2){string who=Rules.CleanName(Value(At(old,r,1)));if(who!="")layout.People[r]=who;}
                if(layout.OldHelper>0&&Text(Value(At(old,1,layout.OldHelper)))==CalculationMarker) {
                    int start=Convert.ToInt32(Value(At(old,1,layout.OldHelper+3))??48,Inv);
                    for(int r=44;r<start-2;r+=2){string who=Rules.CleanName(Value(At(old,r,1)));if(who!="")layout.People[r]=who;}
                    for(int r=start;r<=layout.OldEnd;r++){string key=Text(Value(At(old,r,1)))+"|"+Text(Value(At(old,r,2)));if(!key.EndsWith("|",StringComparison.Ordinal))layout.Expected[key]=Value(At(old,r,5));}
                }
            }
            this.referenceLayout=layout;
            var known=layout.People.Values.ToList();
            foreach(string raw in result.Rows.Select(r=>Rules.CleanName(r.Values[8])).Where(v=>v!="").Distinct(StringComparer.OrdinalIgnoreCase).ToList()) {
                string who=ResolvePerson(raw,known,result);layout.Names[raw]=who;
                if(!known.Contains(who,StringComparer.OrdinalIgnoreCase))known.Add(who);
                if(!layout.People.Values.Contains(who,StringComparer.OrdinalIgnoreCase)){int row=Enumerable.Range(0,15).Select(i=>12+i*2).FirstOrDefault(r=>!layout.People.ContainsKey(r));if(row==0)row=Math.Max(44,layout.People.Keys.DefaultIfEmpty(42).Max()+2);layout.People[row]=who;}
            }
            // Statistics show the surname; the full name only when two different people share it.
            var used=layout.Names.Values.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            foreach(string who in used){string surname=Rules.Surname(who);layout.StatKeys[who]=surname==""||used.Count(o=>Rules.SurnameKey(o)==Rules.SurnameKey(who))>1?who:surname;}
            for(int i=0;i<result.Rows.Count;i++){string raw=Rules.CleanName(result.Rows[i].Values[8]),who;if(!layout.Names.TryGetValue(raw,out who))who=raw;var boxes=Rules.BoxKeys(result.Rows[i].Values[10],result.Rows[i].Values[11]);layout.BoxFiles.Add(new VolumeRecord{Boxes=boxes,List=Rules.BoxList(boxes),Who=who,Date=dates[i]});}
            foreach(double day in dates.Where(v=>v.HasValue).Select(v=>v.Value).Concat(new[]{asOf.ToOADate()}).Distinct().OrderBy(v=>v))if(!layout.Days.Values.Contains(day)){int col=2;while(layout.Days.ContainsKey(col))col+=7;layout.Days[col]=day;}
            int existingEnd=old==null?1:old.Root.Element(N+"sheetData").Elements(N+"row").Elements(N+"c").Select(e=>Column((string)e.Attribute("r"))).Where(c=>layout.OldHelper==0||c<layout.OldHelper).DefaultIfEmpty(1).Max();
            layout.End=Math.Max(existingEnd,layout.Days.Keys.DefaultIfEmpty(2).Max()+6);
            layout.Helper=Math.Min(16384-HelperWidth,Math.Max(Math.Max(96,layout.End+3),minHelper));return layout;
        }
        string ChooseVolumeField(int h,int rn,List<int> indices,int field,bool reviewedOnly) {
            string d=MergeEngine.ExcelColumn(h+13)+rn,fallback=reviewedOnly?"\"\"":MergeEngine.ExcelColumn(h+field)+(indices[0]+7);
            if(indices.Count>60) {
                string match="MATCH(1,INDEX(("+LocalRange(h+28,7,indices.Max()+7)+"="+MergeEngine.ExcelColumn(h+12)+rn+")*("+LocalRange(h+4,7,indices.Max()+7)+"="+d+"),0),0)";
                return "IF("+d+"=\"\","+fallback+",IFERROR(INDEX("+LocalRange(h+field,7,indices.Max()+7)+","+match+"),"+fallback+"))";
            }
            string f=fallback;foreach(int i in indices.AsEnumerable().Reverse()){string date=MergeEngine.ExcelColumn(h+4)+(i+7);f="IF(AND("+date+"<>\"\","+date+"="+d+"),"+MergeEngine.ExcelColumn(h+field)+(i+7)+","+f+")";}return f;
        }
        // The first filled value among the records of a volume: its kind or boxes, also when the ИУЛ row has none.
        string FirstFilled(int h,int rn,List<int> indices,int field) {
            if(indices.Count>60) {
                int lastRow=indices.Max()+7;string values=LocalRange(h+field,7,lastRow);
                return "IFERROR(INDEX("+values+",MATCH(1,INDEX(("+LocalRange(h+28,7,lastRow)+"="+MergeEngine.ExcelColumn(h+12)+rn+")*("+values+"<>\"\"),0),0)),\"\")";
            }
            string f="\"\"";foreach(int i in indices.AsEnumerable().Reverse()){string c=MergeEngine.ExcelColumn(h+field)+(i+7);f="IF("+c+"<>\"\","+c+","+f+")";}return f;
        }
        void ReferenceVolume(Grid grid,int h,int rn,List<int> indices,MergeResult result,List<List<string>> keys,List<double?> dates,int issue,bool isVolume) {
            double? d=indices.All(i=>dates[i].HasValue)?(double?)indices.Max(i=>dates[i].Value):null;int chosen=d.HasValue?indices.First(i=>dates[i]==d):indices[0];string raw=d.HasValue?Rules.CleanName(result.Rows[chosen].Values[8]):"",who;
            if(!referenceLayout.Names.TryGetValue(raw,out who))who=raw;
            var boxes=indices.Select(i=>keys[i]).FirstOrDefault(k=>k.Count>0)??new List<string>();string list=Rules.BoxList(boxes),display=Rules.BoxListDisplay(list);
            string kind=indices.Select(i=>Text(result.Rows[i].Values[10])).FirstOrDefault(v=>v!="")??"";
            referenceLayout.Volumes.Add(new VolumeRecord{Boxes=boxes,List=list,Who=who,Date=d,Issue=issue,IsVolume=isVolume});
            grid.Set(rn,h+37,raw,bodyStyle,ChooseVolumeField(h,rn,indices,9,true));grid.Set(rn,h+38,kind,bodyStyle,FirstFilled(h,rn,indices,8));grid.Set(rn,h+39,list,bodyStyle,FirstFilled(h,rn,indices,0));
            string rawRef=MergeEngine.ExcelColumn(h+37)+rn,keyRef=MergeEngine.ExcelColumn(h+39)+rn,whoRef=MergeEngine.ExcelColumn(h+40)+rn,dRef=MergeEngine.ExcelColumn(h+13)+rn,volumeRef=MergeEngine.ExcelColumn(h+2)+rn,names=NamesRange(h);
            grid.Set(rn,h+40,who,bodyStyle,"IF("+rawRef+"=\"\",\"\",IFERROR(VLOOKUP("+rawRef+","+names+",2,FALSE),"+rawRef+"))");
            // One box is shown as "45пд"; boxes of a list cell are shown as read at the build, e.g. "45пд, 46пд".
            if(boxes.Count>1)grid.Set(rn,h+41,display,bodyStyle);
            else grid.Set(rn,h+41,display,bodyStyle,"IF("+keyRef+"=\"\",\"\",MID("+keyRef+",FIND(\"|\","+keyRef+")+1,LEN("+keyRef+")-FIND(\"|\","+keyRef+")-1)&LOWER(MID("+keyRef+",2,FIND(\"|\","+keyRef+")-2)))");
            int unique=isVolume&&d.HasValue&&list!=""&&who!=""&&!referenceLayout.Volumes.Take(referenceLayout.Volumes.Count-1).Any(v=>v.IsVolume&&v.List==list&&v.Who==who&&v.Date==d)?1:0;
            grid.Set(rn,h+42,unique,intStyle,"IF(OR("+volumeRef+"=0,"+dRef+"=\"\","+keyRef+"=\"\","+whoRef+"=\"\"),0,IF(COUNTIFS("+LocalRange(h+39,7,rn)+","+keyRef+","+LocalRange(h+40,7,rn)+","+whoRef+","+LocalRange(h+13,7,rn)+","+dRef+","+LocalRange(h+2,7,rn)+",1)=1,1,0))");
            grid.Set(rn,h+43,StatKeyOf(who),bodyStyle,"IF("+rawRef+"=\"\",\"\",IFERROR(VLOOKUP("+rawRef+","+names+",3,FALSE),"+rawRef+"))");
        }
        static int BoxNumber(string key){int n;int.TryParse(key.Substring(key.IndexOf('|')+1),NumberStyles.Integer,CultureInfo.InvariantCulture,out n);return n;}
        void WriteReference(Grid calculations,int h,int summaryReportEnd) {
            var l=referenceLayout;var g=new Grid();var old=l.Previous;
            if(old!=null)foreach(var row in old.Root.Element(N+"sheetData").Elements(N+"row")){int rn=(int)row.Attribute("r");g.Templates[rn]=new XElement(row);foreach(var c in row.Elements(N+"c")){int col=Column((string)c.Attribute("r"));if(l.OldHelper>0&&col>=l.OldHelper&&col<l.OldHelper+HelperWidth)continue;if(l.OldHelper>0&&rn>=44&&rn<=l.OldEnd&&col<=l.End)continue;if(!g.Rows.ContainsKey(rn))g.Rows[rn]=new SortedDictionary<int,XElement>();g.Rows[rn][col]=new XElement(c);}}
            foreach(var row in calculations.Rows)foreach(var cell in row.Value.Where(c=>c.Key>=h&&c.Key<h+HelperWidth)){if(!g.Rows.ContainsKey(row.Key))g.Rows[row.Key]=new SortedDictionary<int,XElement>();g.Rows[row.Key][cell.Key]=new XElement(cell.Value);}
            g.Set(6,h+45,"Проверяющий в файлах",bodyStyle);g.Set(6,h+46,"Проверяющий в справке",bodyStyle);g.Set(6,h+47,"Фамилия в статистике",bodyStyle);
            int n=7;foreach(var map in l.Names){g.Set(n,h+45,map.Key,bodyStyle);g.Set(n,h+46,map.Value,bodyStyle);g.Set(n++,h+47,StatKeyOf(map.Value),bodyStyle);}
            int last=l.Volumes.Count+6;string dates=LocalRange(h+13,7,last),keys=LocalRange(h+39,7,last),owners=LocalRange(h+40,7,last),display=LocalRange(h+41,7,last),flags=LocalRange(h+42,7,last),isv=","+LocalRange(h+2,7,last)+",1";
            int label=StyleAt(old,12,1,bodyStyle),num=StyleAt(old,13,7,intStyle),boxStyle=StyleAt(old,12,2,bodyStyle),countStyle=StyleAt(old,13,2,intStyle),head=StyleAt(old,9,2,headStyle);
            if(old==null){g.Set(1,1,"Коробов получено",bodyStyle);g.Set(2,1,"Томов по проекту акта",bodyStyle);g.Set(8,1,"ФИО проверяющего",headStyle);}
            g.Set(3,1,"ИТОГО коробов проверить",bodyStyle);g.Set(4,1,"ИТОГО томов проверить",bodyStyle);g.Set(42,1,"ИТОГО проверено коробов",StyleAt(old,42,1,headStyle));g.Set(43,1,"ИТОГО проверено томов",StyleAt(old,43,1,headStyle));
            var known=l.BoxFiles.SelectMany(v=>v.Boxes).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v=>v.Substring(0,v.IndexOf('|')),StringComparer.Ordinal).ThenBy(BoxNumber).ToList();
            l.BoxCount=known.Count;l.ManifestStart=Math.Max(48,l.People.Keys.DefaultIfEmpty(40).Max()+6);l.ManifestEnd=l.ManifestStart+Math.Max(1,known.Count)-1;
            int title=l.ManifestStart-2,header=l.ManifestStart-1;g.Set(title,1,"УЧЁТ ПРОВЕРЕННЫХ КОРОБОВ",headStyle);
            string[] headers={"Вид","Короб","Томов в реестре","Проверено томов","Всего томов по описи","Дата проверки короба","Состояние","Проверил короб"};for(int c=0;c<headers.Length;c++)g.Set(header,c+1,headers[c],headStyle);g.Heights[header]=75;
            var completion=l.Completed;var completedBy=l.CompletedBy;int fileLast=l.BoxFiles.Count+6;string fileDates=LocalRange(h+4,7,fileLast),fileKeys=LocalRange(h,7,fileLast),fileOwners=LocalRange(h+9,7,fileLast);
            for(int i=0;i<known.Count;i++) {
                int row=l.ManifestStart+i;string key=known[i];var members=l.Volumes.Where(v=>v.IsVolume&&v.Boxes.Contains(key)).ToList();int checkedCount=members.Count(v=>v.Date.HasValue);object expected;l.Expected.TryGetValue(key,out expected);
                var checkedFiles=l.BoxFiles.Where(v=>v.Boxes.Contains(key)&&v.Date.HasValue&&v.Who!="").ToList();double? date=checkedFiles.Count>0?(double?)checkedFiles.Min(v=>v.Date.Value):null;completion[key]=date;string who=date.HasValue?checkedFiles.First(v=>v.Date==date).Who:"";completedBy[key]=who;
                // A file belongs to every box of its list: ";ПД|45;ПД|46;" contains ";ПД|45;".
                string k="$A"+row+"&\"|\"&TEXT($B"+row+",\"0\")",inList="ISNUMBER(SEARCH(\";\"&"+k+"&\";\","+fileKeys+"))",volumeCriteria="\"*;\"&"+k+"&\";*\"";
                g.Set(row,1,key.Split('|')[0],bodyStyle);g.Set(row,2,(double)BoxNumber(key),intStyle);g.Set(row,3,members.Count,intStyle,"COUNTIFS("+keys+","+volumeCriteria+isv+")");g.Set(row,4,checkedCount,intStyle,"COUNTIFS("+keys+","+volumeCriteria+","+dates+",\">0\""+isv+")");g.Set(row,5,expected,StyleAt(old,row,5,intStyle));
                g.Set(row,6,date.HasValue?(object)date.Value:"",dateStyle,"IFERROR(_xlfn.AGGREGATE(15,6,"+fileDates+"/"+inList+"/("+fileDates+">0),1),\"\")");
                g.Set(row,7,date.HasValue?"Проверен":"Нет даты и фамилии",noteStyle,"IF(F"+row+"<>\"\",\"Проверен\",\"Нет даты и фамилии\")");g.Heights[row]=36;
                string rawOwner="INDEX("+fileOwners+",MATCH(1,INDEX("+inList+"*("+fileDates+"=F"+row+"),0),0))";
                g.Set(row,8,who,bodyStyle,"IF(F"+row+"=\"\",\"\",IFERROR(VLOOKUP("+rawOwner+","+NamesRange(h)+",2,FALSE),IFERROR("+rawOwner+",\"\")))");
            }
            string completed=LocalRange(6,l.ManifestStart,l.ManifestEnd),finishedWho=LocalRange(8,l.ManifestStart,l.ManifestEnd);var boxTotals=new List<string>();var volumeTotals=new List<string>();
            foreach(var day in l.Days) {
                int start=day.Key;double serial=day.Value;string d="$"+MergeEngine.ExcelColumn(start)+"$8";g.Set(8,start,serial,StyleAt(old,8,start,dateHeadStyle));
                for(int slot=0;slot<5;slot++){g.Set(9,start+slot,"Короб, вид",StyleAt(old,9,start+slot,head));g.Set(10,start+slot,"Томов за сутки",StyleAt(old,10,start+slot,head));g.Set(11,start+slot,slot+1,StyleAt(old,11,start+slot,intStyle));}
                g.Set(9,start+5,"Коробов / томов",StyleAt(old,9,start+5,head));g.Set(9,start+6,"Примечание",StyleAt(old,9,start+6,head));
                foreach(var person in l.People) {
                    int row=person.Key;string who=person.Value;g.Set(row,1,who,label);var items=l.Volumes.Where(v=>v.IsVolume&&v.Date==serial&&v.Who==who).ToList();var boxList=items.Where(v=>v.List!="").Select(v=>v.List).Distinct().ToList();string personRef="$A"+row;
                    for(int slot=0;slot<5;slot++) {
                        int col=start+slot;string cell=MergeEngine.ExcelColumn(col)+row;string list=boxList.Count>slot?boxList[slot]:"";
                        g.Set(row,col,Rules.BoxListDisplay(list),boxStyle,"IFERROR(INDEX("+display+",_xlfn.AGGREGATE(15,6,(ROW("+dates+")-6)/(("+dates+"="+d+")*("+owners+"="+personRef+")*("+flags+"=1)),"+(slot+1)+")),\"\")");
                        g.Set(row+1,col,list==""?(object)"":items.Count(v=>v.List==list),countStyle,"IF("+cell+"=\"\",\"\",COUNTIFS("+display+","+cell+","+dates+","+d+","+owners+","+personRef+isv+"))");
                    }
                    g.Set(row,start+5,completion.Count(v=>v.Value==serial&&completedBy[v.Key]==who),num,"COUNTIFS("+completed+","+d+","+finishedWho+","+personRef+")");g.Set(row+1,start+5,items.Count,num,"COUNTIFS("+dates+","+d+","+owners+","+personRef+isv+")");
                }
                int totalCol=start+5;g.Set(42,totalCol,completion.Count(v=>v.Value==serial),StyleAt(old,42,totalCol,num),"COUNTIF("+completed+","+d+")");g.Set(43,totalCol,l.Volumes.Count(v=>v.IsVolume&&v.Date==serial),StyleAt(old,43,totalCol,num),"COUNTIFS("+dates+","+d+isv+")");boxTotals.Add(MergeEngine.ExcelColumn(totalCol)+"42");volumeTotals.Add(MergeEngine.ExcelColumn(totalCol)+"43");
            }
            object receivedBoxes=g.Rows.ContainsKey(1)&&g.Rows[1].ContainsKey(2)?Value(g.Rows[1][2]):null,receivedVolumes=g.Rows.ContainsKey(2)&&g.Rows[2].ContainsKey(2)?Value(g.Rows[2][2]):null;
            g.Set(3,2,receivedBoxes is double?(object)((double)receivedBoxes-completion.Count(v=>v.Value.HasValue&&l.Days.ContainsValue(v.Value.Value))):"",intStyle,"IF(ISNUMBER(B1),B1-SUM("+string.Join(",",boxTotals)+"),\"\")");
            g.Set(4,2,receivedVolumes is double?(object)((double)receivedVolumes-l.Volumes.Count(v=>v.IsVolume&&v.Date.HasValue&&l.Days.ContainsValue(v.Date.Value))):"",intStyle,"IF(ISNUMBER(B2),B2-SUM("+string.Join(",",volumeTotals)+"),\"\")");
            g.Set(1,h,CalculationMarker,bodyStyle);g.Set(1,h+2,summaryReportEnd,intStyle);g.Set(1,h+3,l.ManifestStart,intStyle);g.Set(1,h+4,l.ManifestEnd,intStyle);
            var doc=old==null?Sheet(g,new[]{Col(1,1,32),Col(2,l.End,14)},false,l.End):new XDocument(old);if(old!=null)doc.Root.Element(N+"sheetData").ReplaceWith(g.Data());
            doc.Root.Element(N+"dimension").SetAttributeValue("ref","A1:"+MergeEngine.ExcelColumn(h+HelperWidth-1)+g.Rows.Keys.Max());var cols=doc.Root.Element(N+"cols");if(cols==null){cols=new XElement(N+"cols");doc.Root.Element(N+"sheetData").AddBeforeSelf(cols);}var visible=cols.Elements().Where(c=>l.OldHelper==0||(int)c.Attribute("min")<l.OldHelper).Select(c=>new XElement(c)).ToList();foreach(var c in visible)if(l.OldHelper>0&&(int)c.Attribute("max")>=l.OldHelper)c.SetAttributeValue("max",l.OldHelper-1);cols.RemoveNodes();cols.Add(visible);for(int c=2;c<=l.End;c++)if(!visible.Any(v=>(int)v.Attribute("min")<=c&&(int)v.Attribute("max")>=c))cols.Add(Col(c,c,14));cols.Add(Col(h,h+HelperWidth-1,18,true));
            var merges=doc.Root.Element(N+"mergeCells");if(merges==null){merges=new XElement(N+"mergeCells");doc.Root.Element(N+"sheetData").AddAfterSelf(merges);}foreach(var day in l.Days){string range=MergeEngine.ExcelColumn(day.Key)+"8:"+MergeEngine.ExcelColumn(day.Key+6)+"8";if(!merges.Elements().Any(m=>(string)m.Attribute("ref")==range))merges.Add(new XElement(N+"mergeCell",new XAttribute("ref",range)));}merges.SetAttributeValue("count",merges.Elements().Count());PutSheet("Справка по томам",doc);
        }
        static string ReferenceFormula(string formula,int h) {
            return Regex.Replace(formula,@"(?<![A-Z0-9_!])\$?([A-Z]{1,3})\$?\d+(?::\$?([A-Z]{1,3})\$?\d+)?",m=>Column(m.Groups[1].Value)>=h&&Column(m.Groups[1].Value)<h+HelperWidth?"'Справка по томам'!"+m.Value:m.Value);
        }
    }
}
