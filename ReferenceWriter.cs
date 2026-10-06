using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ReviewMerge {
    public sealed partial class XlsxWriter {
        const string CalculationMarker="ReviewMerge.FirstDates.v3";
        const int HelperWidth=47;
        sealed class VolumeRecord {public string Box,Who;public double? Date;public int Issue;}
        sealed class ReferenceLayout {
            public XDocument Previous;
            public int Helper,End,OldHelper,OldEnd,ManifestStart,ManifestEnd;
            public readonly SortedDictionary<int,double> Days=new SortedDictionary<int,double>();
            public readonly SortedDictionary<int,string> People=new SortedDictionary<int,string>();
            public readonly Dictionary<string,string> Names=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            public readonly Dictionary<string,object> Expected=new Dictionary<string,object>(StringComparer.OrdinalIgnoreCase);
            public readonly List<VolumeRecord> Volumes=new List<VolumeRecord>();
            public readonly Dictionary<string,double?> Completed=new Dictionary<string,double?>();
            public readonly Dictionary<string,string> CompletedBy=new Dictionary<string,string>();
        }
        ReferenceLayout referenceLayout;
        static string Surname(string value){return Regex.Split(value.Trim(),@"\s+")[0];}
        int HelperStart(XDocument doc) {
            var marker=doc==null?null:doc.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==1).Elements(N+"c").FirstOrDefault(e=>Text(Value(e)).StartsWith("ReviewMerge.FirstDates.v",StringComparison.Ordinal));
            return marker==null?0:Column((string)marker.Attribute("r"));
        }
        ReferenceLayout PrepareReference(MergeResult result,List<double?> dates,DateTime asOf) {
            string path=PathFor("Справка по томам");var layout=new ReferenceLayout{Previous=path==null?null:Xml(path)};var old=layout.Previous;layout.OldHelper=HelperStart(old);
            layout.OldEnd=layout.OldHelper>0?Convert.ToInt32(Value(At(old,1,layout.OldHelper+4))??43,Inv):43;
            if(old!=null) {
                foreach(var c in old.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==8).Elements(N+"c")) {
                    int col=Column((string)c.Attribute("r"));double? day=XlsxReader.DateSerial(Value(c));if(day.HasValue&&col>=2&&(layout.OldHelper==0||col<layout.OldHelper))layout.Days[col]=day.Value;
                }
                for(int r=12;r<=40;r+=2){string who=Text(Value(At(old,r,1)));if(who!="")layout.People[r]=who;}
                if(layout.OldHelper>0&&Text(Value(At(old,1,layout.OldHelper)))==CalculationMarker) {
                    int start=Convert.ToInt32(Value(At(old,1,layout.OldHelper+3))??48,Inv);
                    for(int r=44;r<start-2;r+=2){string who=Text(Value(At(old,r,1)));if(who!="")layout.People[r]=who;}
                    for(int r=start;r<=layout.OldEnd;r++){string key=Text(Value(At(old,r,1)))+"|"+Text(Value(At(old,r,2)));if(!key.EndsWith("|",StringComparison.Ordinal))layout.Expected[key]=Value(At(old,r,5));}
                }
            }
            foreach(string raw in result.Rows.Select(r=>Text(r.Values[8])).Where(v=>v!="").Distinct(StringComparer.OrdinalIgnoreCase)) {
                var exact=layout.People.Values.Where(v=>string.Equals(v,raw,StringComparison.OrdinalIgnoreCase)).ToList();var candidates=exact.Count==1?exact:layout.People.Values.Where(v=>string.Equals(Surname(v),raw,StringComparison.OrdinalIgnoreCase)).ToList();
                string who=candidates.Count==1?candidates[0]:raw;layout.Names[raw]=who;
                if(!layout.People.Values.Contains(who,StringComparer.OrdinalIgnoreCase)){int row=Enumerable.Range(0,15).Select(i=>12+i*2).FirstOrDefault(r=>!layout.People.ContainsKey(r));if(row==0)row=Math.Max(44,layout.People.Keys.DefaultIfEmpty(42).Max()+2);layout.People[row]=who;}
            }
            foreach(double day in dates.Where(v=>v.HasValue).Select(v=>v.Value).Concat(new[]{asOf.ToOADate()}).Distinct().OrderBy(v=>v))if(!layout.Days.Values.Contains(day)){int col=2;while(layout.Days.ContainsKey(col))col+=7;layout.Days[col]=day;}
            int existingEnd=old==null?1:old.Root.Element(N+"sheetData").Elements(N+"row").Elements(N+"c").Select(e=>Column((string)e.Attribute("r"))).Where(c=>layout.OldHelper==0||c<layout.OldHelper).DefaultIfEmpty(1).Max();
            layout.End=Math.Max(existingEnd,layout.Days.Keys.DefaultIfEmpty(2).Max()+6);layout.Helper=Math.Max(96,layout.End+3);return layout;
        }
        string ChooseVolumeField(int h,int rn,List<int> indices,int field,bool reviewedOnly) {
            string d=MergeEngine.ExcelColumn(h+13)+rn,fallback=reviewedOnly?"\"\"":MergeEngine.ExcelColumn(h+field)+(indices[0]+7);
            if(indices.Count>60) {
                string match="MATCH(1,INDEX(("+LocalRange(h+28,7,indices.Max()+7)+"="+MergeEngine.ExcelColumn(h+12)+rn+")*("+LocalRange(h+4,7,indices.Max()+7)+"="+d+"),0),0)";
                return "IF("+d+"=\"\","+fallback+",IFERROR(INDEX("+LocalRange(h+field,7,indices.Max()+7)+","+match+"),"+fallback+"))";
            }
            string f=fallback;foreach(int i in indices.AsEnumerable().Reverse()){string date=MergeEngine.ExcelColumn(h+4)+(i+7);f="IF(AND("+date+"<>\"\","+date+"="+d+"),"+MergeEngine.ExcelColumn(h+field)+(i+7)+","+f+")";}return f;
        }
        void ReferenceVolume(Grid grid,int h,int rn,List<int> indices,MergeResult result,List<string> keys,List<double?> dates,int issue) {
            double? d=indices.All(i=>dates[i].HasValue)?(double?)indices.Max(i=>dates[i].Value):null;int chosen=d.HasValue?indices.First(i=>dates[i]==d):indices[0];string raw=d.HasValue?Text(result.Rows[chosen].Values[8]):"",who;
            if(!referenceLayout.Names.TryGetValue(raw,out who))who=raw;
            string key=keys[chosen],display=key==""?"":key.Split('|')[1]+key.Split('|')[0].ToLowerInvariant();referenceLayout.Volumes.Add(new VolumeRecord{Box=key,Who=who,Date=d,Issue=issue});
            grid.Set(rn,h+37,raw,bodyStyle,ChooseVolumeField(h,rn,indices,9,true));grid.Set(rn,h+38,Text(result.Rows[chosen].Values[10]),bodyStyle,ChooseVolumeField(h,rn,indices,8,false));grid.Set(rn,h+39,key,bodyStyle,ChooseVolumeField(h,rn,indices,0,false));
            string rawRef=MergeEngine.ExcelColumn(h+37)+rn,keyRef=MergeEngine.ExcelColumn(h+39)+rn,whoRef=MergeEngine.ExcelColumn(h+40)+rn,dRef=MergeEngine.ExcelColumn(h+13)+rn;
            grid.Set(rn,h+40,who,bodyStyle,"IF("+rawRef+"=\"\",\"\",IFERROR(VLOOKUP("+rawRef+",$"+MergeEngine.ExcelColumn(h+45)+"$7:$"+MergeEngine.ExcelColumn(h+46)+"$"+(referenceLayout.Names.Count+6)+",2,FALSE),"+rawRef+"))");
            grid.Set(rn,h+41,display,bodyStyle,"IF("+keyRef+"=\"\",\"\",MID("+keyRef+",FIND(\"|\","+keyRef+")+1,99)&LOWER(LEFT("+keyRef+",FIND(\"|\","+keyRef+")-1)))");
            int unique=referenceLayout.Volumes.Take(referenceLayout.Volumes.Count-1).Any(v=>v.Box==key&&v.Who==who&&v.Date==d)?0:1;
            grid.Set(rn,h+42,d.HasValue&&key!=""&&who!=""?unique:0,intStyle,"IF(OR("+dRef+"=\"\","+keyRef+"=\"\","+whoRef+"=\"\"),0,IF(COUNTIFS("+LocalRange(h+39,7,rn)+","+keyRef+","+LocalRange(h+40,7,rn)+","+whoRef+","+LocalRange(h+13,7,rn)+","+dRef+")=1,1,0))");
            grid.Set(rn,h+43,Surname(who),bodyStyle,"LEFT(TRIM("+whoRef+"),FIND(\" \",TRIM("+whoRef+")&\" \")-1)");
        }
        void WriteReference(Grid calculations,int h,int summaryReportEnd) {
            var l=referenceLayout;var g=new Grid();var old=l.Previous;
            if(old!=null)foreach(var row in old.Root.Element(N+"sheetData").Elements(N+"row")){int rn=(int)row.Attribute("r");g.Templates[rn]=new XElement(row);foreach(var c in row.Elements(N+"c")){int col=Column((string)c.Attribute("r"));if(l.OldHelper>0&&col>=l.OldHelper&&col<l.OldHelper+HelperWidth)continue;if(l.OldHelper>0&&rn>=44&&rn<=l.OldEnd&&col<=l.End)continue;if(!g.Rows.ContainsKey(rn))g.Rows[rn]=new SortedDictionary<int,XElement>();g.Rows[rn][col]=new XElement(c);}}
            foreach(var row in calculations.Rows)foreach(var cell in row.Value.Where(c=>c.Key>=h&&c.Key<h+HelperWidth)){if(!g.Rows.ContainsKey(row.Key))g.Rows[row.Key]=new SortedDictionary<int,XElement>();g.Rows[row.Key][cell.Key]=new XElement(cell.Value);}
            int n=7;foreach(var map in l.Names){g.Set(n,h+45,map.Key,bodyStyle);g.Set(n++,h+46,map.Value,bodyStyle);}
            int last=l.Volumes.Count+6;string dates=LocalRange(h+13,7,last),keys=LocalRange(h+39,7,last),owners=LocalRange(h+40,7,last),display=LocalRange(h+41,7,last),flags=LocalRange(h+42,7,last);
            int label=StyleAt(old,12,1,bodyStyle),num=StyleAt(old,13,7,intStyle),boxStyle=StyleAt(old,12,2,bodyStyle),countStyle=StyleAt(old,13,2,intStyle),head=StyleAt(old,9,2,headStyle);
            if(old==null){g.Set(1,1,"Коробов получено",bodyStyle);g.Set(2,1,"Томов по проекту акта",bodyStyle);g.Set(8,1,"ФИО проверяющего",headStyle);}
            g.Set(3,1,"ИТОГО коробов проверить",bodyStyle);g.Set(4,1,"ИТОГО томов проверить",bodyStyle);g.Set(42,1,"ИТОГО полностью проверено коробов",StyleAt(old,42,1,headStyle));g.Set(43,1,"ИТОГО проверено томов",StyleAt(old,43,1,headStyle));
            l.ManifestStart=Math.Max(48,l.People.Keys.DefaultIfEmpty(40).Max()+6);var known=l.Volumes.Where(v=>v.Box!="").Select(v=>v.Box).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v=>v,StringComparer.Ordinal).ToList();l.ManifestEnd=l.ManifestStart+Math.Max(1,known.Count)-1;
            int title=l.ManifestStart-2,header=l.ManifestStart-1;g.Set(title,1,"КОНТРОЛЬ ПОЛНОЙ ПРОВЕРКИ КОРОБОВ",headStyle);
            string[] headers={"Вид","Короб","Томов в реестре","Проверено томов","Всего томов по описи","Дата полной проверки","Состояние","Завершил проверку"};for(int c=0;c<headers.Length;c++)g.Set(header,c+1,headers[c],headStyle);g.Heights[header]=75;
            var completion=l.Completed;var completedBy=l.CompletedBy;bool mapped=l.Volumes.All(v=>v.Box!="");
            for(int i=0;i<known.Count;i++) {
                int row=l.ManifestStart+i;string key=known[i];var members=l.Volumes.Where(v=>v.Box==key).ToList();int checkedCount=members.Count(v=>v.Date.HasValue);object expected;l.Expected.TryGetValue(key,out expected);double total;bool supplied=double.TryParse(Text(expected),NumberStyles.Float,Inv,out total)&&total>0;
                double? date=checkedCount==members.Count&&((supplied&&total==members.Count)||(!supplied&&mapped))?(double?)members.Max(v=>v.Date.Value):null;completion[key]=date;string who=date.HasValue?members.First(v=>v.Date==date).Who:"";completedBy[key]=who;
                string k="$A"+row+"&\"|\"&TEXT($B"+row+",\"0\")";g.Set(row,1,key.Split('|')[0],bodyStyle);g.Set(row,2,double.Parse(key.Split('|')[1],Inv),intStyle);g.Set(row,3,members.Count,intStyle,"COUNTIF("+keys+","+k+")");g.Set(row,4,checkedCount,intStyle,"COUNTIFS("+keys+","+k+","+dates+",\">0\")");g.Set(row,5,expected,StyleAt(old,row,5,intStyle));
                g.Set(row,6,date.HasValue?(object)date.Value:"",dateStyle,"IF(AND(C"+row+">0,D"+row+"=C"+row+",OR(AND(ISNUMBER(E"+row+"),E"+row+"=C"+row+"),AND(E"+row+"=\"\",COUNTIF("+keys+",\"\")=0))),IFERROR(_xlfn.AGGREGATE(14,6,"+dates+"/("+keys+"="+k+")/("+dates+">0),1),\"\"),\"\")");
                g.Set(row,7,date.HasValue?"Проверен полностью":checkedCount<members.Count?"Есть непроверенные тома":"Полный состав не подтвержден",noteStyle,"IF(F"+row+"<>\"\",\"Проверен полностью\",IF(D"+row+"<C"+row+",\"Есть непроверенные тома\",\"Полный состав не подтвержден\"))");
                g.Set(row,8,who,bodyStyle,"IF(F"+row+"=\"\",\"\",IFERROR(INDEX("+owners+",MATCH(1,INDEX(("+keys+"="+k+")*("+dates+"=F"+row+"),0),0)),\"\"))");
            }
            string completed=LocalRange(6,l.ManifestStart,l.ManifestEnd),finishedWho=LocalRange(8,l.ManifestStart,l.ManifestEnd);var boxTotals=new List<string>();var volumeTotals=new List<string>();
            foreach(var day in l.Days) {
                int start=day.Key;double serial=day.Value;string d="$"+MergeEngine.ExcelColumn(start)+"$8";g.Set(8,start,serial,StyleAt(old,8,start,dateHeadStyle));
                for(int slot=0;slot<5;slot++){g.Set(9,start+slot,"Короб, вид",StyleAt(old,9,start+slot,head));g.Set(10,start+slot,"Томов за сутки",StyleAt(old,10,start+slot,head));g.Set(11,start+slot,slot+1,StyleAt(old,11,start+slot,intStyle));}
                g.Set(9,start+5,"Завершено / томов",StyleAt(old,9,start+5,head));g.Set(9,start+6,"Примечание",StyleAt(old,9,start+6,head));
                foreach(var person in l.People) {
                    int row=person.Key;string who=person.Value;g.Set(row,1,who,label);var items=l.Volumes.Where(v=>v.Date==serial&&v.Who==who).ToList();var boxList=items.Where(v=>v.Box!="").Select(v=>v.Box).Distinct().ToList();string personRef="$A"+row;
                    for(int slot=0;slot<5;slot++) {
                        int col=start+slot;string cell=MergeEngine.ExcelColumn(col)+row;string key=boxList.Count>slot?boxList[slot]:"";string shown=key==""?"":key.Split('|')[1]+key.Split('|')[0].ToLowerInvariant();
                        g.Set(row,col,shown,boxStyle,"IFERROR(INDEX("+display+",_xlfn.AGGREGATE(15,6,(ROW("+dates+")-6)/(("+dates+"="+d+")*("+owners+"="+personRef+")*("+flags+"=1)),"+(slot+1)+")),\"\")");
                        g.Set(row+1,col,key==""?(object)"":items.Count(v=>v.Box==key),countStyle,"IF("+cell+"=\"\",\"\",COUNTIFS("+display+","+cell+","+dates+","+d+","+owners+","+personRef+"))");
                    }
                    g.Set(row,start+5,completion.Count(v=>v.Value==serial&&completedBy[v.Key]==who),num,"COUNTIFS("+completed+","+d+","+finishedWho+","+personRef+")");g.Set(row+1,start+5,items.Count,num,"COUNTIFS("+dates+","+d+","+owners+","+personRef+")");
                }
                int totalCol=start+5;g.Set(42,totalCol,completion.Count(v=>v.Value==serial),StyleAt(old,42,totalCol,num),"COUNTIF("+completed+","+d+")");g.Set(43,totalCol,l.Volumes.Count(v=>v.Date==serial),StyleAt(old,43,totalCol,num),"COUNTIF("+dates+","+d+")");boxTotals.Add(MergeEngine.ExcelColumn(totalCol)+"42");volumeTotals.Add(MergeEngine.ExcelColumn(totalCol)+"43");
            }
            g.Set(3,2,null,intStyle,"IF(ISNUMBER(B1),B1-SUM("+string.Join(",",boxTotals)+"),\"\")");g.Set(4,2,null,intStyle,"IF(ISNUMBER(B2),B2-SUM("+string.Join(",",volumeTotals)+"),\"\")");
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
