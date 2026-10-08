using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ReviewMerge {
    public sealed partial class XlsxWriter {
        const string CalculationMarker="ReviewMerge.FirstDates.v3";
        const int HelperWidth=48,MinSlots=7,SpareDays=2,SparePeople=2,SpareBoxes=3;
        sealed class VolumeRecord {public string List="",Who,Raw="",Kind="";public List<string> Boxes=new List<string>();public double? Date;public int Issue;public bool IsVolume=true;}
        // One day of the reference sheet: box slots, the "boxes / volumes" total and the note column. A spare day gets a date when a new date appears in the main sheet.
        sealed class DayBlock {public double? Date;public int Start,Slots;public int Total{get{return Start+Slots;}}public int Note{get{return Start+Slots+1;}}}
        sealed class Person {public string Name;public int Row;public bool FromSheet,Used;}
        sealed class ReferenceLayout {
            public XDocument Previous;
            public int Helper,End,OldHelper,OldEnd,OldPeopleEnd=43,ManifestStart,BoxEnd,ManifestEnd,BoxCount,OldBlocksEnd=1,BlocksEnd=1,Shift;
            public readonly List<DayBlock> Blocks=new List<DayBlock>(),OldBlocks=new List<DayBlock>();
            public readonly List<List<int>> VolumeGroups=new List<List<int>>();
            public readonly SortedDictionary<int,string> People=new SortedDictionary<int,string>();
            // Rows for reviewers who appear in the main sheet after the build: their names are formulas.
            public readonly List<int> SpareRows=new List<int>();
            // The row of a short spelling ("АА Яковлева") merged into the row of the person: its notes move there.
            public readonly Dictionary<int,int> MergedRows=new Dictionary<int,int>();
            public readonly List<string> Boxes=new List<string>();
            // Reviewer as written in the files -> person of the reference sheet; person -> row label of the surname statistics.
            public readonly Dictionary<string,string> Names=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            public readonly Dictionary<string,string> StatKeys=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            // The matching table of the sheet: every spelling of the files and every name of column A, so a name typed later in the main sheet finds its row.
            public readonly List<KeyValuePair<string,string>> NameTable=new List<KeyValuePair<string,string>>();
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
        string NamesRange(int h){return "$"+MergeEngine.ExcelColumn(h+45)+"$7:$"+MergeEngine.ExcelColumn(h+47)+"$"+Math.Max(7,referenceLayout.NameTable.Count+6);}
        string StatKeyOf(string who){string key;return string.IsNullOrEmpty(who)?"":referenceLayout.StatKeys.TryGetValue(who,out key)?key:who;}
        string StatKeyOfRaw(object raw){string name=Rules.CleanName(raw),who;if(name=="")return "";if(!referenceLayout.Names.TryGetValue(name,out who))who=name;return StatKeyOf(who);}
        // A reviewer of the main sheet becomes a person of the reference sheet. A full name already there is kept: "АА Яковлева" and "Яковлева А.А." go to
        // "Яковлева Александра Алексеевна" when the initials agree; other people with the same surname are merged when the operator confirms it.
        Person ResolvePerson(string raw,List<Person> persons,MergeResult result) {
            var exact=persons.FirstOrDefault(p=>Rules.NameKey(p.Name)==Rules.NameKey(raw));
            if(!options.MatchSurnames) {
                if(exact!=null)return exact;
                var bySurname=persons.Where(p=>string.Equals(Rules.Surname(p.Name),raw,StringComparison.OrdinalIgnoreCase)).ToList();
                return bySurname.Count==1?bySurname[0]:null;
            }
            string surname=Rules.SurnameKey(raw);if(surname=="")return exact;
            int fullness=exact==null?-1:Rules.Fullness(exact.Name);
            var others=persons.Where(p=>p!=exact&&Rules.SurnameKey(p.Name)==surname&&Rules.Fullness(p.Name)>fullness).OrderByDescending(p=>Rules.Fullness(p.Name)).ToList();
            var agreeing=others.Where(p=>Rules.InitialsAgree(raw,p.Name)).ToList();
            foreach(var p in others) {
                bool auto=agreeing.Count==1&&agreeing[0]==p&&(options.SavedSamePerson==null||options.SavedSamePerson(raw,p.Name)!=false);
                if(!auto&&options.SamePerson!=null&&!options.SamePerson(raw,p.Name))continue;
                // A short spelling of the sheet ("Галышева Е.И.") takes the full name from the files.
                string before=p.Name;if(Rules.Fullness(p.Name)==0&&Rules.Fullness(raw)>0)p.Name=raw;
                result.PeopleMerges.Add("«"+raw+"» → «"+p.Name+"»"+(before!=p.Name?" (было «"+before+"»)":"")+(auto?", инициалы совпадают":""));
                return p;
            }
            return exact;
        }
        // Days of the previous sheet: box slots, the total ("ИТОГО" in a manual sheet or "Коробов / томов") and the note column. Spare days without a date follow the dated ones.
        void ReadOldBlocks(ReferenceLayout layout) {
            var old=layout.Previous;Func<int,IEnumerable<XElement>> cells=r=>old.Root.Element(N+"sheetData").Elements(N+"row").Where(e=>(int)e.Attribute("r")==r).Elements(N+"c");
            Func<int,bool> inside=c=>c>=2&&(layout.OldHelper==0||c<layout.OldHelper);
            var days=new SortedDictionary<int,double>();
            foreach(var c in cells(8)){int col=Column((string)c.Attribute("r"));double? day=XlsxReader.DateSerial(Value(c));if(day.HasValue&&inside(col))days[col]=day.Value;}
            var heads=new Dictionary<int,string>();foreach(var c in cells(9)){int col=Column((string)c.Attribute("r"));if(inside(col))heads[col]=XlsxReader.Normal(Value(c));}
            Func<int,bool> isTotal=c=>{string t;return heads.TryGetValue(c,out t)&&(t.Contains("ИТОГО")||t.Contains("КОРОБОВ / ТОМОВ"));};
            var starts=days.Keys.ToList();
            for(int k=0;k<starts.Count;k++){
                int start=starts[k],next=k+1<starts.Count?starts[k+1]:int.MaxValue;
                int total=heads.Keys.Where(c=>c>start&&c<next-1&&isTotal(c)).DefaultIfEmpty(start+5).Min();
                layout.OldBlocks.Add(new DayBlock{Date=days[start],Start=start,Slots=Math.Max(1,total-start)});
            }
            int col2=layout.OldBlocks.Select(x=>x.Note+1).DefaultIfEmpty(2).Max();
            while(true){
                string head;if(!heads.TryGetValue(col2,out head)||!head.Contains("КОРОБ")||isTotal(col2))break;
                int start=col2,total=heads.Keys.Where(c=>c>start&&isTotal(c)).DefaultIfEmpty(0).Min();if(total==0||total-start>200)break;
                layout.OldBlocks.Add(new DayBlock{Start=start,Slots=total-start});col2=total+2;
            }
            layout.OldBlocksEnd=layout.OldBlocks.Select(x=>x.Note).DefaultIfEmpty(1).Max();
        }
        ReferenceLayout PrepareReference(MergeResult result,List<double?> dates,List<int> issues,DateTime asOf,int minHelper) {
            string path=PathFor("Справка по томам");var layout=new ReferenceLayout{Previous=path==null?null:Xml(path)};var old=layout.Previous;layout.OldHelper=HelperStart(old);
            layout.OldEnd=layout.OldHelper>0?Convert.ToInt32(Value(At(old,1,layout.OldHelper+4))??43,Inv):43;
            if(old!=null) {
                ReadOldBlocks(layout);
                // Names typed in column A are kept; a formula there is a spare row of the previous build.
                Func<int,string> name=r=>{var c=At(old,r,1);return c==null||c.Element(N+"f")!=null?"":Rules.CleanName(Value(c));};
                for(int r=12;r<=40;r+=2){string who=name(r);if(who!="")layout.People[r]=who;}
                if(layout.OldHelper>0&&Text(Value(At(old,1,layout.OldHelper)))==CalculationMarker) {
                    int start=Convert.ToInt32(Value(At(old,1,layout.OldHelper+3))??48,Inv);layout.OldPeopleEnd=Math.Max(43,start-3);
                    for(int r=44;r<start-2;r+=2){string who=name(r);if(who!="")layout.People[r]=who;}
                    for(int r=start;r<=layout.OldEnd;r++){string key=Text(Value(At(old,r,1)))+"|"+Text(Value(At(old,r,2)));if(!key.StartsWith("|",StringComparison.Ordinal)&&!key.EndsWith("|",StringComparison.Ordinal))layout.Expected[key]=Value(At(old,r,5));}
                }
            }
            this.referenceLayout=layout;
            // Reviewers come from the main sheet; fuller spellings go first, so a new "Петров П.П." joins "Петров Пётр Петрович" of the same files.
            var persons=layout.People.Select(p=>new Person{Name=p.Value,Row=p.Key,FromSheet=true}).ToList();var map=new Dictionary<string,Person>(StringComparer.OrdinalIgnoreCase);
            foreach(string raw in result.Rows.Select(r=>Rules.CleanName(r.Values[8])).Where(v=>v!="").Distinct(StringComparer.OrdinalIgnoreCase).OrderByDescending(Rules.Fullness).ToList()) {
                var p=ResolvePerson(raw,persons,result);if(p==null){p=new Person{Name=raw};persons.Add(p);}p.Used=true;map[raw]=p;
            }
            // A row of a short spelling that no reviewer uses any more ("АА Яковлева" next to "Яковлева Александра Алексеевна") is merged into its person.
            var mergedInto=new Dictionary<int,Person>();
            if(options.MatchSurnames)foreach(var p in persons.Where(p=>p.FromSheet&&!p.Used&&Rules.Fullness(p.Name)==0&&Rules.SurnameKey(p.Name)!="").ToList()) {
                var same=persons.Where(o=>o.Used&&Rules.SurnameKey(o.Name)==Rules.SurnameKey(p.Name)&&(Rules.Initials(p.Name)==""||Rules.InitialsAgree(o.Name,p.Name))).ToList();
                if(same.Count!=1)continue;
                persons.Remove(p);mergedInto[p.Row]=same[0];result.PeopleMerges.Add("Строка справки «"+p.Name+"» объединена со строкой «"+same[0].Name+"»");
            }
            var taken=new HashSet<int>(persons.Where(p=>p.FromSheet).Select(p=>p.Row));
            Func<int> nextRow=()=>{int r=Enumerable.Range(0,15).Select(i=>12+i*2).FirstOrDefault(x=>!taken.Contains(x));if(r==0)r=Math.Max(44,taken.Max()+2);taken.Add(r);return r;};
            foreach(var p in persons.Where(p=>!p.FromSheet))p.Row=nextRow();
            for(int i=0;i<SparePeople;i++)layout.SpareRows.Add(nextRow());
            layout.People.Clear();foreach(var p in persons)layout.People[p.Row]=p.Name;
            foreach(var pair in mergedInto)layout.MergedRows[pair.Key]=pair.Value.Row;
            foreach(var pair in map)layout.Names[pair.Key]=pair.Value.Name;
            layout.NameTable.AddRange(layout.Names);layout.NameTable.AddRange(layout.People.Values.Where(v=>!layout.Names.ContainsKey(v)).Select(v=>new KeyValuePair<string,string>(v,v)));
            // Statistics show the surname; the full name only when two different people share it.
            var used=layout.Names.Values.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            foreach(string who in used){string surname=Rules.Surname(who);layout.StatKeys[who]=surname==""||used.Count(o=>Rules.SurnameKey(o)==Rules.SurnameKey(who))>1?who:surname;}
            var keys=new List<List<string>>();
            for(int i=0;i<result.Rows.Count;i++){string raw=Rules.CleanName(result.Rows[i].Values[8]),who;if(!layout.Names.TryGetValue(raw,out who))who=raw;var boxes=Rules.BoxKeys(result.Rows[i].Values[10],result.Rows[i].Values[11]);keys.Add(boxes);layout.BoxFiles.Add(new VolumeRecord{Boxes=boxes,List=Rules.BoxList(boxes),Who=who,Date=dates[i]});}
            // Volumes: date of the last part when every part is checked, reviewer of that part, the first filled kind and boxes.
            foreach(var g in Enumerable.Range(0,result.Rows.Count).GroupBy(i=>Rules.VolumeKey(result.Rows[i].Values[2],result.Rows[i].Values[3]),StringComparer.Ordinal)) {
                var indices=g.ToList();double? d=indices.All(i=>dates[i].HasValue)?(double?)indices.Max(i=>dates[i].Value):null;int chosen=d.HasValue?indices.First(i=>dates[i]==d):indices[0];
                string raw=d.HasValue?Rules.CleanName(result.Rows[chosen].Values[8]):"",who;if(!layout.Names.TryGetValue(raw,out who))who=raw;
                var boxes=indices.Select(i=>keys[i]).FirstOrDefault(k=>k.Count>0)??new List<string>();
                layout.VolumeGroups.Add(indices);
                layout.Volumes.Add(new VolumeRecord{Boxes=boxes,List=Rules.BoxList(boxes),Who=who,Raw=raw,Date=d,Issue=indices.Any(i=>issues[i]==1)?1:0,Kind=indices.Select(i=>Text(result.Rows[i].Values[10])).FirstOrDefault(v=>v!="")??"",
                    IsVolume=indices.Any(i=>!Rules.IsUl(result.Rows[i].Values[2],result.Rows[i].Values[3]))});
            }
            // Days are the dates of the main sheet in date order, each with the same number of box slots (at least seven); spare days take dates added after the build.
            var days=dates.Where(v=>v.HasValue).Select(v=>v.Value).Distinct().OrderBy(v=>v).Select(v=>(double?)v).ToList();
            int slots=Math.Max(MinSlots,layout.Volumes.Where(v=>v.IsVolume&&v.Date.HasValue&&v.List!="").GroupBy(v=>v.Date.Value.ToString(Inv)+"|"+v.Who).Select(x=>x.Select(v=>v.List).Distinct().Count()).DefaultIfEmpty(0).Max());
            int col=2;
            foreach(double? day in days.Concat(Enumerable.Repeat((double?)null,SpareDays))){layout.Blocks.Add(new DayBlock{Date=day,Start=col,Slots=slots});col+=slots+2;}
            layout.BlocksEnd=col-1;layout.Shift=layout.OldBlocks.Count>0?layout.BlocksEnd-layout.OldBlocksEnd:0;
            layout.Boxes.AddRange(layout.BoxFiles.SelectMany(v=>v.Boxes).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v=>v.Substring(0,v.IndexOf('|')),StringComparer.Ordinal).ThenBy(BoxNumber));
            layout.BoxCount=layout.Boxes.Count;layout.ManifestStart=Math.Max(48,layout.People.Keys.Concat(layout.SpareRows).Max()+6);layout.BoxEnd=layout.ManifestStart+layout.Boxes.Count-1;layout.ManifestEnd=layout.BoxEnd+SpareBoxes;
            int existingEnd=old==null?1:old.Root.Element(N+"sheetData").Elements(N+"row").Elements(N+"c").Select(e=>Column((string)e.Attribute("r"))).Where(c=>layout.OldHelper==0||c<layout.OldHelper).DefaultIfEmpty(1).Max();
            if(layout.OldBlocks.Count>0&&existingEnd>layout.OldBlocksEnd)existingEnd+=layout.Shift;
            layout.End=Math.Max(existingEnd,layout.BlocksEnd);
            layout.Helper=Math.Min(16384-HelperWidth,Math.Max(Math.Max(96,layout.End+3),minHelper));return layout;
        }
        DayBlock OldBlockAt(int col){return referenceLayout.OldBlocks.FirstOrDefault(x=>col>=x.Start&&col<=x.Note);}
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
        void ReferenceVolume(Grid grid,int h,int rn,int j,List<int> indices) {
            var v=referenceLayout.Volumes[j];string display=Rules.BoxListDisplay(v.List);
            grid.Set(rn,h+37,v.Raw,bodyStyle,ChooseVolumeField(h,rn,indices,9,true));grid.Set(rn,h+38,v.Kind,bodyStyle,FirstFilled(h,rn,indices,8));grid.Set(rn,h+39,v.List,bodyStyle,FirstFilled(h,rn,indices,0));
            string rawRef=MergeEngine.ExcelColumn(h+37)+rn,keyRef=MergeEngine.ExcelColumn(h+39)+rn,whoRef=MergeEngine.ExcelColumn(h+40)+rn,dRef=MergeEngine.ExcelColumn(h+13)+rn,volumeRef=MergeEngine.ExcelColumn(h+2)+rn,names=NamesRange(h);
            grid.Set(rn,h+40,v.Who,bodyStyle,"IF("+rawRef+"=\"\",\"\",IFERROR(VLOOKUP("+rawRef+","+names+",2,FALSE),"+rawRef+"))");
            // One box is shown as "45пд"; boxes of a list cell are shown as read at the build, e.g. "45пд, 46пд".
            if(v.Boxes.Count>1)grid.Set(rn,h+41,display,bodyStyle);
            else grid.Set(rn,h+41,display,bodyStyle,"IF("+keyRef+"=\"\",\"\",MID("+keyRef+",FIND(\"|\","+keyRef+")+1,LEN("+keyRef+")-FIND(\"|\","+keyRef+")-1)&LOWER(MID("+keyRef+",2,FIND(\"|\","+keyRef+")-2)))");
            bool dated=v.IsVolume&&v.Date.HasValue&&v.Who!="";int unique=dated&&v.List!=""&&!referenceLayout.Volumes.Take(j).Any(o=>o.IsVolume&&o.List==v.List&&o.Who==v.Who&&o.Date==v.Date)?1:0;
            // Keys "date|reviewer" let the day cells compare one column instead of three; a record of the day is shown once per box list.
            string day=dated?v.Date.Value.ToString(Inv)+"|"+v.Who:"",tripleRef=MergeEngine.ExcelColumn(h+6)+rn,dayKey=dRef+"&\"|\"&"+whoRef;
            grid.Set(rn,h+6,dated&&v.List!=""?v.List+"|"+day:"",bodyStyle,"IF(OR("+volumeRef+"=0,"+dRef+"=\"\","+keyRef+"=\"\","+whoRef+"=\"\"),\"\","+keyRef+"&\"|\"&"+dayKey+")");
            grid.Set(rn,h+42,unique,intStyle,"IF("+tripleRef+"=\"\",0,IF(MATCH("+tripleRef+","+LocalRange(h+6,7,rn)+",0)="+(j+1)+",1,0))");
            grid.Set(rn,h+26,unique==1?day:"",bodyStyle,"IF("+MergeEngine.ExcelColumn(h+42)+rn+"=1,"+dayKey+",\"\")");
            grid.Set(rn,h+27,day,bodyStyle,"IF(OR("+volumeRef+"=0,"+dRef+"=\"\","+whoRef+"=\"\"),\"\","+dayKey+")");
            grid.Set(rn,h+43,StatKeyOf(v.Who),bodyStyle,"IF("+rawRef+"=\"\",\"\",IFERROR(VLOOKUP("+rawRef+","+names+",3,FALSE),"+rawRef+"))");
        }
        static int BoxNumber(string key){int n;int.TryParse(key.Substring(key.IndexOf('|')+1),NumberStyles.Integer,CultureInfo.InvariantCulture,out n);return n;}
        void WriteReference(Grid calculations,int h,int summaryReportEnd) {
            var l=referenceLayout;var g=new Grid();var old=l.Previous;var notes=new List<Tuple<int,DayBlock,XElement>>();
            if(old!=null)foreach(var row in old.Root.Element(N+"sheetData").Elements(N+"row")){
                int rn=(int)row.Attribute("r");g.Templates[rn]=new XElement(row);
                foreach(var c in row.Elements(N+"c")){
                    int col=Column((string)c.Attribute("r"));if(l.OldHelper>0&&col>=l.OldHelper&&col<l.OldHelper+HelperWidth)continue;
                    // Notes stay with their day and person; other cells of the days are rebuilt.
                    var block=rn>=8&&rn<=l.OldPeopleEnd?OldBlockAt(col):null;
                    if(block!=null){if(col==block.Note&&rn>=10)notes.Add(Tuple.Create(rn,block,c));continue;}
                    if(l.OldHelper>0&&rn>=44&&rn<=l.OldEnd&&col<=Math.Max(l.End,l.OldBlocksEnd))continue;
                    if(rn>=8&&rn<=43) {
                        if(col==1&&rn>=12&&rn<=41&&(c.Element(N+"f")!=null||l.MergedRows.ContainsKey(rn)))continue;
                        if(l.OldBlocks.Count>0&&col>l.OldBlocksEnd)col+=l.Shift;
                    }
                    var cell=new XElement(c);cell.SetAttributeValue("r",MergeEngine.ExcelColumn(col)+rn);
                    if(!g.Rows.ContainsKey(rn))g.Rows[rn]=new SortedDictionary<int,XElement>();g.Rows[rn][col]=cell;
                }
            }
            foreach(var row in calculations.Rows)foreach(var cell in row.Value.Where(c=>c.Key>=h&&c.Key<h+HelperWidth)){if(!g.Rows.ContainsKey(row.Key))g.Rows[row.Key]=new SortedDictionary<int,XElement>();g.Rows[row.Key][cell.Key]=new XElement(cell.Value);}
            g.Set(6,h+45,"Проверяющий в файлах",bodyStyle);g.Set(6,h+46,"Проверяющий в справке",bodyStyle);g.Set(6,h+47,"Фамилия в статистике",bodyStyle);g.Set(6,h+29,"Новый короб в учёте",bodyStyle);g.Set(6,h+44,"Ключ короба в учёте",bodyStyle);g.Set(6,h+26,"Запись дня для ячеек коробов",bodyStyle);g.Set(6,h+27,"Дата и проверяющий тома",bodyStyle);
            int n=7;foreach(var map in l.NameTable){g.Set(n,h+45,map.Key,bodyStyle);g.Set(n,h+46,map.Value,bodyStyle);g.Set(n++,h+47,StatKeyOf(map.Value),bodyStyle);}
            int last=l.Volumes.Count+6;string dates=LocalRange(h+13,7,last),keys=LocalRange(h+39,7,last),display=LocalRange(h+41,7,last),slotKeys=LocalRange(h+26,7,last),dayKeys=LocalRange(h+27,7,last),isv=","+LocalRange(h+2,7,last)+",1";
            // The first day of the previous sheet is the style template of every day.
            var first=l.OldBlocks.FirstOrDefault()??new DayBlock{Start=2,Slots=MinSlots};
            int label=StyleAt(old,12,1,bodyStyle),num=StyleAt(old,13,first.Total,intStyle),boxStyle=StyleAt(old,12,first.Start,bodyStyle),countStyle=StyleAt(old,13,first.Start,intStyle),head=StyleAt(old,9,first.Start,headStyle);
            if(old==null){g.Set(1,1,"Коробов получено",bodyStyle);g.Set(2,1,"Томов по проекту акта",bodyStyle);g.Set(8,1,"ФИО проверяющего",headStyle);}
            g.Set(3,1,"ИТОГО коробов проверить",bodyStyle);g.Set(4,1,"ИТОГО томов проверить",bodyStyle);g.Set(42,1,"ИТОГО проверено коробов",StyleAt(old,42,1,headStyle));g.Set(43,1,"ИТОГО проверено томов",StyleAt(old,43,1,headStyle));
            int title=l.ManifestStart-2,header=l.ManifestStart-1;g.Set(title,1,"УЧЁТ ПРОВЕРЕННЫХ КОРОБОВ",headStyle);
            string[] headers={"Вид","Короб","Томов в реестре","Проверено томов","Всего томов по описи","Дата проверки короба","Состояние","Проверил короб"};for(int c=0;c<headers.Length;c++)g.Set(header,c+1,headers[c],headStyle);g.Heights[header]=75;
            var completion=l.Completed;var completedBy=l.CompletedBy;int fileLast=l.BoxFiles.Count+6;string fileDates=LocalRange(h+4,7,fileLast),fileKeys=LocalRange(h,7,fileLast),fileOwners=LocalRange(h+9,7,fileLast);
            for(int i=0;i<=l.ManifestEnd-l.ManifestStart;i++) {
                int row=l.ManifestStart+i;bool spare=i>=l.Boxes.Count;string key=spare?"":l.Boxes[i],a="$A"+row;
                // A file belongs to every box of its list: ";ПД|45;ПД|46;" contains ";ПД|45;".
                string k=a+"&\"|\"&TEXT($B"+row+",\"0\")",inList="ISNUMBER(SEARCH(\";\"&"+k+"&\";\","+fileKeys+"))",volumeCriteria="\"*;\"&"+k+"&\";*\"";
                Func<string,string> live=f=>spare?"IF("+a+"=\"\",\"\","+f+")":f;
                object volumes="",checkedVolumes="",date="",state="",owner="";
                if(spare) {
                    // A box typed in the main sheet after the build: the next number that the table does not have yet.
                    string found=MergeEngine.ExcelColumn(h+29)+row;
                    g.Set(row,h+29,"",bodyStyle,"IFERROR(INDEX("+fileKeys+",_xlfn.AGGREGATE(15,6,(ROW("+fileKeys+")-6)/("+LocalRange(h+25,7,fileLast)+"=1),"+(i-l.Boxes.Count+1)+")),\"\")");
                    g.Set(row,1,"",bodyStyle,"IF("+found+"=\"\",\"\",MID("+found+",2,FIND(\"|\","+found+")-2))");
                    g.Set(row,2,"",intStyle,"IF("+found+"=\"\",\"\",VALUE(MID("+found+",FIND(\"|\","+found+")+1,LEN("+found+")-FIND(\"|\","+found+")-1)))");
                    g.Set(row,5,null,StyleAt(old,row,5,intStyle));
                } else {
                    var members=l.Volumes.Where(v=>v.IsVolume&&v.Boxes.Contains(key)).ToList();object expected;l.Expected.TryGetValue(key,out expected);
                    var checkedFiles=l.BoxFiles.Where(v=>v.Boxes.Contains(key)&&v.Date.HasValue&&v.Who!="").ToList();double? earliest=checkedFiles.Count>0?(double?)checkedFiles.Min(v=>v.Date.Value):null;completion[key]=earliest;completedBy[key]=earliest.HasValue?checkedFiles.First(v=>v.Date==earliest).Who:"";
                    g.Set(row,1,key.Split('|')[0],bodyStyle);g.Set(row,2,(double)BoxNumber(key),intStyle);g.Set(row,5,expected,StyleAt(old,row,5,intStyle));g.Set(row,h+44,";"+key+";",bodyStyle,"\";\"&"+k+"&\";\"");
                    volumes=members.Count;checkedVolumes=members.Count(v=>v.Date.HasValue);date=earliest.HasValue?(object)earliest.Value:"";state=earliest.HasValue?"Проверен":"Нет даты и фамилии";owner=completedBy[key];
                }
                string rawOwner="INDEX("+fileOwners+",MATCH(1,INDEX("+inList+"*("+fileDates+"=F"+row+"),0),0))";
                g.Set(row,3,volumes,intStyle,live("COUNTIFS("+keys+","+volumeCriteria+isv+")"));g.Set(row,4,checkedVolumes,intStyle,live("COUNTIFS("+keys+","+volumeCriteria+","+dates+",\">0\""+isv+")"));
                g.Set(row,6,date,dateStyle,live("IFERROR(_xlfn.AGGREGATE(15,6,"+fileDates+"/"+inList+"/("+fileDates+">0),1),\"\")"));g.Set(row,7,state,noteStyle,live("IF(F"+row+"<>\"\",\"Проверен\",\"Нет даты и фамилии\")"));
                g.Set(row,8,owner,bodyStyle,live("IF(F"+row+"=\"\",\"\",IFERROR(VLOOKUP("+rawOwner+","+NamesRange(h)+",2,FALSE),IFERROR("+rawOwner+",\"\")))"));g.Heights[row]=36;
            }
            string completed=LocalRange(6,l.ManifestStart,l.ManifestEnd),finishedWho=LocalRange(8,l.ManifestStart,l.ManifestEnd);
            // Reviewers: names of column A stay; spare rows show reviewers that appear in the main sheet after the build.
            foreach(var person in l.People)g.Set(person.Key,1,person.Value,label);
            for(int i=0;i<l.SpareRows.Count;i++)g.Set(l.SpareRows[i],1,"",label,"IFERROR(INDEX("+fileOwners+",_xlfn.AGGREGATE(15,6,(ROW("+fileOwners+")-6)/("+LocalRange(h+7,7,fileLast)+"=1),"+(i+1)+")),\"\")");
            var rows=l.People.Keys.Concat(l.SpareRows).OrderBy(r=>r).ToList();
            for(int b=0;b<l.Blocks.Count;b++) {
                // A day shows the next date of the main sheet, so a changed date moves its boxes to the right day.
                var block=l.Blocks[b];int start=block.Start;string d="$"+MergeEngine.ExcelColumn(start)+"$8",previous=b==0?null:"$"+MergeEngine.ExcelColumn(l.Blocks[b-1].Start)+"$8";
                string day=previous==null?"IFERROR(_xlfn.AGGREGATE(15,6,"+fileDates+"/("+fileDates+">0),1),\"\")":"IF("+previous+"=\"\",\"\",IFERROR(_xlfn.AGGREGATE(15,6,"+fileDates+"/("+fileDates+">"+previous+"),1),\"\"))";
                g.Set(8,start,block.Date.HasValue?(object)block.Date.Value:"",StyleAt(old,8,first.Start,dateHeadStyle),day);
                for(int slot=0;slot<block.Slots;slot++){int like=first.Start+Math.Min(slot,first.Slots-1);g.Set(9,start+slot,"Короб, вид",StyleAt(old,9,like,head));g.Set(10,start+slot,"Томов за сутки",StyleAt(old,10,like,head));g.Set(11,start+slot,slot+1,StyleAt(old,11,like,intStyle));}
                g.Set(9,block.Total,"Коробов / томов",StyleAt(old,9,first.Total,head));g.Set(9,block.Note,"Примечание",StyleAt(old,9,first.Note,head));
                foreach(int row in rows) {
                    string who;l.People.TryGetValue(row,out who);bool shown=who!=null&&block.Date.HasValue;string personRef="$A"+row,blank="IF(OR("+d+"=\"\","+personRef+"=\"\"),\"\",",dayKey=d+"&\"|\"&"+personRef;
                    var items=shown?l.Volumes.Where(v=>v.IsVolume&&v.Date==block.Date&&v.Who==who).ToList():new List<VolumeRecord>();var boxList=items.Where(v=>v.List!="").Select(v=>v.List).Distinct().ToList();
                    for(int slot=0;slot<block.Slots;slot++) {
                        int col=start+slot;string cell=MergeEngine.ExcelColumn(col)+row;string list=boxList.Count>slot?boxList[slot]:"";
                        g.Set(row,col,Rules.BoxListDisplay(list),boxStyle,blank+"IFERROR(INDEX("+display+",_xlfn.AGGREGATE(15,6,(ROW("+slotKeys+")-6)/("+slotKeys+"="+dayKey+"),"+(slot+1)+")),\"\"))");
                        g.Set(row+1,col,list==""?(object)"":items.Count(v=>v.List==list),countStyle,"IF("+cell+"=\"\",\"\",COUNTIFS("+display+","+cell+","+dayKeys+","+dayKey+"))");
                    }
                    g.Set(row,block.Total,shown?(object)completion.Count(v=>v.Value==block.Date&&completedBy[v.Key]==who):"",num,blank+"COUNTIFS("+completed+","+d+","+finishedWho+","+personRef+"))");
                    g.Set(row+1,block.Total,shown?(object)items.Count:"",num,blank+"COUNTIF("+dayKeys+","+dayKey+"))");
                }
                g.Set(42,block.Total,block.Date.HasValue?(object)completion.Count(v=>v.Value==block.Date):"",StyleAt(old,42,first.Total,num),"IF("+d+"=\"\",\"\",COUNTIF("+completed+","+d+"))");
                g.Set(43,block.Total,block.Date.HasValue?(object)l.Volumes.Count(v=>v.IsVolume&&v.Date==block.Date):"",StyleAt(old,43,first.Total,num),"IF("+d+"=\"\",\"\",COUNTIFS("+dates+","+d+isv+"))");
            }
            PlaceNotes(g,notes);
            object receivedBoxes=g.Rows.ContainsKey(1)&&g.Rows[1].ContainsKey(2)?Value(g.Rows[1][2]):null,receivedVolumes=g.Rows.ContainsKey(2)&&g.Rows[2].ContainsKey(2)?Value(g.Rows[2][2]):null;
            g.Set(3,2,receivedBoxes is double?(object)((double)receivedBoxes-completion.Count(v=>v.Value.HasValue)):"",intStyle,"IF(ISNUMBER(B1),B1-COUNTIF("+completed+",\">0\"),\"\")");
            g.Set(4,2,receivedVolumes is double?(object)((double)receivedVolumes-l.Volumes.Count(v=>v.IsVolume&&v.Date.HasValue)):"",intStyle,"IF(ISNUMBER(B2),B2-COUNTIFS("+dates+",\">0\""+isv+"),\"\")");
            g.Set(1,h,CalculationMarker,bodyStyle);g.Set(1,h+2,summaryReportEnd,intStyle);g.Set(1,h+3,l.ManifestStart,intStyle);g.Set(1,h+4,l.ManifestEnd,intStyle);
            var doc=old==null?Sheet(g,new[]{Col(1,1,32),Col(2,l.End,14)},false,l.End):new XDocument(old);if(old!=null)doc.Root.Element(N+"sheetData").ReplaceWith(g.Data());
            doc.Root.Element(N+"dimension").SetAttributeValue("ref","A1:"+MergeEngine.ExcelColumn(h+HelperWidth-1)+g.Rows.Keys.Max());
            var cols=doc.Root.Element(N+"cols");if(cols==null){cols=new XElement(N+"cols");doc.Root.Element(N+"sheetData").AddBeforeSelf(cols);}
            var oldCols=cols.Elements().Where(c=>l.OldHelper==0||(int)c.Attribute("min")<l.OldHelper).ToList();
            Func<int,XElement> oldCol=c=>oldCols.FirstOrDefault(e=>(int)e.Attribute("min")<=c&&(int)e.Attribute("max")>=c);
            // Column widths follow the role of the column in its day (box slot, total or note) as in the first day of the previous sheet.
            var widths=new List<XElement>();
            for(int c=1;c<=l.End;c++) {
                var block=l.Blocks.FirstOrDefault(x=>c>=x.Start&&c<=x.Note);int like=c;
                if(block!=null)like=l.OldBlocks.Count==0?0:c==block.Note?first.Note:c==block.Total?first.Total:first.Start+Math.Min(c-block.Start,first.Slots-1);
                else if(l.OldBlocks.Count>0&&c>l.BlocksEnd)like=c-l.Shift;
                var source=like>0?oldCol(like):null;var e=source!=null?new XElement(source):Col(c,c,c==1?32:block!=null&&c==block.Note?30:14);
                e.SetAttributeValue("min",c);e.SetAttributeValue("max",c);widths.Add(e);
            }
            cols.RemoveNodes();cols.Add(widths);cols.Add(Col(h,h+HelperWidth-1,18,true));
            var merges=doc.Root.Element(N+"mergeCells");if(merges==null){merges=new XElement(N+"mergeCells");doc.Root.Element(N+"sheetData").AddAfterSelf(merges);}
            int dayArea=Math.Max(l.OldBlocksEnd,l.BlocksEnd);
            merges.Elements().Where(m=>{var r=Regex.Match((string)m.Attribute("ref")??"",@"^([A-Z]+)8:[A-Z]+8$");return r.Success&&Column(r.Groups[1].Value)>=2&&Column(r.Groups[1].Value)<=dayArea;}).Remove();
            foreach(var block in l.Blocks)merges.Add(new XElement(N+"mergeCell",new XAttribute("ref",MergeEngine.ExcelColumn(block.Start)+"8:"+MergeEngine.ExcelColumn(block.Note)+"8")));
            merges.SetAttributeValue("count",merges.Elements().Count());PutSheet("Справка по томам",doc);
        }
        // A note goes to the same date (a spare day to a spare day) and to the row of its person. When its date is gone, it joins the note of the previous day with its date in front.
        void PlaceNotes(Grid g,List<Tuple<int,DayBlock,XElement>> notes) {
            var l=referenceLayout;var oldSpare=l.OldBlocks.Where(b=>!b.Date.HasValue).ToList();var newSpare=l.Blocks.Where(b=>!b.Date.HasValue).ToList();var dated=l.Blocks.Where(b=>b.Date.HasValue).ToList();
            foreach(var note in notes) {
                int rn=note.Item1,target;var from=note.Item2;var source=note.Item3;
                if(l.MergedRows.TryGetValue(rn,out target))rn=target;else if(l.MergedRows.TryGetValue(rn-1,out target))rn=target+1;
                int spareIndex=oldSpare.IndexOf(from);var to=from.Date.HasValue?l.Blocks.FirstOrDefault(b=>b.Date==from.Date):spareIndex<newSpare.Count?newSpare[spareIndex]:null;string prefix="";
                if(to==null){to=dated.LastOrDefault(b=>from.Date.HasValue&&b.Date<from.Date)??dated.FirstOrDefault()??l.Blocks[0];if(from.Date.HasValue)prefix=DateTime.FromOADate(from.Date.Value).ToString("dd.MM.yyyy",Inv)+": ";}
                if(!g.Rows.ContainsKey(rn))g.Rows[rn]=new SortedDictionary<int,XElement>();
                XElement existing;g.Rows[rn].TryGetValue(to.Note,out existing);string text=Text(Value(source)),before=existing==null?"":Text(Value(existing));
                if(text==""&&existing!=null)continue;
                if(before==""&&prefix==""){var moved=new XElement(source);moved.SetAttributeValue("r",MergeEngine.ExcelColumn(to.Note)+rn);g.Rows[rn][to.Note]=moved;}
                else g.Set(rn,to.Note,before==""?prefix+text:before+"\n"+prefix+text,(int?)source.Attribute("s")??noteStyle);
            }
        }
        static string ReferenceFormula(string formula,int h) {
            return Regex.Replace(formula,@"(?<![A-Z0-9_!])\$?([A-Z]{1,3})\$?\d+(?::\$?([A-Z]{1,3})\$?\d+)?",m=>Column(m.Groups[1].Value)>=h&&Column(m.Groups[1].Value)<h+HelperWidth?"'Справка по томам'!"+m.Value:m.Value);
        }
    }
}
