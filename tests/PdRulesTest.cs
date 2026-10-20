using System;
using System.Collections.Generic;
using System.Reflection;
using ReviewMerge;

class PdRulesTest {
    static void Check(bool ok,string message) { if(!ok)throw new Exception(message); }
    static SourceRow Document(object crc,object time,string extension) {
        var row=new SourceRow();row.Values[2]="Документ.pdf";row.Values[3]=extension;
        row.Values[5]="01.10.2026";row.Values[6]=time;row.Values[7]=crc;return row;
    }
    static int Main() {
        foreach(object value in new object[]{0,0.0,"0.0","-0","+0","0,00","0e3",false,"\u00a0\t"," нет "})Check(!Rules.IsResolved(value),"False resolved mark: "+value);
        foreach(object value in new object[]{1,2,true,"да","1",46300.0})Check(Rules.IsResolved(value),"Missing resolved mark: "+value);
        Check(!Rules.IsRemarkMark("+")&&!Rules.IsRemarkMark(2)&&Rules.IsRemarkMark("1 "),"Invalid mark counted as a remark");
        Check(XlsxReader.Normal("№ тома")=="№ ТОМА","Numero sign changed");
        Check(Rules.VolumeKey("Документ Фрагмент №1","pdf")==Rules.VolumeKey("Документ Фрагмент №2","pdf"),"Numbered fragments not grouped");
        Check(!Rules.IsUl("Путепровод на ул Мира","pdf")&&Rules.IsUl("Том-УЛ-ГТСС1","pdf"),"Street confused with IUL");
        Check(Rules.Surname("КИМ Алексей Петрович")=="Ким"&&Rules.Surname("ЦОЙ Виктор Робертович")=="Цой"&&Rules.InitialsAgree("КИМ Алексей Петрович","Ким А.П.")&&Rules.Fullness("КИМ Алексей Петрович")==2,"Short surname in capitals read as initials");
        Check(Rules.Surname("Пётр Петрович Петров")=="Петров"&&Rules.Surname("Петров Пётр Петрович")=="Петров"&&Rules.Surname("АА Яковлева")=="Яковлева"&&Rules.Surname("Яковлева А.А.")=="Яковлева"&&Rules.Fullness("Яковлева Александра Алексеевна")==2&&Rules.Fullness("АА Яковлева")==0,"Surname of a full name");
        Check(Rules.SamePersonSpelling("Петров","Петров П.П.")&&Rules.SamePersonSpelling("Иванов И.И.","Иванов Иван Иванович")&&!Rules.SamePersonSpelling("Анна","Борис")&&!Rules.SamePersonSpelling("Иванов А.А.","Иванов Б.Б."),"Spellings of one person");
        Check(Rules.VolumeKey("Путепровод на ул Мира","pdf")!=Rules.VolumeKey("Путепровод на ул Ленина","pdf"),"Street names merged");
        var rows=new List<MergedRow>();foreach(string name in new[]{"Проект А","Проект Б"}){var row=new MergedRow{Values=new object[20],VolumeNumber="б/н 2"};row.Values[2]=name;row.Values[3]="pdf";rows.Add(row);}
        var keys=Rules.VolumeKeys(rows);Check(keys[0]!=keys[1],"Placeholder volume numbers merged unrelated projects");
        string key=XlsxReader.Key(Document("05688907","10:00","pdf"));
        Check(key==XlsxReader.Key(Document(5688907,46300+10.0/24,".pdf")),"CRC leading zero, extension or timestamp not normalized");
        foreach(string date in new[]{"06.10.2026 10:15:00","06.10.2026 9:15","6/10/2026","2.10.26","02.10.2026."})Check(XlsxReader.DateSerial(date).HasValue,"Date not read: "+date);
        var method=typeof(XlsxWriter).GetMethod("ShiftMainFormula",BindingFlags.Static|BindingFlags.NonPublic);
        string formula="HEX2DEC(M7)+LOG10(A1)+SUM(M:M)+SUM(M:R)+INDIRECT(\"M7\")+Other!M7";
        string actual=(string)method.Invoke(null,new object[]{formula,"Основной",true});
        Check(actual=="HEX2DEC(N7)+LOG10(A1)+SUM(N:N)+SUM(N:S)+INDIRECT(\"M7\")+Other!M7","Formula corrupted: "+actual);
        var addresses=typeof(XlsxWriter).GetMethod("ShiftColumnAddress",BindingFlags.Static|BindingFlags.NonPublic);
        Check((string)addresses.Invoke(null,new object[]{"M7:M20 S7:S20"})=="N7:N20 T7:T20","Conditional formatting ranges not shifted");
        Console.WriteLine("PD rules passed: resolved marks, dates, CRC32, numero sign, IUL, placeholders and safe formula migration");return 0;
    }
}
