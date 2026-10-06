using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ReviewMerge {
    public static class RemarkCounter {
        public static string Normalize(string text) {return Regex.Replace((text??"").Replace("\r","").Replace('\t',' ')," +"," ").Trim(' ').Replace("\n ","\n");}
        public static int Count(string text) {
            if(string.IsNullOrWhiteSpace(text))return 0;
            text=text.Replace("\r","").Trim();
            int listed=Regex.Matches(text,@"(?m)^[ \t]*(?:\d{1,3}[.)]|[-*•●▪])[ \t]+\S").Count;
            return listed>0?listed:Regex.Split(text,@"\n[ \t]*\n+").Count(p=>!string.IsNullOrWhiteSpace(p));
        }
        // Excel counts numbered/bulleted lines, otherwise nonempty paragraphs. Text stays untouched.
        public static string NormalizeFormula(string cell) {
            return "SUBSTITUTE(TRIM(SUBSTITUTE(SUBSTITUTE("+cell+",CHAR(13),\"\"),CHAR(9),\" \")),CHAR(10)&\" \",CHAR(10))";
        }
        public static string Formula(string cell,string clean,string tokens) {
            string lines="CHAR(10)&"+clean;
            string count="SUMPRODUCT((LEN("+lines+")-LEN(SUBSTITUTE("+lines+","+tokens+",\"\")))/LEN("+tokens+"))";
            string paragraphs=clean;for(int n=0;n<6;n++)paragraphs="SUBSTITUTE("+paragraphs+",CHAR(10)&CHAR(10)&CHAR(10),CHAR(10)&CHAR(10))";
            string paragraphsCount="1+(LEN("+paragraphs+")-LEN(SUBSTITUTE("+paragraphs+",CHAR(10)&CHAR(10),\"\")))/2";
            return "IF(LEN(TRIM("+cell+"))=0,0,IF("+count+">0,"+count+","+paragraphsCount+"))";
        }
    }
}
