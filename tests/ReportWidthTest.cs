using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

// Run against the freshly built Windows application; no production sources are copied.
// Fixed vectors deliberately separate the stored OOXML width, pixels and Excel UI units.
class ReportWidthTest {
    static readonly XNamespace N="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    static readonly XNamespace P="http://schemas.openxmlformats.org/package/2006/relationships";
    static readonly XNamespace A="http://schemas.openxmlformats.org/drawingml/2006/main";
    const BindingFlags PrivateInstance=BindingFlags.Instance|BindingFlags.NonPublic;
    static Type writerType,resultType;
    static MethodInfo columnWidth,digitWidth,maximumDigitWidth;
    static int checks;

    static void Equal(double expected,double actual,string label) {
        if(Math.Abs(expected-actual)>0.00000001||double.IsNaN(actual))
            throw new Exception(label+": expected "+expected.ToString("R",CultureInfo.InvariantCulture)+", got "+actual.ToString("R",CultureInfo.InvariantCulture));
        checks++;
    }
    static void True(bool condition,string label) {if(!condition)throw new Exception(label);checks++;}
    static double Width(double characters,int mdw) {return (double)columnWidth.Invoke(null,new object[]{characters,mdw});}
    static int Digit(string name,double size,bool bold,bool italic,int charset) {
        return (int)digitWidth.Invoke(null,new object[]{name,size,bold,italic,charset});
    }
    // Independent inverse from the Microsoft SpreadsheetML Column documentation.
    static int Pixels(double raw,int mdw) {
        return (int)Math.Truncate(((256*raw+Math.Truncate(128.0/mdw))/256)*mdw);
    }
    static double DialogWidth(double raw,int mdw) {
        int padding=2*(int)Math.Ceiling(mdw/4.0)+1;
        return Math.Truncate((Pixels(raw,mdw)-padding)/(double)mdw*100+0.5)/100;
    }
    static void Invalid(double characters,int mdw) {
        try {Width(characters,mdw);}
        catch(TargetInvocationException ex) {if(ex.InnerException is ArgumentOutOfRangeException){checks++;return;}throw;}
        throw new Exception("Invalid column width was accepted");
    }
    static void TestConversions() {
        int[] mdw={7,8,9,10};
        double[] raw={9.5703125,9.5,9.6640625,9.59765625};
        int[] pixels={67,76,87,96};
        double[] dialog={8.86,8.88,8.89,8.90};
        for(int i=0;i<mdw.Length;i++) {
            double actual=Width(8.9,mdw[i]);
            Equal(raw[i],actual,"Serialized 8.9 characters, MDW "+mdw[i]);
            Equal(pixels[i],Pixels(actual,mdw[i]),"Pixel width, MDW "+mdw[i]);
            Equal(dialog[i],DialogWidth(actual,mdw[i]),"Excel dialog width, MDW "+mdw[i]);
            Equal(Math.Truncate(actual*256),actual*256,"OOXML 1/256-character quantization");
        }
        Equal(8.22,DialogWidth(9,9),"Old raw width 9 reproduces the 8.22 regression");
        Equal(0,Width(0,7),"Zero width");
        Equal(0.85546875,Width(0.5,7),"Sub-character width uses proportional padding");
        Equal(1.7109375,Width(1,7),"One-character boundary");
        Equal(9.33203125,Width(8.5,9),"Half-pixel rounds away from zero");
        Equal(84,Pixels(Width(8.5,9),9),"Half-pixel result");
        Invalid(-0.01,7);Invalid(255.01,7);Invalid(double.NaN,7);
        Invalid(double.PositiveInfinity,7);Invalid(double.NegativeInfinity,7);
        Invalid(8.9,0);Invalid(8.9,-1);
        Console.WriteLine("PASS: fixed conversion vectors, inverse Excel units, rounding and invalid inputs");
    }

    static XElement Font(string name,double size,string scheme) {
        var font=new XElement(N+"font",new XElement(N+"name",new XAttribute("val",name)),
            new XElement(N+"sz",new XAttribute("val",size)),new XElement(N+"charset",new XAttribute("val",1)));
        if(scheme!=null)font.Add(new XElement(N+"scheme",new XAttribute("val",scheme)));
        return font;
    }
    static XDocument Styles(XElement target,string styleName,bool builtin,bool misleadingNormal) {
        var names=new XElement(N+"cellStyles");
        if(misleadingNormal)names.Add(new XElement(N+"cellStyle",new XAttribute("name","Normal"),new XAttribute("xfId",0),new XAttribute("builtinId",1)));
        var normal=new XElement(N+"cellStyle",new XAttribute("name",styleName),new XAttribute("xfId",3));
        if(builtin)normal.Add(new XAttribute("builtinId",0));
        names.Add(normal);
        // Both IDs are nonzero and different: cellStyle xfId=3 -> cellStyleXfs[3] fontId=2.
        string face=(string)target.Element(N+"name").Attribute("val");
        return new XDocument(new XElement(N+"styleSheet",
            new XElement(N+"fonts",Font(face,24,null),Font(face,30,null),target),
            new XElement(N+"cellStyleXfs",new XElement(N+"xf",new XAttribute("fontId",0)),
                new XElement(N+"xf",new XAttribute("fontId",1)),new XElement(N+"xf",new XAttribute("fontId",0)),
                new XElement(N+"xf",new XAttribute("fontId",2))),names));
    }
    sealed class Fixture {
        public object Writer,Result;
        public IDictionary Entries;
        public Fixture(XDocument styles) {
            Writer=Activator.CreateInstance(writerType,new object[]{null});
            Result=Activator.CreateInstance(resultType);
            writerType.GetField("styles",PrivateInstance).SetValue(Writer,styles);
            writerType.GetField("rels",PrivateInstance).SetValue(Writer,new XDocument(new XElement(P+"Relationships")));
            Entries=(IDictionary)writerType.GetField("entries",PrivateInstance).GetValue(Writer);
        }
        public IList Problems {get{return (IList)resultType.GetField("Problems").GetValue(Result);}}
        public int Measure() {return (int)maximumDigitWidth.Invoke(Writer,new object[]{Result});}
        public void Theme(string target,string entry,string major,string minor,bool external) {
            var rel=new XElement(P+"Relationship",new XAttribute("Id","rIdTheme"),
                new XAttribute("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
                new XAttribute("Target",target));
            if(external)rel.Add(new XAttribute("TargetMode","External"));
            writerType.GetField("rels",PrivateInstance).SetValue(Writer,new XDocument(new XElement(P+"Relationships",rel)));
            var theme=new XDocument(new XElement(A+"theme",new XElement(A+"themeElements",new XElement(A+"fontScheme",
                new XElement(A+"majorFont",new XElement(A+"latin",new XAttribute("typeface",major))),
                new XElement(A+"minorFont",new XElement(A+"latin",new XAttribute("typeface",minor)))))));
            Entries[entry]=Encoding.UTF8.GetBytes(theme.ToString(SaveOptions.DisableFormatting));
        }
    }
    sealed class KnownFont {
        public string Name;public int Mdw;
        public KnownFont(string name,int mdw){Name=name;Mdw=mdw;}
    }
    static List<KnownFont> TestNativeFonts() {
        var available=new List<KnownFont>();
        var expected=new[]{new KnownFont("Courier New",10),new KnownFont("Times New Roman",8),new KnownFont("Arial",9)};
        using(var installed=new InstalledFontCollection()) {
            var names=new HashSet<string>(installed.Families.Select(f=>f.Name),StringComparer.OrdinalIgnoreCase);
            foreach(var font in expected) {
                if(!names.Contains(font.Name)){Console.WriteLine("SKIP: native metric for uninstalled font "+font.Name);continue;}
                Equal(font.Mdw,Digit(font.Name,12,false,false,1),font.Name+" 12 pt, 96-DPI pixel font");
                var fixture=new Fixture(Styles(Font(font.Name,12,null),"Normal",true,false));
                Equal(font.Mdw,fixture.Measure(),font.Name+" Normal-style native metric");
                True(fixture.Problems.Count==0,"Installed font was substituted; not native verification: "+font.Name);
                available.Add(font);
                Console.WriteLine("PASS: installed "+font.Name+" 12 pt has native maximum digit width "+font.Mdw+" px; no fallback");
            }
        }
        True(available.Count>0,"No known test font is installed; native font verification did not run");
        return available;
    }
    static void TestStyles(List<KnownFont> fonts) {
        var font=fonts[0];
        True(Digit(font.Name,24,false,false,1)!=font.Mdw,"Style fixture needs distinct 12-pt and 24-pt metrics");
        var builtin=new Fixture(Styles(Font(font.Name,12,null),"Localized default",true,false));
        Equal(font.Mdw,builtin.Measure(),"builtinId=0 resolves nonzero xfId and fontId");
        var priority=new Fixture(Styles(Font(font.Name,12,null),"Localized default",true,true));
        Equal(font.Mdw,priority.Measure(),"builtinId=0 takes priority over a misleading Normal name");
        var named=new Fixture(Styles(Font(font.Name,12,null),"Normal",false,false));
        Equal(font.Mdw,named.Measure(),"Normal name fallback without builtinId");
        var localized=new Fixture(Styles(Font(font.Name,12,null),"Обычный",false,false));
        Equal(font.Mdw,localized.Measure(),"Localized Normal name fallback");
        var toggles=Font(font.Name,12,null);toggles.Add(new XElement(N+"b",new XAttribute("val",0)),new XElement(N+"i",new XAttribute("val","false")));
        Equal(font.Mdw,new Fixture(Styles(toggles,"Normal",true,false)).Measure(),"Explicit false font flags");
        var emphatic=Font(font.Name,12,null);emphatic.Add(new XElement(N+"b"),new XElement(N+"i"));
        Equal(Digit(font.Name,12,true,true,1),new Fixture(Styles(emphatic,"Normal",true,false)).Measure(),"Bold and italic Normal font flags");
        Console.WriteLine("PASS: Normal style identity, xf/font indirection, fallback names and font flags");

        const string missing="ReviewMergeMissingFont123456";
        foreach(string scheme in new[]{"minor","major"}) {
            var themed=new Fixture(Styles(Font(missing,12,scheme),"Normal",true,false));
            // A wrong scheme selection must produce a missing-font warning even if MDWs happen to match.
            themed.Theme("theme/width%20theme.xml","xl/theme/width theme.xml",scheme=="major"?font.Name:missing,scheme=="minor"?font.Name:missing,false);
            Equal(font.Mdw,themed.Measure(),scheme+" theme Latin typeface");
            True(themed.Problems.Count==0,scheme+" theme did not resolve an installed typeface");
        }
        var absolute=new Fixture(Styles(Font(missing,12,"minor"),"Normal",true,false));
        absolute.Theme("/xl/theme/absolute.xml","xl/theme/absolute.xml",missing,font.Name,false);
        Equal(font.Mdw,absolute.Measure(),"Absolute theme relationship target");
        True(absolute.Problems.Count==0,"Absolute theme target was not resolved");
        var external=new Fixture(Styles(Font(font.Name,12,"minor"),"Normal",true,false));
        external.Theme("theme/external.xml","xl/theme/external.xml",missing,missing,true);
        Equal(font.Mdw,external.Measure(),"External theme relationship is ignored");
        True(external.Problems.Count==0,"External theme altered Normal font selection");
        var absent=new Fixture(Styles(Font(missing,12,null),"Normal",true,false));
        True(absent.Measure()>0,"Missing-font fallback did not return a usable metric");
        True(absent.Problems.Count==1,"Font substitution must produce exactly one warning");
        object problem=absent.Problems[0];
        string detail=(string)problem.GetType().GetField("Detail").GetValue(problem);
        True(detail!=null&&detail.Contains(missing),"Substitution warning must identify the unavailable font");
        Equal(font.Mdw,Digit(font.Name,12,false,false,1),"Native font remains usable after substitution");
        Console.WriteLine("PASS: major/minor themes, URI targets, external-theme exclusion and explicit substitution warning");
    }
    static int Main(string[] args) {
        try {
            if(args.Length!=1||!File.Exists(args[0]))throw new ArgumentException("Pass the freshly built application executable path");
            var app=Assembly.LoadFrom(Path.GetFullPath(args[0]));
            writerType=app.GetType("ReviewMerge.XlsxWriter",true);resultType=app.GetType("ReviewMerge.MergeResult",true);
            columnWidth=writerType.GetMethod("ReportColumnWidth",BindingFlags.Static|BindingFlags.NonPublic);
            digitWidth=writerType.GetMethod("ReportDigitWidth",BindingFlags.Static|BindingFlags.NonPublic);
            maximumDigitWidth=writerType.GetMethod("ReportMaximumDigitWidth",PrivateInstance);
            True(columnWidth!=null&&digitWidth!=null&&maximumDigitWidth!=null,"Report-width methods were not found in the built application");
            TestConversions();TestStyles(TestNativeFonts());
            Console.WriteLine("Report width tests passed ("+checks+" assertions)");return 0;
        } catch(Exception ex) {Console.Error.WriteLine(ex);return 1;}
    }
}
