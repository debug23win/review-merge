"""Main window layout: the settings rows with the link to forget the answers, the toolbar buttons and a short window that scrolls."""
import os, shutil, subprocess, uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/.runs' / uuid.uuid4().hex
OUT.mkdir(parents=True)
# The harness opens the window, shows the link as after a saved answer (no settings are written) and checks the controls.
HARNESS = r'''
using System; using System.Collections.Generic; using System.Drawing; using System.Reflection; using System.Windows.Forms; using ReviewMerge;
class UiTest {
    static MainForm f;
    static T Get<T>(string name){return (T)typeof(MainForm).GetField(name,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(f);}
    static void Pump(){for(int i=0;i<20;i++){Application.DoEvents();System.Threading.Thread.Sleep(25);}}
    // Fully inside the client area of every parent up to the page, which scrolls.
    static bool Inside(Control c){var r=new Rectangle(Point.Empty,c.Size);for(Control k=c;k.Parent!=null&&!(k is TabPage);k=k.Parent){r.Offset(k.Location);if(r.Left<0||r.Top<0||r.Right>k.Parent.ClientSize.Width||r.Bottom>k.Parent.ClientSize.Height)return false;}return true;}
    [STAThread] static int Main() {
        f=new MainForm();f.StartPosition=FormStartPosition.Manual;f.Location=new Point(0,0);f.Show();Pump();
        float scale=f.CurrentAutoScaleDimensions.Height/96f;var errors=new List<string>();
        var link=Get<LinkLabel>("forgetNames");link.Text="Забыть сохранённые ответы (12)";link.Visible=true;f.PerformLayout();Pump();
        if(!Get<RadioButton>("rdMode").Checked&&!Inside(Get<CheckBox>("matchSurnames")))errors.Add("match surnames checkbox hidden "+Get<CheckBox>("matchSurnames").Bounds);
        if(!Inside(link)||link.Width<link.GetPreferredSize(Size.Empty).Width-1)errors.Add("forget link "+link.Bounds);
        int n=0;foreach(var b in Get<List<Button>>("editingButtons")){n++;if(b.Visible&&!Inside(b))errors.Add("button "+n+" clipped "+b.Bounds);}
        // A window as low as the working area of 1920x1080 at 150 %: the list and the log keep about four rows and the page scrolls.
        f.MinimumSize=Size.Empty;f.Size=new Size(f.Width,(int)(693*scale));Pump();
        var page=Get<TabControl>("tabs").TabPages[0];var files=Get<ListView>("files");var log=Get<RichTextBox>("log");
        if(files.Height<95*scale||log.Height<95*scale||!page.VerticalScroll.Visible)errors.Add("short window: list "+files.Height+", log "+log.Height+", scroll "+page.VerticalScroll.Visible);
        // A path copied in Explorer with «Копировать как путь» comes in quotes.
        if(MainForm.CleanPath(" \"C:\\Свод\\ПД.xlsx\" ")!="C:\\Свод\\ПД.xlsx"||MainForm.CleanPath(null)!="")errors.Add("quoted path not cleaned");
        foreach(string e in errors)Console.WriteLine(e);
        f.Close();return errors.Count;
    }
}
'''
shutil.copy(ROOT / 'Свод_проверки.exe', OUT / 'Свод_проверки.exe')
(OUT / 'UiTest.cs').write_text(HARNESS, encoding='utf-8')
# csc of the .NET Framework on Windows, mcs of Mono elsewhere; the same switches with «/» or «-».
windows = os.name == 'nt'
compiler = str(Path(os.environ.get('WINDIR', '')) / 'Microsoft.NET/Framework64/v4.0.30319/csc.exe') if windows else 'mcs'
flag = '/' if windows else '-'
subprocess.run([compiler, flag + 'nologo', flag + 'codepage:65001', flag + 'target:exe', flag + 'out:' + str(OUT / 'UiTest.exe'), flag + 'r:' + str(OUT / 'Свод_проверки.exe'),
                flag + 'r:System.Windows.Forms.dll', flag + 'r:System.Drawing.dll', str(OUT / 'UiTest.cs')], check=True, timeout=120)
run = [str(OUT / 'UiTest.exe')] if windows else ['mono', str(OUT / 'UiTest.exe')]
result = subprocess.run(run, cwd=OUT, capture_output=True, text=True, timeout=120)
assert result.returncode == 0, result.stdout + result.stderr
print('UI tests passed: surname matching and the forget link stay visible, toolbar buttons are not clipped, a short window scrolls and keeps the list and the log, quoted paths')
