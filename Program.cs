using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace ReviewMerge {
    static class Program {
        [STAThread]
        static int Main(string[] args) {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            try {
                if(args.Length>0&&args[0]=="--apply-update")return DesktopUpdates.Updates.Apply(args);
                if (args.Length > 0 && args[0] == "--batch") {
                    // Options of the batch mode answer every question with "yes": --match-surnames merges people with one surname, --fill-incomplete fills the remaining rows of partly checked boxes.
                    string[] switches={"--current-date-for-new","--match-surnames","--fill-incomplete"};
                    var options=new MergeOptions{CurrentDateForNewReviews=args.Contains(switches[0]),MatchSurnames=args.Contains(switches[1])};
                    if(args.Contains(switches[2]))options.FillIncomplete=list=>list;
                    args=args.Where(a=>!switches.Contains(a)).ToArray();
                    if (args.Length < 5) throw new ArgumentException("--batch результат.xlsx yyyy-MM-dd основной-лист файл1.xlsx [файл2.xlsx ...] [--current-date-for-new] [--match-surnames] [--fill-incomplete]");
                    DateTime date = DateTime.ParseExact(args[2], "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                    var logs = new List<string>();
                    var result = MergeEngine.Run(args.Skip(4).Select(Path.GetFullPath).ToList(), args[1], args[3], date, (n,s) => logs.Add(n+"%: "+s), CancellationToken.None, options);
                    var report = new { result.Files, result.Documents, result.Reviewed, result.WithIssues, result.Boxes, result.DatesAssignedToday, result.FilledRows, Conflicts=result.Conflicts.Count, Problems=result.Problems.Count,
                        IncompleteBoxes=result.IncompleteBoxes.Select(v=>v.ToString()).ToList(), FilledBoxes=result.FilledBoxes.Select(v=>v.ToString()).ToList(), result.PeopleMerges, result.Output, Log=logs };
                    File.WriteAllText(args[1]+".run.json", new JavaScriptSerializer().Serialize(report), new UTF8Encoding(false));
                    return 0;
                }
                if (args.Length > 1 && args[0] == "--snapshot") {
                    using (var form = new MainForm()) {
                        foreach (string file in args.Skip(2)) form.AddPath(file);
                        form.Opacity=0; form.Show(); Application.DoEvents(); form.PerformLayout();
                        using (var bitmap = new Bitmap(form.Width,form.Height)) { form.DrawToBitmap(bitmap,new Rectangle(0,0,form.Width,form.Height)); bitmap.Save(args[1]); }
                    }
                    return 0;
                }
                Application.Run(new MainForm()); return 0;
            } catch(Exception ex) {
                if(args.Length>1 && args[0]=="--batch") File.WriteAllText(args[1]+".error.txt",ex.ToString(),Encoding.UTF8);
                else MessageBox.Show(ex.Message,"Свод проверки",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return 1;
            }
        }
    }
    public sealed class MainForm : Form {
        readonly List<string> paths = new List<string>();
        readonly ListView files = new ListView();
        readonly TextBox output = new TextBox(), sheetName = new TextBox();
        readonly DateTimePicker asOf = new DateTimePicker();
        readonly CheckBox currentDate = new CheckBox(), matchSurnames = new CheckBox();
        readonly LinkLabel forgetNames = new LinkLabel();
        readonly RichTextBox log = new RichTextBox();
        readonly ProgressBar progress = new ProgressBar();
        readonly Label totals = new Label(), rule = new Label();
        readonly Button build = new Button(), cancel = new Button(), open = new Button();
        readonly DataGridView conflicts = new DataGridView();
        readonly List<Button> editingButtons = new List<Button>();
        readonly TabControl tabs = new TabControl();
        CancellationTokenSource cancellation;
        bool running;
        string lastOutput;
        public MainForm() {
            Text="Свод проверки документации"; Font=new Font("Segoe UI",10); BackColor=Color.FromArgb(246,248,250);
            ClientSize=new Size(1060,820);MinimumSize=new Size(980,815);StartPosition=FormStartPosition.CenterScreen;
            AutoScaleMode=AutoScaleMode.Dpi;AllowDrop=true;
            var layout=new TableLayoutPanel { Dock=DockStyle.Fill,Padding=new Padding(22),ColumnCount=1,RowCount=4 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,43));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,38));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));
            Controls.Add(layout);
            var titleBar=new FlowLayoutPanel{Dock=DockStyle.Fill,WrapContents=false};titleBar.Controls.Add(new Label {Text="Свод проверки документации",Font=new Font("Segoe UI",19,FontStyle.Bold),AutoSize=true,ForeColor=Color.FromArgb(28,46,65)});titleBar.Controls.Add(DesktopUpdates.Updates.Attach(this,()=>running));layout.Controls.Add(titleBar,0,0);
            layout.Controls.Add(new Label {Text="Объединение отметок и замечаний. Статистика рассчитывается формулами из основного листа.",AutoSize=true,ForeColor=Color.FromArgb(73,91,108)},0,1);
            tabs.Dock=DockStyle.Fill;layout.Controls.Add(tabs,0,2);
            var page=new TabPage("Сборка свода") {Padding=new Padding(14),BackColor=Color.White};tabs.TabPages.Add(page);
            var content=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=9};
            content.RowStyles.Add(new RowStyle(SizeType.Absolute,38));content.RowStyles.Add(new RowStyle(SizeType.Percent,48));
            content.RowStyles.Add(new RowStyle(SizeType.Absolute,40));content.RowStyles.Add(new RowStyle(SizeType.Absolute,64));content.RowStyles.Add(new RowStyle(SizeType.Absolute,66));
            content.RowStyles.Add(new RowStyle(SizeType.Absolute,44));content.RowStyles.Add(new RowStyle(SizeType.Absolute,28));
            content.RowStyles.Add(new RowStyle(SizeType.Percent,52));content.RowStyles.Add(new RowStyle(SizeType.Absolute,31));page.Controls.Add(content);
            var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,WrapContents=false};
            buttons.Controls.Add(EditButton("Добавить файлы",130,(s,e)=>ChooseFiles()));
            buttons.Controls.Add(EditButton("Добавить папку",135,(s,e)=>ChooseFolder()));
            buttons.Controls.Add(EditButton("Удалить",90,(s,e)=>RemoveSelected()));
            buttons.Controls.Add(EditButton("Выше",74,(s,e)=>MoveSelected(-1)));
            buttons.Controls.Add(EditButton("Ниже",74,(s,e)=>MoveSelected(1)));content.Controls.Add(buttons,0,0);
            files.Dock=DockStyle.Fill;files.View=View.Details;files.FullRowSelect=true;files.MultiSelect=true;files.HideSelection=false;
            files.Columns.Add("№",38);files.Columns.Add("Файл",265);files.Columns.Add("Папка",630);content.Controls.Add(files,0,1);
            var parameters=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=4};parameters.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,120));parameters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));parameters.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,125));parameters.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,150));
            parameters.Controls.Add(new Label {Text="Основной лист:",AutoSize=true,Anchor=AnchorStyles.Left},0,0);
            sheetName.Text="Все загруженные файлы";sheetName.Dock=DockStyle.Fill;parameters.Controls.Add(sheetName,1,0);
            parameters.Controls.Add(new Label {Text="Дата свода:",AutoSize=true,Anchor=AnchorStyles.Left},2,0);
            asOf.Format=DateTimePickerFormat.Custom;asOf.CustomFormat="dd.MM.yyyy";asOf.Dock=DockStyle.Fill;parameters.Controls.Add(asOf,3,0);content.Controls.Add(parameters,0,2);
            var settings=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=2,RowCount=2,Margin=new Padding(0)};settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            currentDate.Text="Ставить текущую дату новым проверенным томам";currentDate.AutoSize=true;currentDate.Anchor=AnchorStyles.Left;currentDate.Checked=UserSettings.GetFlag("CurrentDateForNewReviews",false);currentDate.CheckedChanged+=(s,e)=>{SaveFlag("CurrentDateForNewReviews",currentDate.Checked);UpdateRule();};settings.Controls.Add(currentDate,0,0);
            matchSurnames.Text="Сопоставлять проверяющих по фамилии (спрашивать при каждом совпадении)";matchSurnames.AutoSize=true;matchSurnames.Anchor=AnchorStyles.Left;matchSurnames.Checked=UserSettings.GetFlag("MatchSurnames",true);matchSurnames.CheckedChanged+=(s,e)=>SaveFlag("MatchSurnames",matchSurnames.Checked);settings.Controls.Add(matchSurnames,0,1);
            forgetNames.AutoSize=true;forgetNames.Anchor=AnchorStyles.Left;forgetNames.Margin=new Padding(18,3,0,0);forgetNames.LinkClicked+=(s,e)=>{if(MessageBox.Show(this,"Забыть сохранённые ответы о совпадающих фамилиях? При следующей сборке программа спросит снова.",Text,MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes)return;try{UserSettings.ForgetSamePerson();}catch(Exception ex){Append("Не удалось сбросить ответы: "+ex.Message);}RefreshForget();};settings.Controls.Add(forgetNames,1,1);RefreshForget();
            content.Controls.Add(settings,0,3);
            rule.Dock=DockStyle.Fill;rule.ForeColor=Color.FromArgb(73,91,108);UpdateRule();content.Controls.Add(rule,0,4);
            var save=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=3};save.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,120));save.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));save.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,115));
            save.Controls.Add(new Label {Text="Сводный документ:",AutoSize=true,Anchor=AnchorStyles.Left},0,0);output.Dock=DockStyle.Fill;save.Controls.Add(output,1,0);
            save.Controls.Add(EditButton("Обзор…",105,(s,e)=>ChooseOutput()),2,0);content.Controls.Add(save,0,5);
            progress.Dock=DockStyle.Fill;progress.Maximum=100;progress.Margin=new Padding(0,3,0,6);content.Controls.Add(progress,0,6);
            log.Dock=DockStyle.Fill;log.ReadOnly=true;log.BackColor=Color.FromArgb(248,250,252);log.BorderStyle=BorderStyle.FixedSingle;log.Font=new Font("Segoe UI",9);
            log.Text="Добавьте файлы проверяющих и выберите существующий сводный документ.\nЕго строки и оформление сохраняются; новые документы добавляются в конец. Перед записью создаётся резервная копия. Закройте документ в Excel на время сборки.";content.Controls.Add(log,0,7);
            totals.Dock=DockStyle.Fill;totals.Text="Файлы не выбраны";totals.ForeColor=Color.FromArgb(28,46,65);content.Controls.Add(totals,0,8);
            var conflictPage=new TabPage("Конфликты") {Padding=new Padding(12),BackColor=Color.White};tabs.TabPages.Add(conflictPage);
            conflicts.Dock=DockStyle.Fill;conflicts.ReadOnly=true;conflicts.AllowUserToAddRows=false;conflicts.AllowUserToDeleteRows=false;conflicts.AutoGenerateColumns=false;conflicts.BackgroundColor=Color.White;conflicts.RowHeadersVisible=false;conflicts.AutoSizeRowsMode=DataGridViewAutoSizeRowsMode.DisplayedCells;
            conflicts.DefaultCellStyle.WrapMode=DataGridViewTriState.True;conflicts.DefaultCellStyle.Font=new Font("Segoe UI",9);conflicts.ColumnHeadersHeightSizeMode=DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            foreach(var item in new [] {new [] {"Document","Документ"},new [] {"Field","Поле"},new [] {"Details","Различия в источниках"},new [] {"Decision","Применённое правило"}})
                conflicts.Columns.Add(new DataGridViewTextBoxColumn {DataPropertyName=item[0],HeaderText=item[1],AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill,FillWeight=item[0]=="Details"?160:100});
            conflictPage.Controls.Add(conflicts);
            var actions=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft,WrapContents=false,Padding=new Padding(0,6,0,0)};
            build.Text="Собрать свод";build.Width=160;build.Height=33;build.BackColor=Color.FromArgb(35,69,100);build.ForeColor=Color.White;build.FlatStyle=FlatStyle.Flat;build.Click+=(s,e)=>StartBuild();actions.Controls.Add(build);
            cancel.Text="Отмена";cancel.Width=105;cancel.Height=33;cancel.Enabled=false;cancel.Click+=(s,e)=>CancelBuild();actions.Controls.Add(cancel);
            open.Text="Открыть результат";open.Width=175;open.Height=33;open.Enabled=false;open.Click+=(s,e)=> { if(File.Exists(lastOutput)) Process.Start(new ProcessStartInfo(lastOutput) {UseShellExecute=true}); };actions.Controls.Add(open);layout.Controls.Add(actions,0,3);
            DragEnter+=(s,e)=> { if(!running && e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect=DragDropEffects.Copy; };
            DragDrop+=(s,e)=> { if(running)return;foreach(string path in (string[])e.Data.GetData(DataFormats.FileDrop)) AddPath(path); };
            FormClosing+=(s,e)=> { if(running) {e.Cancel=true;CancelBuild();Append("Отмена запрошена. Дождитесь завершения текущей операции.");} };
        }
        Button EditButton(string text,int width,EventHandler handler) {
            var b=new Button {Text=text,Width=width,Height=31,FlatStyle=FlatStyle.Standard};b.Click+=handler;editingButtons.Add(b);return b;
        }
        void ChooseFiles() { using(var dialog=new OpenFileDialog {Title="Выберите файлы проверяющих",Filter="Книги Excel (*.xlsx)|*.xlsx",Multiselect=true}) if(dialog.ShowDialog(this)==DialogResult.OK) foreach(string f in dialog.FileNames) AddPath(f); }
        void ChooseFolder() {using(var dialog=new FolderBrowserDialog {Description="Выберите папку с сохранёнными файлами проверяющих (включая вложенные папки)"}) if(dialog.ShowDialog(this)==DialogResult.OK) AddPath(dialog.SelectedPath);}
        public void AddPath(string path) {
            try {
                if(Directory.Exists(path)) { foreach(string f in Directory.EnumerateFiles(path,"*.xlsx",SearchOption.AllDirectories).OrderBy(f=>f,StringComparer.OrdinalIgnoreCase)) AddPath(f);return; }
                path=Path.GetFullPath(path);
                if(!File.Exists(path) || !path.EndsWith(".xlsx",StringComparison.OrdinalIgnoreCase) || Path.GetFileName(path).StartsWith("~",StringComparison.Ordinal))return;
                if(paths.Any(p=>string.Equals(p,path,StringComparison.OrdinalIgnoreCase)))return;
                if(string.Equals(path,output.Text,StringComparison.OrdinalIgnoreCase))return;
                paths.Add(path);RefreshList();
            } catch(Exception ex) {Append("Не удалось добавить: "+path+". "+ex.Message);}
        }
        void RefreshList() {
            files.Items.Clear();for(int i=0;i<paths.Count;i++)files.Items.Add(new ListViewItem(new [] {(i+1).ToString(),Path.GetFileName(paths[i]),Path.GetDirectoryName(paths[i])}));
            totals.Text="Выбрано файлов: "+paths.Count;
        }
        void RemoveSelected() {foreach(int index in files.SelectedIndices.Cast<int>().OrderByDescending(i=>i).ToArray())paths.RemoveAt(index);RefreshList();}
        void MoveSelected(int delta) {if(files.SelectedIndices.Count!=1)return;int a=files.SelectedIndices[0],b=a+delta;if(b<0 || b>=paths.Count)return;string p=paths[a];paths[a]=paths[b];paths[b]=p;RefreshList();files.Items[b].Selected=true;files.Items[b].Focused=true;}
        void ChooseOutput() {using(var dialog=new OpenFileDialog {Title="Выберите сводный документ для дополнения",Filter="Книга Excel (*.xlsx)|*.xlsx",FileName=output.Text,CheckFileExists=true})if(dialog.ShowDialog(this)==DialogResult.OK)output.Text=dialog.FileName;}
        void Append(string message) {log.AppendText("\n"+DateTime.Now.ToString("HH:mm:ss")+"  "+message);log.SelectionStart=log.TextLength;log.ScrollToCaret();}
        void Ui(Action action) {if(!IsDisposed && IsHandleCreated)BeginInvoke(action);}
        void Busy(bool value) {
            running=value;build.Enabled=!value;cancel.Enabled=value;open.Enabled=!value && lastOutput!=null;
            foreach(var b in editingButtons)b.Enabled=!value;sheetName.Enabled=!value;output.Enabled=!value;asOf.Enabled=!value;files.Enabled=!value;currentDate.Enabled=!value;matchSurnames.Enabled=!value;forgetNames.Enabled=!value;
        }
        void StartBuild() {
            if(paths.Count==0) {MessageBox.Show(this,"Добавьте файлы проверяющих.",Text,MessageBoxButtons.OK,MessageBoxIcon.Information);return;}
            if(string.IsNullOrWhiteSpace(output.Text) || string.IsNullOrWhiteSpace(sheetName.Text)) {MessageBox.Show(this,"Укажите основной лист и итоговый файл.",Text);return;}
            string destination;
            try {destination=Path.GetFullPath(output.Text);}catch(Exception ex){MessageBox.Show(this,ex.Message,Text);return;}
            if(!File.Exists(destination)){MessageBox.Show(this,"Выберите существующий сводный документ, в который нужно добавить данные.",Text);return;}
            var selected=paths.ToList();string main=sheetName.Text.Trim();DateTime date=asOf.Value.Date;
            var options=new MergeOptions{CurrentDateForNewReviews=currentDate.Checked,MatchSurnames=matchSurnames.Checked,SamePerson=AskSamePerson,FillIncomplete=AskFillIncomplete};
            cancellation=new CancellationTokenSource();Busy(true);progress.Value=0;conflicts.DataSource=null;Append("Начата сборка "+selected.Count+" файлов.");
            var worker=new Thread(()=> {
                try {
                    var result=MergeEngine.Run(selected,destination,main,date,(n,s)=>Ui(()=>{progress.Value=n;Append(s);}),cancellation.Token,options);
                    Ui(()=>{lastOutput=result.Output;conflicts.DataSource=result.Conflicts;totals.Text="Документов: "+result.Documents+"   Проверено: "+result.Reviewed+"   Коробов проверено не полностью: "+result.IncompleteBoxes.Count+"   Конфликтов: "+result.Conflicts.Count+"   Проблем данных: "+result.Problems.Count;
                        Append("Дополнен сводный документ: "+result.Output);foreach(var merge in result.PeopleMerges)Append("Один проверяющий: "+merge);
                        foreach(var v in result.FilledBoxes)Append("Заполнены без замечаний остальные строки. "+v);foreach(var v in result.IncompleteBoxes)Append("Проверено не полностью, неотмеченные тома не входят в статистику. "+v);
                        foreach(var problem in result.Problems)Append(Path.GetFileName(problem.File)+", строка "+problem.Row+": "+problem.Detail);Busy(false);cancellation.Dispose();cancellation=null;});
                } catch(OperationCanceledException) {Ui(()=>{Append("Сборка отменена. Итоговый файл не заменён.");Busy(false);cancellation.Dispose();cancellation=null;});}
                catch(Exception ex) {Ui(()=>{Append("Ошибка: "+ex.Message);Busy(false);cancellation.Dispose();cancellation=null;MessageBox.Show(this,ex.Message,"Не удалось собрать свод",MessageBoxButtons.OK,MessageBoxIcon.Error);});}
            });worker.SetApartmentState(ApartmentState.STA);worker.IsBackground=true;worker.Start();
        }
        void SaveFlag(string name,bool value){try{UserSettings.SetFlag(name,value);}catch(Exception ex){Append("Не удалось сохранить настройку: "+ex.Message);}}
        void RefreshForget(){int n=UserSettings.SamePersonCount();forgetNames.Text=n==0?"":"Забыть ответы о фамилиях ("+n+")";forgetNames.Visible=n>0;}
        // Called from the build thread: asks on the window thread and remembers the answer for the next evenings.
        bool AskSamePerson(string raw,string known) {
            bool? saved=UserSettings.GetSamePerson(raw,known);
            if(saved.HasValue){Ui(()=>Append("«"+raw+"» и «"+known+"»: "+(saved.Value?"один проверяющий":"разные проверяющие")+" (сохранённый ответ)."));return saved.Value;}
            var answer=DialogResult.None;
            Invoke((Action)(()=>{answer=MessageBox.Show(this,"У проверяющих одна фамилия:\n\n«"+raw+"» — в файле проверяющего\n«"+known+"» — уже в своде\n\nЭто один и тот же человек?\n\nДа — объединить в справке и статистике.\nНет — считать разными людьми.\nОтмена — прервать сборку.\n\nОтвет запомнится для следующих сборок.","Сопоставление по фамилии",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question);}));
            if(answer==DialogResult.Cancel)throw new OperationCanceledException();
            bool same=answer==DialogResult.Yes;
            try{UserSettings.SetSamePerson(raw,known,same);}catch(Exception ex){Ui(()=>Append("Не удалось сохранить ответ: "+ex.Message));}
            Ui(RefreshForget);return same;
        }
        IList<IncompleteBox> AskFillIncomplete(IList<IncompleteBox> boxes) {
            IList<IncompleteBox> chosen=new List<IncompleteBox>();bool cancel=false;
            Invoke((Action)(()=>{using(var dialog=new IncompleteBoxesDialog(boxes)){var answer=dialog.ShowDialog(this);if(answer==DialogResult.OK)chosen=dialog.Selected;cancel=answer==DialogResult.Cancel;}}));
            if(cancel)throw new OperationCanceledException();
            return chosen;
        }
        void UpdateRule(){rule.Text="Правило: сохранять все отметки «1» и все разные тексты замечаний.\n"+(currentDate.Checked?"Новые проверки — с текущей датой; даты уже внесённых проверок сохраняются.":"Проверяющий и дата — из последней датированной проверки; при одинаковой дате — из последнего файла.")+"\nДанные дописываются в выбранный сводный документ. Статистика — на вкладке «Свод».";}
        void CancelBuild() {if(cancellation!=null) {cancellation.Cancel();cancel.Enabled=false;Append("Запрошена отмена.");}}
    }
    sealed class IncompleteBoxesDialog : Form {
        readonly CheckedListBox list = new CheckedListBox();
        public IncompleteBoxesDialog(IList<IncompleteBox> boxes) {
            Text="Короба проверены не полностью";Font=new Font("Segoe UI",10);ClientSize=new Size(940,540);MinimumSize=new Size(720,420);StartPosition=FormStartPosition.CenterParent;ShowInTaskbar=false;MinimizeBox=false;
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(14),ColumnCount=1,RowCount=3};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,70));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,48));
            layout.Controls.Add(new Label {Dock=DockStyle.Fill,Text="В коробах ниже отмечены не все строки: ИУЛ, фрагменты или тома без фамилии и даты. Обычно так быть не должно; неотмеченные тома не входят в статистику.\nЗаполнить остальные строки короба без замечаний? Им будут поставлены проверяющий, дата, вид и короб проверенной части; замечания не добавляются."},0,0);
            list.Dock=DockStyle.Fill;list.CheckOnClick=true;list.HorizontalScrollbar=true;list.IntegralHeight=false;foreach(var v in boxes)list.Items.Add(v,true);layout.Controls.Add(list,0,1);
            var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft,WrapContents=false,Padding=new Padding(0,6,0,0)};
            var mark=new Button {Text="Заполнить выбранные",Width=190,Height=33,DialogResult=DialogResult.OK};
            var skip=new Button {Text="Не заполнять",Width=130,Height=33,DialogResult=DialogResult.No};
            var stop=new Button {Text="Отменить сборку",Width=160,Height=33,DialogResult=DialogResult.Cancel};
            var none=new Button {Text="Снять все",Width=110,Height=33};none.Click+=(s,e)=>{for(int i=0;i<list.Items.Count;i++)list.SetItemChecked(i,false);};
            var all=new Button {Text="Выбрать все",Width=120,Height=33};all.Click+=(s,e)=>{for(int i=0;i<list.Items.Count;i++)list.SetItemChecked(i,true);};
            buttons.Controls.AddRange(new Control[] {mark,skip,stop,none,all});layout.Controls.Add(buttons,0,2);
            Controls.Add(layout);AcceptButton=mark;CancelButton=stop;
        }
        public IList<IncompleteBox> Selected { get { return list.CheckedItems.Cast<IncompleteBox>().ToList(); } }
    }
    static class UserSettings {
        static string PathName {get{return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"ReviewMerge","settings.json");}}
        static Dictionary<string,object> Load(){try{if(!File.Exists(PathName))return new Dictionary<string,object>();return new JavaScriptSerializer().Deserialize<Dictionary<string,object>>(File.ReadAllText(PathName,Encoding.UTF8))??new Dictionary<string,object>();}catch{return new Dictionary<string,object>();}}
        static void Save(Dictionary<string,object> values){Directory.CreateDirectory(Path.GetDirectoryName(PathName));File.WriteAllText(PathName,new JavaScriptSerializer().Serialize(values),new UTF8Encoding(false));}
        public static bool GetFlag(string name,bool fallback){object value;return Load().TryGetValue(name,out value)&&value is bool?(bool)value:fallback;}
        public static void SetFlag(string name,bool value){var values=Load();values[name]=value;Save(values);}
        static Dictionary<string,object> Answers(Dictionary<string,object> values){object map;return values.TryGetValue("SamePerson",out map)?map as Dictionary<string,object>??new Dictionary<string,object>():new Dictionary<string,object>();}
        static string PairKey(string raw,string known){return Rules.NameKey(raw)+" = "+Rules.NameKey(known);}
        public static bool? GetSamePerson(string raw,string known){object value;return Answers(Load()).TryGetValue(PairKey(raw,known),out value)&&value is bool?(bool?)(bool)value:null;}
        public static void SetSamePerson(string raw,string known,bool same){var values=Load();var answers=Answers(values);answers[PairKey(raw,known)]=same;values["SamePerson"]=answers;Save(values);}
        public static int SamePersonCount(){return Answers(Load()).Count;}
        public static void ForgetSamePerson(){var values=Load();values.Remove("SamePerson");Save(values);}
    }
}
