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
                    if (args.Length < 5) throw new ArgumentException("--batch результат.xlsx yyyy-MM-dd основной-лист файл1.xlsx [файл2.xlsx ...]");
                    DateTime date = DateTime.ParseExact(args[2], "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                    var logs = new List<string>();
                    var result = MergeEngine.Run(args.Skip(4).Select(Path.GetFullPath).ToList(), args[1], args[3], date, (n,s) => logs.Add(n+"%: "+s), CancellationToken.None);
                    var report = new { result.Files, result.Documents, result.Reviewed, result.WithIssues, result.Boxes, Conflicts=result.Conflicts.Count, Problems=result.Problems.Count, result.Output, Log=logs };
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
            ClientSize=new Size(1060,755);MinimumSize=new Size(980,750);StartPosition=FormStartPosition.CenterScreen;
            AutoScaleMode=AutoScaleMode.Dpi;AllowDrop=true;
            var layout=new TableLayoutPanel { Dock=DockStyle.Fill,Padding=new Padding(22),ColumnCount=1,RowCount=4 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,43));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,38));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));
            Controls.Add(layout);
            var titleBar=new FlowLayoutPanel{Dock=DockStyle.Fill,WrapContents=false};titleBar.Controls.Add(new Label {Text="Свод проверки документации",Font=new Font("Segoe UI",19,FontStyle.Bold),AutoSize=true,ForeColor=Color.FromArgb(28,46,65)});titleBar.Controls.Add(DesktopUpdates.Updates.Attach(this,()=>running));layout.Controls.Add(titleBar,0,0);
            layout.Controls.Add(new Label {Text="Объединение отметок и замечаний. Статистика рассчитывается формулами из основного листа.",AutoSize=true,ForeColor=Color.FromArgb(73,91,108)},0,1);
            tabs.Dock=DockStyle.Fill;layout.Controls.Add(tabs,0,2);
            var page=new TabPage("Сборка свода") {Padding=new Padding(14),BackColor=Color.White};tabs.TabPages.Add(page);
            var content=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=8};
            content.RowStyles.Add(new RowStyle(SizeType.Absolute,38));content.RowStyles.Add(new RowStyle(SizeType.Percent,48));
            content.RowStyles.Add(new RowStyle(SizeType.Absolute,40));content.RowStyles.Add(new RowStyle(SizeType.Absolute,66));
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
            rule.Dock=DockStyle.Fill;rule.ForeColor=Color.FromArgb(73,91,108);rule.Text="Правило: сохранять все отметки «1» и все разные тексты замечаний.\nПроверяющий и дата — из последней датированной проверки; при одинаковой дате — из последнего файла.\nВсе исходные записи и различия сохраняются в отдельных листах результата.";content.Controls.Add(rule,0,3);
            var save=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=3};save.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,120));save.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));save.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,115));
            save.Controls.Add(new Label {Text="Сохранить свод:",AutoSize=true,Anchor=AnchorStyles.Left},0,0);output.Dock=DockStyle.Fill;output.Text=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Свод_"+DateTime.Today.ToString("yyyy-MM-dd")+".xlsx");save.Controls.Add(output,1,0);
            save.Controls.Add(EditButton("Обзор…",105,(s,e)=>ChooseOutput()),2,0);content.Controls.Add(save,0,4);
            progress.Dock=DockStyle.Fill;progress.Maximum=100;progress.Margin=new Padding(0,3,0,6);content.Controls.Add(progress,0,5);
            log.Dock=DockStyle.Fill;log.ReadOnly=true;log.BackColor=Color.FromArgb(248,250,252);log.BorderStyle=BorderStyle.FixedSingle;log.Font=new Font("Segoe UI",9);
            log.Text="Добавьте сохранённые файлы проверяющих. Первый файл задаёт структуру и оформление результата.\nФормулы пересчитываются при открытии Excel. Исходные файлы не изменяются.";content.Controls.Add(log,0,6);
            totals.Dock=DockStyle.Fill;totals.Text="Файлы не выбраны";totals.ForeColor=Color.FromArgb(28,46,65);content.Controls.Add(totals,0,7);
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
        void ChooseOutput() {using(var dialog=new SaveFileDialog {Title="Сохранить общий свод",Filter="Книга Excel (*.xlsx)|*.xlsx",FileName=output.Text,OverwritePrompt=true})if(dialog.ShowDialog(this)==DialogResult.OK)output.Text=dialog.FileName;}
        void Append(string message) {log.AppendText("\n"+DateTime.Now.ToString("HH:mm:ss")+"  "+message);log.SelectionStart=log.TextLength;log.ScrollToCaret();}
        void Ui(Action action) {if(!IsDisposed && IsHandleCreated)BeginInvoke(action);}
        void Busy(bool value) {
            running=value;build.Enabled=!value;cancel.Enabled=value;open.Enabled=!value && lastOutput!=null;
            foreach(var b in editingButtons)b.Enabled=!value;sheetName.Enabled=!value;output.Enabled=!value;asOf.Enabled=!value;files.Enabled=!value;
        }
        void StartBuild() {
            if(paths.Count==0) {MessageBox.Show(this,"Добавьте файлы проверяющих.",Text,MessageBoxButtons.OK,MessageBoxIcon.Information);return;}
            if(string.IsNullOrWhiteSpace(output.Text) || string.IsNullOrWhiteSpace(sheetName.Text)) {MessageBox.Show(this,"Укажите основной лист и итоговый файл.",Text);return;}
            string destination;
            try {destination=Path.GetFullPath(output.Text);}catch(Exception ex){MessageBox.Show(this,ex.Message,Text);return;}
            if(paths.Any(p=>string.Equals(p,destination,StringComparison.OrdinalIgnoreCase))){MessageBox.Show(this,"Выберите отдельный итоговый файл: исходные файлы перезаписывать нельзя.",Text);return;}
            if(File.Exists(destination) && MessageBox.Show(this,"Заменить существующий итоговый файл? Предыдущая версия будет сохранена рядом с расширением .backup.",Text,MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes)return;
            var selected=paths.ToList();string main=sheetName.Text.Trim();DateTime date=asOf.Value.Date;
            cancellation=new CancellationTokenSource();Busy(true);progress.Value=0;conflicts.DataSource=null;Append("Начата сборка "+selected.Count+" файлов.");
            var worker=new Thread(()=> {
                try {
                    var result=MergeEngine.Run(selected,destination,main,date,(n,s)=>Ui(()=>{progress.Value=n;Append(s);}),cancellation.Token);
                    Ui(()=>{lastOutput=result.Output;conflicts.DataSource=result.Conflicts;totals.Text="Документов: "+result.Documents+"   Проверено: "+result.Reviewed+"   Конфликтов: "+result.Conflicts.Count+"   Проблем данных: "+result.Problems.Count;
                        Append("Сохранено: "+result.Output);if(result.Problems.Count>0)Append("Откройте лист «Проблемы данных»: строки без документа и некорректные реквизиты сохранены для разбора.");Busy(false);cancellation.Dispose();cancellation=null;});
                } catch(OperationCanceledException) {Ui(()=>{Append("Сборка отменена. Итоговый файл не заменён.");Busy(false);cancellation.Dispose();cancellation=null;});}
                catch(Exception ex) {Ui(()=>{Append("Ошибка: "+ex.Message);Busy(false);cancellation.Dispose();cancellation=null;MessageBox.Show(this,ex.Message,"Не удалось собрать свод",MessageBoxButtons.OK,MessageBoxIcon.Error);});}
            });worker.SetApartmentState(ApartmentState.STA);worker.IsBackground=true;worker.Start();
        }
        void CancelBuild() {if(cancellation!=null) {cancellation.Cancel();cancel.Enabled=false;Append("Запрошена отмена.");}}
    }
}
