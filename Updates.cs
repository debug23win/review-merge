using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace DesktopUpdates {
    public sealed class ReleaseInfo { public Version Version; public string Tag,Url,Digest; public long Size; }
    public static class Updates {
        public static readonly Version Current = Assembly.GetExecutingAssembly().GetName().Version;
        static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        static readonly string Repo = AppInfo.Repository;
        static string Asset { get { return AppInfo.Asset; } }
        static string[] Allowed { get { return AppInfo.UpdateFiles; } }
        public static ReleaseInfo ParseRelease(string json) {
            var r=(Dictionary<string,object>)Json.DeserializeObject(json);
            Version version; string tag=Convert.ToString(r["tag_name"]);
            if(!Regex.IsMatch(tag??"",@"^v?\d+\.\d+\.\d+$") || !Version.TryParse(tag.TrimStart('v'),out version))return null;
            if(Convert.ToBoolean(r["draft"]) || Convert.ToBoolean(r["prerelease"]))return null;
            foreach(var item in (object[])r["assets"]){var a=(Dictionary<string,object>)item;
                if(Convert.ToString(a["name"])!=Asset)continue;
                string url=Convert.ToString(a["browser_download_url"]),digest=a.ContainsKey("digest")?Convert.ToString(a["digest"]):"";
                if(!url.StartsWith("https://github.com/"+Repo+"/releases/download/",StringComparison.Ordinal))throw new InvalidDataException("Неверный адрес пакета обновления.");
                if(!Regex.IsMatch(digest??"",@"^sha256:[0-9a-fA-F]{64}$"))throw new InvalidDataException("У релиза отсутствует SHA-256 пакета. Обновление не установлено.");
                long size=Convert.ToInt64(a["size"]);if(size<=0 || size>50000000)throw new InvalidDataException("Неверный размер пакета обновления.");
                return new ReleaseInfo {Version=version,Tag=tag,Url=url,Digest=digest.Substring(7),Size=size};
            }return null;
        }
        static HttpWebRequest Request(string url) {
            ServicePointManager.SecurityProtocol=SecurityProtocolType.Tls12;
            var request=(HttpWebRequest)WebRequest.Create(url);request.UserAgent=AppInfo.Product+"/"+Current;request.Timeout=20000;request.ReadWriteTimeout=20000;
            request.Accept="application/vnd.github+json";request.Headers["X-GitHub-Api-Version"]="2026-03-10";return request;
        }
        public static ReleaseInfo Check() {
            try{using(var response=Request("https://api.github.com/repos/"+Repo+"/releases/latest").GetResponse())using(var stream=response.GetResponseStream())using(var reader=new StreamReader(stream,Encoding.UTF8))return ParseRelease(reader.ReadToEnd());}
            catch(WebException ex){var r=ex.Response as HttpWebResponse;if(r!=null&&r.StatusCode==HttpStatusCode.NotFound)return null;throw;}
        }
        public static void Verify(string path,string expected) {
            string hash;using(var h=SHA256.Create())using(var f=File.OpenRead(path))hash=BitConverter.ToString(h.ComputeHash(f)).Replace("-","");
            if(!hash.Equals(expected,StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("SHA-256 скачанного пакета не совпадает. Обновление отклонено.");
        }
        public static void Extract(string archive,string stage) {
            Directory.CreateDirectory(stage);long total=0;var seen=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using(var f=File.OpenRead(archive))using(var zip=new ZipArchive(f,ZipArchiveMode.Read))foreach(var entry in zip.Entries){
                if(entry.FullName!=Path.GetFileName(entry.FullName)||entry.FullName.Contains('\\'))throw new InvalidDataException("В пакете недопустимый путь.");
                if(!Allowed.Contains(entry.FullName,StringComparer.OrdinalIgnoreCase))continue;
                if(!seen.Add(entry.FullName))throw new InvalidDataException("В пакете повторяется файл.");
                total+=entry.Length;if(total>50000000)throw new InvalidDataException("Пакет обновления слишком большой.");
                using(var input=entry.Open())using(var output=File.Create(Path.Combine(stage,entry.FullName)))input.CopyTo(output);
            }
            if(!File.Exists(Path.Combine(stage,AppInfo.Executable)))throw new InvalidDataException("В пакете отсутствует программа.");
        }
        public static string Download(ReleaseInfo release,string cacheRoot=null) {
            string cache=cacheRoot??Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"DocumentReviewTools",AppInfo.Product,"updates");
            string root=Path.Combine(cache,Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
            string zip=Path.Combine(root,"package.zip"),stage=Path.Combine(root,"package");
            using(var response=Request(release.Url).GetResponse())using(var input=response.GetResponseStream())using(var output=File.Create(zip)){
                byte[] buffer=new byte[65536];int n;long received=0;while((n=input.Read(buffer,0,buffer.Length))>0){received+=n;if(received>50000000)throw new InvalidDataException("Пакет слишком большой.");output.Write(buffer,0,n);}if(received!=release.Size)throw new InvalidDataException("Пакет скачан не полностью.");
            }
            Verify(zip,release.Digest);Extract(zip,stage);
            var actual=AssemblyName.GetAssemblyName(Path.Combine(stage,AppInfo.Executable)).Version;
            if(actual.Major!=release.Version.Major||actual.Minor!=release.Version.Minor||actual.Build!=release.Version.Build)throw new InvalidDataException("Версия программы в пакете не соответствует релизу.");
            return stage;
        }
        public static LinkLabel Attach(Form form,Func<bool> busy) {
            var link=new LinkLabel {AutoSize=true,Text="v"+Current.ToString(3)+" · Обновления",Margin=new Padding(24,10,0,0),LinkColor=Color.FromArgb(35,80,115)};
            ReleaseInfo latest=null;bool checking=false;
            Action<bool> check=null;
            check=async manual=>{
                if(checking)return;checking=true;link.Enabled=false;
                try{latest=await Task.Run(()=>Check());
                    if(latest!=null&&latest.Version>new Version(Current.ToString(3)))link.Text="Установить v"+latest.Version.ToString(3);
                    else {link.Text="v"+Current.ToString(3)+" · Обновления";if(manual)MessageBox.Show(form,"Установлена последняя опубликованная версия.","Обновления",MessageBoxButtons.OK,MessageBoxIcon.Information);}
                }catch(Exception ex){link.Text="v"+Current.ToString(3)+" · Повторить проверку";if(manual)MessageBox.Show(form,"Проверка обновлений недоступна: "+ex.Message,"Обновления",MessageBoxButtons.OK,MessageBoxIcon.Information);}
                finally{checking=false;if(!form.IsDisposed)link.Enabled=true;}
            };
            link.LinkClicked+=async (s,e)=>{
                if(busy()){MessageBox.Show(form,"Дождитесь завершения текущей проверки или свода.","Обновления");return;}
                if(latest==null||latest.Version<=new Version(Current.ToString(3))){check(true);return;}
                if(MessageBox.Show(form,"Установить версию "+latest.Version.ToString(3)+"? Программа перезапустится. Результаты проверки сохраните перед обновлением.","Обновление",MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes)return;
                link.Enabled=false;link.Text="Загрузка обновления…";form.Enabled=false;
                try{
                    string destination=Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
                    string probe=Path.Combine(destination,".update-write-"+Guid.NewGuid().ToString("N"));using(File.Create(probe)){}File.Delete(probe);
                    string stage=await Task.Run(()=>Download(latest));string helper=Path.Combine(Path.GetDirectoryName(stage),"update-helper.exe");File.Copy(Application.ExecutablePath,helper);
                    Process.Start(new ProcessStartInfo(helper,"--apply-update "+Process.GetCurrentProcess().Id+" \""+destination.TrimEnd('\\')+"\" \""+stage+"\""){UseShellExecute=false,CreateNoWindow=true,WindowStyle=ProcessWindowStyle.Hidden});form.Close();
                }catch(Exception ex){form.Enabled=true;MessageBox.Show(form,"Обновление не установлено: "+ex.Message+"\nРаспакуйте программу в папку, где разрешена запись, и повторите попытку.","Обновления",MessageBoxButtons.OK,MessageBoxIcon.Error);link.Text="Установить v"+latest.Version.ToString(3);link.Enabled=true;}
            };
            form.Shown+=(s,e)=>{if(Environment.GetCommandLineArgs().Length==1)check(false);};return link;
        }
        public static int Apply(string[] args) {
            if(args.Length!=4)return 1;string destination=Path.GetFullPath(args[2]),stage=Path.GetFullPath(args[3]);
            string backup=Path.Combine(destination,".update-backup-"+DateTime.Now.ToString("yyyyMMddHHmmss"));var changed=new List<string>();
            try{
                int pid=int.Parse(args[1]);try{using(var old=Process.GetProcessById(pid))if(!old.WaitForExit(60000))throw new IOException("Предыдущий экземпляр программы не завершился.");}catch(ArgumentException){}
                if(!File.Exists(Path.Combine(stage,AppInfo.Executable)))throw new FileNotFoundException("Пакет обновления отсутствует.");Directory.CreateDirectory(backup);
                foreach(string name in Allowed){string src=Path.Combine(stage,name),dst=Path.Combine(destination,name);if(!File.Exists(src))continue;
                    if(!Path.GetFullPath(dst).StartsWith(destination.TrimEnd('\\')+"\\",StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Недопустимая папка обновления.");
                    if(File.Exists(dst))File.Copy(dst,Path.Combine(backup,name),true);changed.Add(name);File.Copy(src,dst,true);
                }
                File.WriteAllText(Path.Combine(backup,"Обновление.txt"),"Обновление установлено "+DateTime.Now+". Предыдущие файлы сохранены в этой папке.",Encoding.UTF8);
                Process.Start(new ProcessStartInfo(Path.Combine(destination,AppInfo.Executable)){UseShellExecute=true,WorkingDirectory=destination});return 0;
            }catch(Exception ex){foreach(string name in changed){string prior=Path.Combine(backup,name);if(File.Exists(prior))try{File.Copy(prior,Path.Combine(destination,name),true);}catch{}}
                MessageBox.Show("Обновление не завершено. Предыдущая версия восстановлена, если её файлы были скопированы.\n"+ex.Message,"Обновление",MessageBoxButtons.OK,MessageBoxIcon.Error);return 1;}
        }
    }
}
