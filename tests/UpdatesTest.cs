using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using DesktopUpdates;
class UpdatesTest {
    static void ExpectFailure(Action action){try{action();}catch(InvalidDataException){return;}throw new Exception("Invalid package accepted");}
    static int Main(){
        string root=Path.Combine(Path.GetTempPath(),"review-update-test-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
        string digest=new string('a',64);
        string json="{\"tag_name\":\"v1.2.3\",\"draft\":false,\"prerelease\":false,\"assets\":[{\"name\":\""+AppInfo.Asset+"\",\"browser_download_url\":\"https://github.com/"+AppInfo.Repository+"/releases/download/v1.2.3/"+AppInfo.Asset+"\",\"size\":100,\"digest\":\"sha256:"+digest+"\"}]}";
        if(Updates.ParseRelease(json).Version!=new Version(1,2,3))throw new Exception("Release version incorrect");
        if(Updates.ParseRelease(json.Replace("\"prerelease\":false","\"prerelease\":true"))!=null)throw new Exception("Prerelease accepted");
        ExpectFailure(()=>Updates.ParseRelease(json.Replace("sha256:"+digest,"")));
        ExpectFailure(()=>Updates.ParseRelease(json.Replace("https://github.com/","https://example.invalid/")));
        string file=Path.Combine(root,"bytes");File.WriteAllText(file,"fixture",Encoding.UTF8);string hash;
        using(var h=SHA256.Create())using(var f=File.OpenRead(file))hash=BitConverter.ToString(h.ComputeHash(f)).Replace("-","");
        Updates.Verify(file,hash);ExpectFailure(()=>Updates.Verify(file,new string('0',64)));
        string zip=Path.Combine(root,"valid.zip");using(var f=File.Create(zip))using(var z=new ZipArchive(f,ZipArchiveMode.Create)){using(var w=new StreamWriter(z.CreateEntry(AppInfo.Executable).Open()))w.Write("fixture");if(AppInfo.UpdateFiles.Length>2)using(var w=new StreamWriter(z.CreateEntry("Recognize.ps1").Open()))w.Write("fixture");using(var w=new StreamWriter(z.CreateEntry("user-data.xlsx").Open()))w.Write("must not install");}
        string stage=Path.Combine(root,"stage");Updates.Extract(zip,stage);if(File.Exists(Path.Combine(stage,"user-data.xlsx")))throw new Exception("Unexpected data installed");
        string bad=Path.Combine(root,"bad.zip");using(var f=File.Create(bad))using(var z=new ZipArchive(f,ZipArchiveMode.Create))using(var w=new StreamWriter(z.CreateEntry("../escape.exe").Open()))w.Write("bad");
        ExpectFailure(()=>Updates.Extract(bad,Path.Combine(root,"bad-stage")));
        // Download: a failed attempt leaves no folder in the cache, folders older than an hour are removed, a recent one is kept.
        string cache=Path.Combine(root,"cache"),old=Path.Combine(cache,"old"),recent=Path.Combine(cache,"recent");Directory.CreateDirectory(old);Directory.CreateDirectory(recent);
        File.WriteAllText(Path.Combine(old,"package.zip"),"old");Directory.SetLastWriteTimeUtc(old,DateTime.UtcNow.AddDays(-2));
        string good=Path.Combine(root,"good.zip");using(var f=File.Create(good))using(var z=new ZipArchive(f,ZipArchiveMode.Create))z.CreateEntryFromFile(typeof(UpdatesTest).Assembly.Location,AppInfo.Executable);
        byte[] package=File.ReadAllBytes(good);string packageHash;using(var h=SHA256.Create())packageHash=BitConverter.ToString(h.ComputeHash(package)).Replace("-","");
        var version=typeof(UpdatesTest).Assembly.GetName().Version;var release=new ReleaseInfo{Version=new Version(version.Major,version.Minor,version.Build),Url="test",Digest=packageHash,Size=package.Length};
        Func<string[]> folders=()=>Directory.GetDirectories(cache).Select(Path.GetFileName).OrderBy(n=>n).ToArray();
        ExpectFailure(()=>Updates.Download(new ReleaseInfo{Version=release.Version,Url="test",Digest=new string('0',64),Size=package.Length},cache,()=>new MemoryStream(package)));
        if(!folders().SequenceEqual(new[]{"recent"}))throw new Exception("Cache after a bad hash: "+string.Join(",",folders()));
        ExpectFailure(()=>Updates.Download(new ReleaseInfo{Version=release.Version,Url="test",Digest=packageHash,Size=package.Length+1},cache,()=>new MemoryStream(package)));
        ExpectFailure(()=>Updates.Download(new ReleaseInfo{Version=new Version(version.Major+1,0,0),Url="test",Digest=packageHash,Size=package.Length},cache,()=>new MemoryStream(package)));
        if(!folders().SequenceEqual(new[]{"recent"}))throw new Exception("Cache after failed downloads: "+string.Join(",",folders()));
        string staged=Updates.Download(release,cache,()=>new MemoryStream(package));
        if(!File.Exists(Path.Combine(staged,AppInfo.Executable))||folders().Length!=2||!folders().Contains("recent"))throw new Exception("Cache after a good download: "+string.Join(",",folders()));
        Console.WriteLine("Update tests passed: stable version, host, digest, checksum, allowlist, ZIP paths and the download cache");return 0;
    }
}
