$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$temp=Join-Path ([IO.Path]::GetTempPath()) ('update-tests-'+[guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($temp)|Out-Null
$compiler=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$exe=Join-Path $temp 'tests.exe'
& $compiler /nologo /codepage:65001 /target:exe (('/out:'+ $exe)) /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll (Join-Path $root 'Updates.cs') (Join-Path $root 'AppInfo.cs') (Join-Path $PSScriptRoot 'UpdatesTest.cs')
if($LASTEXITCODE -ne 0){throw 'Update test compilation failed'}
& $exe
if($LASTEXITCODE -ne 0){throw 'Update tests failed'}
