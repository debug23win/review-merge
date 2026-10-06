param([ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version='1.3.2')
$ErrorActionPreference='Stop'
Push-Location -LiteralPath $PSScriptRoot
$infoPath=Join-Path $PSScriptRoot 'AppInfo.cs'
$infoText=[IO.File]::ReadAllText($infoPath,[Text.Encoding]::UTF8)
$infoText=[regex]::Replace($infoText,'Assembly(Version|FileVersion)\("[0-9.]+"\)',('Assembly$1("'+$Version+'.0")'))
[IO.File]::WriteAllText($infoPath,$infoText,[Text.UTF8Encoding]::new($false))
$compiler=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
if(!(Test-Path -LiteralPath $compiler)){$compiler=Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319/csc.exe'}
& $compiler /nologo /target:winexe /platform:anycpu /optimize+ /codepage:65001 '/out:Свод_проверки.exe' /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Xml.Linq.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:System.Web.Extensions.dll Program.cs XlsxReader.cs MergeEngine.cs XlsxWriter.cs SummaryWriter.cs ReferenceWriter.cs RemarkCounter.cs Updates.cs AppInfo.cs
if($LASTEXITCODE -ne 0){throw 'Compilation failed'}
$dist=Join-Path $PSScriptRoot 'dist'
[IO.Directory]::CreateDirectory($dist)|Out-Null
$package=Join-Path $dist 'review-merge-windows.zip'
Compress-Archive -LiteralPath @((Join-Path $PSScriptRoot 'Свод_проверки.exe'),(Join-Path $PSScriptRoot 'README.md')) -DestinationPath $package -Force
$hash=(Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText(($package+'.sha256'),$hash+'  review-merge-windows.zip'+[Environment]::NewLine,[Text.Encoding]::ASCII)
Write-Output $package
Pop-Location
