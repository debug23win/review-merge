$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$app=Join-Path $root 'Свод_проверки.exe'
if(!(Test-Path -LiteralPath $app)){throw 'Run build.ps1 before the report width tests'}
$temp=Join-Path ([IO.Path]::GetTempPath()) ('report-width-tests-'+[guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($temp)|Out-Null
try {
    $compiler=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if(!(Test-Path -LiteralPath $compiler)){$compiler=Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319/csc.exe'}
    $exe=Join-Path $temp 'tests.exe'
    & $compiler /nologo /codepage:65001 /target:exe (('/out:'+ $exe)) /r:System.Drawing.dll /r:System.Xml.Linq.dll (Join-Path $PSScriptRoot 'ReportWidthTest.cs')
    if($LASTEXITCODE -ne 0){throw 'Report width test compilation failed'}
    & $exe $app
    if($LASTEXITCODE -ne 0){throw 'Report width tests failed'}
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force
}
