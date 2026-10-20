# Optional native Excel check after test_pd_recheck.py. Only disposable test books are edited.
$ErrorActionPreference='Stop'
$taskCandidate=Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '.runs') -Directory | Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'fragments.xlsx')} | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if(!$taskCandidate){throw 'Run test_pd_recheck.py first'}
$taskExcel=$null;$taskBook=$null
function Assert-Summary([int]$total,[int]$resolved,[int]$pending){
 $taskExcel.CalculateFullRebuild()
 $taskSheet=$taskBook.Worksheets.Item('СВОД')
 foreach($taskPair in @(@('D4',$total),@('H4',$resolved),@('L4',$pending))){
  if($taskSheet.Range($taskPair[0]).Value2 -notmatch (': '+$taskPair[1]+'$')){throw ('Incorrect '+$taskPair[0]+': '+$taskSheet.Range($taskPair[0]).Value2)}
 }
 foreach($taskName in @('СВОД','Стат. по типам','Стат. по зам.','Итоги работы')){
  $taskBad=$null;try{$taskBad=$taskBook.Worksheets.Item($taskName).UsedRange.SpecialCells(-4123,16)}catch{}
  if($null -ne $taskBad){throw ('Formula errors: '+$taskName+' '+$taskBad.Address())}
 }
}
try{
 $taskExcel=New-Object -ComObject Excel.Application
 $taskExcel.Visible=$false;$taskExcel.DisplayAlerts=$false;$taskExcel.EnableEvents=$false;$taskExcel.AskToUpdateLinks=$false;$taskExcel.AutomationSecurity=3
 $taskBook=$taskExcel.Workbooks.Open((Join-Path $taskCandidate.FullName 'marks.xlsx'),0,$true)
 Assert-Summary 11 4 7
 $taskMain=$taskBook.Worksheets.Item('Все загруженные файлы')
 foreach($taskMark in @('0.0','-0','0,00','0e3',([string][char]160+[char]9),'нет')){
  $taskMain.Range('V7').NumberFormat='@';$taskMain.Range('V7').Value2=$taskMark
  Assert-Summary 11 4 7
 }
 $taskMain.Range('V7').Value2=$true;Assert-Summary 11 5 6
 $taskMain.Range('V7').Value2=$false;Assert-Summary 11 4 7
 $taskMain.Range('N7:S7').Value2=0
 $taskMain.Range('T7').Value2=([string][char]160+[char]9)
 Assert-Summary 10 4 6
 $taskMain.Range('N7').Value2=1;$taskMain.Range('T7').Value2=''
 $taskMain.Range('I7').Value2=('Анна'+[char]160+[char]9)
 Assert-Summary 11 4 7
 if($taskBook.Worksheets.Item('Итоги работы').Range('D6').Value2 -ne 11){throw 'Whitespace changed reviewer totals'}
 $taskMain.Range('J7').Value2=5
 Assert-Summary 10 4 6
 $taskBook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook);$taskBook=$null
 $taskBook=$taskExcel.Workbooks.Open((Join-Path $taskCandidate.FullName 'fragments.xlsx'),0,$true)
 Assert-Summary 2 1 1
 $taskMain=$taskBook.Worksheets.Item('Все загруженные файлы')
 $taskMain.Range('V8').Value2=0;Assert-Summary 2 0 2
 $taskMain.Range('V9').Value2=$true;Assert-Summary 2 1 1
 $taskMain.Range('V9').Formula='=DATE(2026,10,29)';Assert-Summary 2 1 1
 Write-Output 'Native Excel passed: cached and live resolved counts, zero variants, booleans, whitespace, date bounds, fragment/IUL marks and future resolution dates'
}finally{
 if($null -ne $taskBook){$taskBook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskBook)}
 if($null -ne $taskExcel){$taskExcel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
