# Optional regression check using installed Microsoft Excel on artificial test_merge.py data.
$ErrorActionPreference='Stop'
$candidate=Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '.runs') -Directory | Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'template.xlsx')} | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if(!$candidate){throw 'Run test_merge.py first'}
$excelApp=$null;$bookCheck=$null
try {
    $excelApp=New-Object -ComObject Excel.Application
    $excelApp.Visible=$false;$excelApp.DisplayAlerts=$false;$excelApp.EnableEvents=$false;$excelApp.AskToUpdateLinks=$false;$excelApp.AutomationSecurity=3
    $bookCheck=$excelApp.Workbooks.Open((Join-Path $candidate.FullName 'result.xlsx'),0,$true)
    $summary=$bookCheck.Worksheets.Item('Свод');$main=$bookCheck.Worksheets.Item('Все загруженные файлы')
    $longNote='1. '+('Текст искусственного замечания. '*15)+"`n2. Второй пункт`n3. Третий пункт`n4. Четвёртый пункт"
    $main.Range('S7:S8').Value2=$longNote
    $excelApp.CalculateFullRebuild()
    if($summary.Range('B38').Value2 -ne 4){throw 'Long duplicate comments were counted incorrectly'}
    $main.Range('S8').Value2=$longNote+' Другое замечание'
    $excelApp.CalculateFullRebuild()
    if($summary.Range('B38').Value2 -ne 8){throw 'Different long comments were counted incorrectly'}
    $main.Range('S7:S8').Value2=''
    $excelApp.CalculateFullRebuild()
    if($summary.Range('B38').Value2 -ne 0){throw 'Blank comments were counted incorrectly'}
    if($bookCheck.Windows.Item(1).SelectedSheets.Count -ne 1){throw 'Grouped sheets remain'}
    Write-Output 'Excel tests passed: long text, fragment duplicates, changed text and blank comments'
} finally {
    if($bookCheck){$bookCheck.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($bookCheck)}
    if($excelApp){$excelApp.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp)}
}

$candidate=Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '.runs') -Directory | Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'boxes.xlsx')} | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$excelApp=$null;$bookCheck=$null
try {
    $excelApp=New-Object -ComObject Excel.Application
    $excelApp.Visible=$false;$excelApp.DisplayAlerts=$false;$excelApp.EnableEvents=$false;$excelApp.AskToUpdateLinks=$false;$excelApp.AutomationSecurity=3
    $bookCheck=$excelApp.Workbooks.Open((Join-Path $candidate.FullName 'result.xlsx'),0,$true)
    $summary=$bookCheck.Worksheets.Item('Свод');$main=$bookCheck.Worksheets.Item('Все загруженные файлы');$ref=$bookCheck.Worksheets.Item('Справка по томам')
    $excelApp.CalculateFullRebuild()
    if($ref.Range('F48').Value2 -ne ''){throw 'Incomplete box was counted'}
    if($ref.Range('F49').Value2 -ne ([datetime]'2026-10-04').ToOADate()){throw 'Complete box was not counted'}
    if($summary.Range('B51').Value2 -ne 1 -or $summary.Range('B52').Value2 -ne 2){throw 'Volume statistics by surname were incorrect'}
    if($summary.Range('G52').Value2 -ne 1){throw 'Completed box by surname was incorrect'}
    $main.Range('I9').Value2='Анна';$main.Range('J9').Formula='=DATE(2026,10,4)'
    $excelApp.CalculateFullRebuild()
    if($ref.Range('F48').Value2 -ne ([datetime]'2026-10-04').ToOADate()){throw 'Finishing last volume did not finish box'}
    if($summary.Range('B51').Value2 -ne 2){throw 'Newly reviewed volume was not counted'}
    $main.Range('I9').Value2='Борис';$excelApp.CalculateFullRebuild()
    if($summary.Range('B51').Value2 -ne 1 -or $summary.Range('B52').Value2 -ne 3){throw 'Surname edit did not update volume statistics'}
    $main.Range('L9').Formula='3';$excelApp.CalculateFullRebuild()
    if($ref.Range('F48').Value2 -ne ''){throw 'Changing box composition left an incorrect completion'}
    if($bookCheck.Windows.Item(1).SelectedSheets.Count -ne 1){throw 'Grouped sheets remain'}
    Write-Output 'Excel tests passed: reference formulas, complete boxes, volume statistics by surname and independent active sheet'
} finally {
    if($bookCheck){$bookCheck.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($bookCheck)}
    if($excelApp){$excelApp.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp)}
}
