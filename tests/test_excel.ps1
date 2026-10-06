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
    Write-Output 'Excel tests passed: long text, fragment duplicates, changed text and blank comments'
} finally {
    if($bookCheck){$bookCheck.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($bookCheck)}
    if($excelApp){$excelApp.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp)}
}
