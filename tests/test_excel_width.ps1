# Optional oracle requiring installed Microsoft Excel and artificial test_recreated_reports.py output.
# The workbook is opened read-only, changed only in memory and closed without saving.
param([string]$WorkbookPath)
$ErrorActionPreference='Stop'
if(!$WorkbookPath){
    $runs=Join-Path $PSScriptRoot '.runs'
    if(!(Test-Path -LiteralPath $runs)){throw 'Run test_recreated_reports.py first'}
    $candidate=Get-ChildItem -LiteralPath $runs -Directory | Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'reviewer-history.xlsx')} | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if(!$candidate){throw 'Run test_recreated_reports.py first'}
    $WorkbookPath=Join-Path $candidate.FullName 'reviewer-history.xlsx'
}
$WorkbookPath=(Resolve-Path -LiteralPath $WorkbookPath).Path
$beforeHash=(Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash
$excelApp=$null;$books=$null;$book=$null;$sheets=$null;$sheet=$null;$columns=$null;$cells=$null;$oracle=$null;$used=$null;$usedColumns=$null
try {
    $excelApp=New-Object -ComObject Excel.Application
    $excelApp.Visible=$false;$excelApp.DisplayAlerts=$false;$excelApp.EnableEvents=$false;$excelApp.AskToUpdateLinks=$false;$excelApp.AutomationSecurity=3
    $books=$excelApp.Workbooks
    $book=$books.Open($WorkbookPath,0,$true)
    if(!$book.ReadOnly){throw 'Excel did not open the artificial workbook read-only'}
    $sheets=$book.Worksheets;$sheet=$sheets.Item('Итоги работы')
    $columns=$sheet.Columns;$cells=$sheet.Cells
    $used=$sheet.UsedRange;$usedColumns=$used.Columns
    $lastUsed=[int]$used.Column+[int]$usedColumns.Count-1
    if($lastUsed -ge 16384){throw 'The oracle column XFD is not unused'}
    $lastReportColumn=0
    for($index=3;$index -le $lastUsed;$index+=2){
        $header=$null
        try {
            $header=$cells.Item(4,$index)
            if([string]$header.Value2 -ceq 'ВСЕГО'){$lastReportColumn=$index+1;break}
        } finally {if($header){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($header)}}
    }
    if($lastReportColumn -lt 4){throw 'The final ВСЕГО reviewer group was not found'}
    $oracle=$columns.Item(16384)
    $oracle.ColumnWidth=8.9
    $oracleCharacters=[double]$oracle.ColumnWidth
    $oraclePoints=[double]$oracle.Width
    if($oracleCharacters -le 0 -or $oraclePoints -le 0){throw 'Excel returned an invalid oracle width'}
    # Excel quantizes ColumnWidth according to Normal's font. Do not demand exactly 8.9.
    # Compare both its character-unit result and physical point width with the oracle.
    for($index=3;$index -le $lastReportColumn;$index++){
        $column=$null
        try {
            $column=$columns.Item($index)
            $actualCharacters=[double]$column.ColumnWidth
            $actualPoints=[double]$column.Width
            if([Math]::Abs($actualCharacters-$oracleCharacters) -gt 0.000001 -or [Math]::Abs($actualPoints-$oraclePoints) -gt 0.000001){
                throw ('Column {0}: ColumnWidth={1}, Width={2}pt; Excel oracle ColumnWidth={3}, Width={4}pt' -f $index,$actualCharacters,$actualPoints,$oracleCharacters,$oraclePoints)
            }
        } finally {if($column){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($column)}}
    }
    Write-Output ('Excel width oracle passed: {0} report columns match ColumnWidth={1}, Width={2}pt' -f ($lastReportColumn-2),$oracleCharacters,$oraclePoints)
} finally {
    foreach($item in @($oracle,$usedColumns,$used,$cells,$columns,$sheet,$sheets)){
        if($null -ne $item){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)}
    }
    if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if($books){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($books)}
    if($excelApp){$excelApp.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp)}
}
$afterHash=(Get-FileHash -LiteralPath $WorkbookPath -Algorithm SHA256).Hash
if($afterHash -ne $beforeHash){throw 'The Excel oracle modified the workbook on disk'}
Write-Output 'Verified: source workbook bytes unchanged'
