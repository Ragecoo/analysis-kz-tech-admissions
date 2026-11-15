param([string]$Definitions, [string]$Stage)
$ErrorActionPreference = 'Stop'
$configuration = Get-Content -LiteralPath $Definitions -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
$null = New-Item -ItemType Directory -Path $Stage -Force
$stagePath = (Resolve-Path -LiteralPath $Stage).Path
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
try {
    $items = $configuration.GetEnumerator() | Sort-Object Key -Descending
    foreach ($item in $items) {
        $name = $item.Key
        $target = Join-Path $stagePath $name
        if (Test-Path -LiteralPath $target) { continue }
        $source = (Resolve-Path -LiteralPath (Join-Path 'excel/bns-5-let' $name)).Path
        $book = $excel.Workbooks.Open($source, 0, $true)
        try {
            $result = $book.Worksheets.Item('итоги')
            $table = $result.ListObjects.Item(1)
            $count = $table.ListRows.Count
            foreach ($query in $item.Value.queries.GetEnumerator()) {
                if ($query.Key -in @('исходник БНС','госзаказ БНС')) {
                    $book.Queries.Item($query.Key).Formula = $query.Value
                } else {
                    $null = $book.Queries.Add($query.Key, $query.Value)
                }
            }
            foreach ($sheet in $item.Value.oldSheets) { $book.Worksheets.Item($sheet).Delete() }
            "Обновление: $name"
            $null = $table.QueryTable.Refresh($false)
            if ($table.ListRows.Count -ne $count) {
                throw "Изменилось число строк в $name : $count → $($table.ListRows.Count)"
            }
            $result.Activate()
            $book.SaveAs($target, 51)
            "$name : $count строк после обновления из файлов БНС"
        }
        catch { Write-Output "Ошибка $name : $($_.Exception.Message)" }
        finally { $book.Close($false) }
    }
}
finally { $excel.Quit() }
