param([string]$DataPath, [string]$OutputPath)
$ErrorActionPreference = 'Stop'
$books = Get-Content -LiteralPath $DataPath -Raw -Encoding utf8 | ConvertFrom-Json -NoEnumerate
$null = New-Item -ItemType Directory -Path $OutputPath -Force
$out = (Resolve-Path -LiteralPath $OutputPath).Path
$numeric = @(
    'Строка БНС', 'Принято всего', 'Принято женщин', 'Численность всего', 'Численность женщин',
    'Выпуск всего', 'Выпуск женщин', 'По госзаказу всего', 'По госзаказу женщин', 'Численность',
    'Контингент', 'Продолжили обучение', 'Продолжили в вузах', 'Продолжили в СУЗах',
    'Трудоустроены', 'Трудоустроены с очной формы'
)

function Set-Layout($sheet, $table) {
    $table.TableStyle = 'TableStyleMedium7'
    $null = $table.Range.Columns.AutoFit()
    foreach ($column in $table.ListColumns) {
        if ($column.Range.ColumnWidth -gt 42) { $column.Range.ColumnWidth = 42 }
        if ($column.Range.ColumnWidth -lt 12) { $column.Range.ColumnWidth = 12 }
        if ($column.Name -match '^Доля|^Темп') {
            try { $sheet.Columns.Item($column.Index).NumberFormatLocal = '0,00%' }
            catch { Write-Output "Процентный формат: $($column.Name)" }
        }
    }
    $table.HeaderRowRange.WrapText = $true
    $table.HeaderRowRange.RowHeight = 32
    $sheet.Activate()
    $sheet.Application.ActiveWindow.SplitRow = 1
    $sheet.Application.ActiveWindow.FreezePanes = $true
    $null = $sheet.Range('A1').Select()
}

function Add-SourceTable($workbook, $sheet, $data, $index) {
    $head = $data[0]
    $rows = $data[1]
    $grid = [object[,]]::new($rows.Count + 1, $head.Count)
    for ($j=0; $j -lt $head.Count; $j++) {
        $grid[0,$j] = [string]$head[$j]
        if ($head[$j] -in ('Код','КАТО','Учебный год','Лист БНС')) {
            $sheet.Columns.Item($j+1).NumberFormat = '@'
        }
    }
    for ($i=0; $i -lt $rows.Count; $i++) {
        for ($j=0; $j -lt $head.Count; $j++) {
            $value = $rows[$i][$j]
            if ($null -ne $value) {
                if ($head[$j] -in $numeric -and "$value" -match '^-?\d+$') { $grid[($i+1),$j] = [double]$value }
                elseif ("$value" -ne '') { $grid[($i+1),$j] = [string]$value }
            }
        }
    }
    $sheet.Name = if ($index -eq 1) { 'исходные данные' } else { 'исходные госзаказ' }
    $range = $sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item($rows.Count+1,$head.Count))
    $range.Value2 = $grid
    $table = $sheet.ListObjects.Add(1,$range,[Type]::Missing,1)
    $table.Name = "Источник$index"
    Set-Layout $sheet $table
    $schema = ($head | ForEach-Object {
        $type = if ($_ -in $numeric) { 'Int64.Type' } else { 'type text' }
        '{"'+$_+'", '+$type+'}'
    }) -join ', '
    $formula = @"
let
    Источник = Excel.CurrentWorkbook(){[Name="Источник$index"]}[Content],
    #"Измененный тип" = Table.TransformColumnTypes(Источник, {$schema}),
    #"Удалены пустые строки" = Table.SelectRows(#"Измененный тип", each [Учебный год] <> null)
in
    #"Удалены пустые строки"
"@
    $query = if ($index -eq 1) { 'исходник БНС' } else { 'госзаказ БНС' }
    $null = $workbook.Queries.Add($query,$formula)
}

$formulas = @(
@'
let
    Источник = #"исходник БНС",
    #"Объединены приём и госзаказ" = Table.NestedJoin(Источник, {"Учебный год","Уровень","Код"}, #"госзаказ БНС", {"Учебный год","Уровень","Код"}, "Госзаказ", JoinKind.LeftOuter),
    #"Развернут госзаказ" = Table.ExpandTableColumn(#"Объединены приём и госзаказ", "Госзаказ", {"По госзаказу всего","По госзаказу женщин","Файл БНС","Книга в архиве","Лист БНС","Строка БНС"}, {"По госзаказу всего","По госзаказу женщин","Файл госзаказа","Книга госзаказа","Лист госзаказа","Строка госзаказа"}),
    #"Добавлена техническая область" = Table.AddColumn(#"Развернут госзаказ", "Техническая область", each let код = if [Уровень] = "Вузы" then Text.Range([Код],2,2) else Text.Start([Код],2) in if код = "06" then "ИКТ" else if код = "07" then "Инженерные, обрабатывающие и строительные отрасли" else "Остальные", type text),
    #"Добавлена доля госзаказа" = Table.AddColumn(#"Добавлена техническая область", "Доля госзаказа", each if [Принято всего] = null or [Принято всего] = 0 then null else [По госзаказу всего]/[Принято всего], Percentage.Type)
in
    #"Добавлена доля госзаказа"
'@,
@'
let
    Источник = #"исходник БНС",
    #"Добавлен год начала" = Table.AddColumn(Источник, "Год", each Number.FromText(Text.Start([Учебный год],4)), Int64.Type),
    #"Добавлен предыдущий год" = Table.AddColumn(#"Добавлен год начала", "Предыдущий год", each [Год]-1, Int64.Type),
    #"Найден прошлогодний приём" = Table.NestedJoin(#"Добавлен предыдущий год", {"Уровень","Регион","Предыдущий год"}, #"Добавлен год начала", {"Уровень","Регион","Год"}, "Прошлый год", JoinKind.LeftOuter),
    #"Развернут прошлогодний приём" = Table.ExpandTableColumn(#"Найден прошлогодний приём", "Прошлый год", {"Принято всего"}, {"Принято годом ранее"}),
    #"Добавлено изменение" = Table.AddColumn(#"Развернут прошлогодний приём", "Изменение в людях", each [Принято всего]-[Принято годом ранее], Int64.Type),
    #"Добавлен темп изменения" = Table.AddColumn(#"Добавлено изменение", "Темп изменения", each if [Принято годом ранее] = null or [Принято годом ранее] = 0 then null else [Изменение в людях]/[Принято годом ранее], Percentage.Type),
    #"Отмечены изменения границ" = Table.AddColumn(#"Добавлен темп изменения", "Сопоставимость", each if [Год]=2022 and List.Contains({"Абай","Жетісу","Ұлытау","Восточно-Казахстанская","Алматинская","Карагандинская"},[Регион]) then "Проверить изменение границ" else null, type text),
    #"Удалены вспомогательные столбцы" = Table.RemoveColumns(#"Отмечены изменения границ", {"Год","Предыдущий год"})
in
    #"Удалены вспомогательные столбцы"
'@,
@'
let
    Источник = #"исходник БНС",
    #"Добавлена техническая область" = Table.AddColumn(Источник, "Техническая область", each if Text.StartsWith([Код],"06") then "ИКТ" else "Инженерные, обрабатывающие и строительные отрасли", type text),
    #"Записи отсортированы" = Table.Sort(#"Добавлена техническая область",{{"Учебный год",Order.Ascending},{"Местность",Order.Ascending},{"Код",Order.Ascending}})
in
    #"Записи отсортированы"
'@,
@'
let
    Источник = #"исходник БНС",
    #"Сгруппированные строки" = Table.Group(Источник, {"Учебный год","Уровень","Регион","Возрастная группа","Файл БНС","Книга в архиве","Лист БНС","Строка БНС"}, {{"Численность",each if List.NonNullCount([Численность])=List.Count([Численность]) then List.Sum([Численность]) else null,Int64.Type},{"Контингент",each List.First([Контингент]),Int64.Type}}),
    #"Добавлена доля возрастной группы" = Table.AddColumn(#"Сгруппированные строки", "Доля возрастной группы", each if [Контингент]=null or [Контингент]=0 then null else [Численность]/[Контингент], Percentage.Type)
in
    #"Добавлена доля возрастной группы"
'@,
@'
let
    Источник = #"исходник БНС",
    #"Сгруппированные строки" = Table.Group(Источник, {"Учебный год","Уровень","Регион"}, {{"Общий приём",each if List.NonNullCount([Принято всего])=List.Count([Принято всего]) then List.Sum([Принято всего]) else null,Int64.Type}}),
    #"Объединены приём и итоги" = Table.NestedJoin(Источник, {"Учебный год","Уровень","Регион"}, #"Сгруппированные строки", {"Учебный год","Уровень","Регион"}, "Итоги", JoinKind.LeftOuter),
    #"Развернут итог" = Table.ExpandTableColumn(#"Объединены приём и итоги", "Итоги", {"Общий приём"}),
    #"Добавлена доля" = Table.AddColumn(#"Развернут итог", "Доля приёма", each if [Общий приём]=null or [Общий приём]=0 then null else [Принято всего]/[Общий приём], Percentage.Type)
in
    #"Добавлена доля"
'@,
@'
let
    Источник = #"исходник БНС",
    #"Добавлена доля трудоустроенных" = Table.AddColumn(Источник, "Доля трудоустроенных", each if [Выпуск всего]=null or [Выпуск всего]=0 then null else [Трудоустроены]/[Выпуск всего], Percentage.Type),
    #"Добавлена доля продолживших" = Table.AddColumn(#"Добавлена доля трудоустроенных", "Доля продолживших обучение", each if [Выпуск всего]=null or [Выпуск всего]=0 then null else [Продолжили обучение]/[Выпуск всего], Percentage.Type)
in
    #"Добавлена доля продолживших"
'@
)

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$workbook = $null
try {
    for ($i=0; $i -lt $books.Count; $i++) {
        $name = [string]$books[$i][0]
        $path = Join-Path $out "$name.xlsx"
        if (Test-Path -LiteralPath $path) { continue }
        $workbook = $excel.Workbooks.Add(-4167)
        $result = $workbook.Worksheets.Item(1)
        $result.Name = 'итоги'
        $tables = $books[$i][1]
        for ($j=0; $j -lt $tables.Count; $j++) {
            $sheet = $workbook.Worksheets.Add([Type]::Missing,$workbook.Worksheets.Item($workbook.Worksheets.Count))
            Add-SourceTable $workbook $sheet $tables[$j] ($j+1)
        }
        $null = $workbook.Queries.Add('результат',$formulas[$i])
        $connection = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="результат";Extended Properties=""'
        $table = $result.ListObjects.Add(0,$connection,[Type]::Missing,1,$result.Range('A1'))
        $table.Name = 'РезультатБНС'
        $queryTable = $table.QueryTable
        $queryTable.CommandType = 2
        $queryTable.CommandText = 'SELECT * FROM [результат]'
        $null = $queryTable.Refresh($false)
        $count = $table.ListRows.Count
        if ($count -lt 100) { throw "Мало строк: $name ($count)" }
        Set-Layout $result $table
        $workbook.SaveAs($path,51)
        $workbook.Close($false)
        $workbook = $null
        "$name : $count строк"
    }
}
finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
}
