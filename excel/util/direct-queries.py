"""Generate per-publication Power Query imports for the five-year workbooks."""

import importlib.util
import json
import re
import subprocess
import sys
from collections import defaultdict
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from extract import HERE, ROOT, SOURCE, normalized, rows_from


spec = importlib.util.spec_from_file_location("prepare", HERE / "prepare-five-year.py")
prepare = importlib.util.module_from_spec(spec)
spec.loader.exec_module(prepare)
BOOKS = ROOT / "excel/bns-5-let"
SUPPLEMENT = SOURCE / "supplementary"


def m(value):
    if value is None:
        return "null"
    if isinstance(value, (int, float)):
        return str(value)
    return '"' + str(value).replace('"', '""').replace("#(", "#(#)(") + '"'


def ml(values):
    return "{" + ", ".join(map(m, values)) + "}"


NUMBER = """(Значение as any) as nullable number =>
    if Значение = null then null
    else if Value.Is(Значение, type number) then Int64.From(Значение)
    else let Текст = Text.Trim(Text.From(Значение)) in
        if List.Contains({"-", "–", "—"}, Текст) then 0
        else if List.Contains({"", "x", "х", "..." , "…"}, Text.Lower(Текст)) then null
        else Int64.From(Текст)"""


def navigation(path, tab, rows):
    legacy = path.suffix.lower() == ".xls"
    # ACE uses # for periods in sheet names and only provides the Name key.
    key = f"Name={m(tab.replace('.', '#'))}" if legacy else f'Item={m(tab)}, Kind="Sheet"'
    candidates = [
        (i, normalized(value)) for i, row in enumerate(rows[:14], 1) for value in row[:7]
        if isinstance(value, str) and len(normalized(value)) >= 12
    ]
    anchor_row, anchor = next((i, value) for i, value in candidates
                              if sum(item == value for _, item in candidates) == 1)
    return (
        f"    Источник = Excel.Workbook(File.Contents(#\"Папка проекта\" & {m('/' + path.relative_to(ROOT).as_posix())}), null, true),\n"
        f"    Навигация = Источник{{[{key}]}}[Data],\n"
        '    #"Нормализован заголовок" = (Значение as any) as text => if Value.Is(Значение, type text) then Text.Combine(List.Select(Text.SplitAny(Значение, " #(tab)#(lf)#(cr)" & Character.FromNumber(160)), each _ <> ""), " ") else "",\n'
        f'    #"Найдена строка заголовка" = List.PositionOf(List.Transform(Table.ToRows(Table.FirstN(Навигация, 18)), each List.Contains(List.Transform(List.FirstN(_, 7), #"Нормализован заголовок"), {m(anchor)})), true),\n'
        f'    #"Первая строка БНС" = if #"Найдена строка заголовка" < 0 then error "Заголовок исходного листа изменился" else {anchor_row} - #"Найдена строка заголовка",\n'
        '    #"Добавлен номер строки БНС" = Table.AddIndexColumn(Навигация, "Строка БНС", #"Первая строка БНС", 1, Int64.Type),\n'
    )


def regional(path, tab, records, header):
    year_start = 2021
    years = [f"Column{i}" for i in range(24, 29)]
    mapping = {record[-1]: (record[2], record[3]) for record in records}
    map_rows = ", ".join(ml((i, code, name)) for i, (code, name) in mapping.items())
    publish = records[0]
    return f"""let
{navigation(path, tab, rows_from(path, tab))}    Территории = #table({{"Строка БНС","КАТО","Регион"}}, {{{map_rows}}}),
    #"Оставлены регионы и пять лет" = Table.SelectColumns(#"Добавлен номер строки БНС", List.Combine({{{{"Строка БНС"}}, {ml(years)}}})),
    #"Присоединены регионы" = Table.NestedJoin(#"Оставлены регионы и пять лет",{{"Строка БНС"}},Территории,{{"Строка БНС"}},"Территория",JoinKind.Inner),
    #"Раскрыты регионы" = Table.ExpandTableColumn(#"Присоединены регионы","Территория",{{"КАТО","Регион"}}),
    #"Сохранены пропуски" = Table.ReplaceValue(#"Раскрыты регионы",null,"...",Replacer.ReplaceValue,{ml(years)}),
    #"Годы перенесены в строки" = Table.Unpivot(#"Сохранены пропуски",{ml(years)},"Столбец года","Принято всего"),
    #"Добавлен учебный год" = Table.AddColumn(#"Годы перенесены в строки","Учебный год", each let год = {year_start}+Number.FromText(Text.AfterDelimiter([Столбец года],"Column"))-24 in Text.From(год)&"/"&Text.End(Text.From(год+1),2),type text),
    #"Приём приведён к числам" = Table.TransformColumns(#"Добавлен учебный год",{{{{"Принято всего",#"Число БНС",Int64.Type}}}}),
    #"Добавлен уровень" = Table.AddColumn(#"Приём приведён к числам","Уровень", each {m(publish[1])},type text),
    #"Добавлен файл" = Table.AddColumn(#"Добавлен уровень","Файл БНС", each {m(publish[-4])},type text),
    #"Добавлена книга" = Table.AddColumn(#"Добавлен файл","Книга в архиве", each null,type text),
    #"Добавлен лист" = Table.AddColumn(#"Добавлена книга","Лист БНС", each {m(tab)},type text),
    #"Упорядочены столбцы" = Table.SelectColumns(#"Добавлен лист",{ml(header)})
in
    #"Упорядочены столбцы" """


def positions(book, filename, tab, rows, header, records):
    first_numeric = {"01": 4, "03": 4, "04": 5, "05": 4, "06": 2}[book]
    metric_names = header[first_numeric:-4]
    if book in {"01", "03"}:
        first = next(
            i for r in rows[:16] for i, v in enumerate(r[:5])
            if i > 0 and re.search(r"при[её]м (учащихся|обучающихся|студентов)|принято в отчетном году", normalized(v), re.I)
        )
        return first_numeric, metric_names, [tuple(range(first, first + len(metric_names)))] * len(records)
    if book == "06":
        return first_numeric, metric_names, [(1, 2, 3, 4, 5, 6)] * len(records)
    if book == "05":
        categories = {"После школы": 0, "После ТиПО": 1, "После высшего образования": 2}
        fields = []
        for row in records:
            early = row[0] in {"2021/22", "2022/23"}
            fields.append((3,) if row[1] == "СУЗы" else (((2, 4, 6) if early else (2, 3, 4))[categories[row[3]]],))
        return first_numeric, metric_names, fields

    age_header = next(r for r in rows[:15] if "до 14 лет" in str(r))
    age_columns = {
        g[0]: i for i, value in enumerate(age_header)
        if (g := prepare.age_group(normalized(value).replace("_x000D_", " ")))
    }
    return first_numeric, metric_names, [(age_columns[row[3]], 1) for row in records]


def detailed(book, path, tab, header, records):
    rows = rows_from(path, tab)
    first, metrics, fields = positions(book, path, tab, rows, header, records)
    keys = header[:first]
    fixed = len(set(fields)) == 1
    marker = () if fixed else ("Столбец БНС",)
    pivot = book == "04" and not fixed
    metadata = []
    for row, columns in zip(records, fields):
        metadata.append(ml((*row[:first], row[-1], *((f"Column{columns[0]+1}",) if marker else ()))))
    specs = ",\n        ".join(metadata)
    source_columns = [f"Column{i + 1}" for i in sorted({i for cols in fields for i in cols})]
    joins = ml(("Строка БНС", *marker))
    if fixed:
        mapping = ", ".join(ml((f"Column{i + 1}", name)) for i, name in zip(fields[0], metrics))
        raw_transform = f'Table.RenameColumns(#"Оставлены столбцы показателей", {{{mapping}}})'
    else:
        retained = [col for col in source_columns if not pivot or col != "Column2"]
        raw_transform = (
            'Table.Unpivot(Table.ReplaceValue('
            '#"Оставлены столбцы показателей",null,"...",Replacer.ReplaceValue,'
            f'{ml(retained)}),{ml(retained)},"Столбец БНС",{m(metrics[0])})'
        )
        if pivot:
            raw_transform = f'Table.RenameColumns({raw_transform}, {{{ml(("Column2", metrics[1]))}}})'
    types = ", ".join("{" + m(name) + ', #"Число БНС", Int64.Type}' for name in metrics)
    source = records[0]
    return f"""let
{navigation(path, tab, rows)}    Разметка = #table({ml((*keys, "Строка БНС", *marker))}, {{
        {specs}
    }}),
    #"Оставлены столбцы показателей" = Table.SelectColumns(#"Добавлен номер строки БНС",{ml(("Строка БНС", *source_columns))}),
    #"Подготовлены показатели листа" = {raw_transform},
    #"Присоединены строки листа" = Table.NestedJoin(Разметка,{joins},Table.Buffer(#"Подготовлены показатели листа"),{joins},"Строка источника",JoinKind.LeftOuter),
    #"Раскрыты показатели" = Table.ExpandTableColumn(#"Присоединены строки листа","Строка источника",{ml(metrics)}),
    #"Измененный тип" = Table.TransformColumns(#"Раскрыты показатели", {{{types}}}),
    #"Добавлен файл" = Table.AddColumn(#"Измененный тип","Файл БНС", each {m(source[-4])},type text),
    #"Добавлена книга" = Table.AddColumn(#"Добавлен файл","Книга в архиве", each {m(source[-3])},type text),
    #"Добавлен лист" = Table.AddColumn(#"Добавлена книга","Лист БНС", each {m(tab)},type text),
    #"Упорядочены столбцы" = Table.SelectColumns(#"Добавлен лист",{ml(header)})
in
    #"Упорядочены столбцы" """


def build():
    archive = SOURCE / "universities/university_2021-2022.rar"
    extracted = SUPPLEMENT / "university_2021-2022_unpacked.xlsx"
    with TemporaryDirectory() as directory:
        subprocess.run(["tar", "-xf", str(archive), "-C", directory], check=True, capture_output=True)
        contents = next(Path(directory).rglob("*.xlsx")).read_bytes()
        if extracted.exists() and extracted.read_bytes() != contents:
            raise ValueError("Existing workbook differs from the BNS archive")
        extracted.write_bytes(contents)
    output = {}
    for filename in sorted(BOOKS.glob("*.xlsx")):
        book = filename.name[:2]
        workbook = load_workbook(filename, read_only=True, data_only=True)
        queries = {"Папка проекта": m(ROOT.as_posix()), "Число БНС": NUMBER}
        expected = {}
        for sheet in workbook:
            if not sheet.title.startswith("исходные"):
                continue
            target = "госзаказ БНС" if "госзаказ" in sheet.title else "исходник БНС"
            iterator = sheet.values
            header = tuple(next(iterator))
            groups = defaultdict(list)
            for record in iterator:
                groups[(record[-4], record[-3], record[-2])].append(record)
            names = []
            for (original, archive_member, tab), records in groups.items():
                path = extracted if archive_member else ROOT / original
                name = f"БНС {target.split()[0]} {len(queries)-1:02} {path.stem} лист {tab.replace('.', '-')}"
                if name in queries:
                    raise ValueError(name)
                queries[name] = regional(path, tab, records, header) if book == "02" else detailed(book, path, tab, header, records)
                names.append('#"' + name + '"')
            queries[target] = (
                'let\n    Источник = Table.Buffer(Table.Combine({' + ', '.join(names) + '})),\n'
                '    #"Учебные годы" = Table.SelectRows(Источник, each List.Contains({"2021/22","2022/23","2023/24","2024/25","2025/26"}, [Учебный год]))\n'
                'in\n    #"Учебные годы"'
            )
            expected[sheet.title] = target
        workbook.close()
        output[filename.name] = {"queries": queries, "oldSheets": list(expected)}
    return output


if __name__ == "__main__":
    Path(sys.argv[1]).write_text(json.dumps(build(), ensure_ascii=False), encoding="utf-8")
