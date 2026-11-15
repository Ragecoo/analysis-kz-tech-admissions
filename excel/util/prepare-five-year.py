"""Prepare source extracts for the six Power Query workbooks."""

import csv
import json
import re
import subprocess
from collections import Counter
from pathlib import Path
from tempfile import TemporaryDirectory

from extract import HERE, ROOT, SOURCE, normalized, number, parse, rows_from


PERIOD = range(2021, 2026)
REGIONS = (
    "Республика Казахстан|Қазақстан Республикасы|Республика Казакстан",
    "Абай", "Акмолинская|Ақмола", "Актюбинская|Ақтөбе",
    "Алматинская|Алматы", "Атырауская|Атырау",
    "Западно-Казахстанская|Батыс Қазақстан", "Жамбылская|Жамбыл",
    "Жетісу|Жетису", "Карагандинская|Қарағанды", "Костанайская|Қостанай",
    "Кызылординская|Қызылорда", "Мангистауская|Маңғыстау", "Павлодарская|Павлодар",
    "Северо-Казахстанская|Северно-Казахстанская|Солтүстік Қазақстан", "Туркестанская|Түркістан|Туркистанская",
    "Ұлытау|Улытау", "Восточно-Казахстанская|Шығыс Қазақстан",
    "город Астана|город Нур-Султан|г. Астана|г.Астана|Астана қаласы|Нұр-Сұлтан қаласы",
    "город Алматы|г. Алматы|г.Алматы|Алматы қаласы",
    "город Шымкент|г. Шымкент|г.Шымкент|Шымкент қаласы",
)
REGION_NAMES = {alias.casefold(): group.split("|")[0] for group in REGIONS for alias in group.split("|")}
LEVEL = {"colleges": "СУЗы", "universities": "Вузы"}
ADMISSION = ("Принято всего", "Принято женщин", "Численность всего", "Численность женщин",
             "Выпуск всего", "Выпуск женщин")
CAPTIONS = ("После 9 класса", "После 11 класса", "После ТиПО")


def year_label(year):
    return f"{year}/{(year + 1) % 100:02}"


def region(value):
    label = normalized(value).replace("_x000D_", " ").strip()
    return REGION_NAMES.get(label.casefold())


def geographic_rows(rows):
    matches = {}
    for index, row in enumerate(rows, 1):
        name = region(row[0]) if row else None
        if name and any(isinstance(number(v), int) for v in row[1:4]):
            if name in matches:
                break
            matches[name] = (index, row)
    return matches


def original(level, year):
    return next((SOURCE / level).glob(f"*_{year}-{year + 1}.*"))


def locator(file, sheet, index, member=""):
    return (file.relative_to(ROOT).as_posix(), member, sheet, index)


def technical(level, code):
    return code.startswith(("06", "07") if level == "colleges" else ("6B06", "6B07"))


def detail_code(level, code):
    return bool(re.fullmatch(r"\d{8}" if level == "colleges" else r"6B\d{3}", code))


def age_group(label):
    if "до 14 лет" in label:
        return "до 14 лет", "до 18 лет"
    match = re.search(r"(\d+(?:\s*[-–]\s*\d+)?)\s*(?:лет|год(?:а)?|и старше)(?:\s+и старше)?$", label)
    if not match:
        return None
    age = int(re.search(r"\d+", match[1]).group())
    bucket = "до 18 лет" if age < 18 else "18–19 лет" if age < 20 else "20–24 года" if age < 25 else "25 лет и старше"
    return match[0], bucket


def prepare():
    profiles, budgets, places, ages, bases, outcomes, regions = ([] for _ in range(7))
    with (HERE / "specialty-source.csv").open(encoding="utf-8-sig", newline="") as file:
        for row in csv.DictReader(file, delimiter=";"):
            level, year, code = row["level"], int(row["year"]), row["code"]
            if year not in PERIOD or not detail_code(level, code):
                continue
            profiles.append((year_label(year), LEVEL[level], code, row["name"], *(
                int(row[k]) if row[k] else None for k in
                ("admissions", "women_admissions", "students", "women_students", "graduates", "women_graduates")
            ), row["file"], row["member"], row["sheet"], int(row["row"])))

    for level, caption in LEVEL.items():
        f = SOURCE / level / ("college_admissions_series.xls" if level == "colleges" else "university_admissions_series.xls")
        sheet = "Показатель" if level == "colleges" else "Прием всего"
        rows = rows_from(f, sheet)
        years = {int(v): i for i, v in enumerate(rows[2]) if isinstance(v, (int, float))}
        for i, row in enumerate(rows[3:], 4):
            if len(row) < 27 or not normalized(row[1]):
                continue
            if not isinstance(row[0], (int, float)) and normalized(row[1]) != "Республика Казахстан":
                continue
            for year in PERIOD:
                regions.append((year_label(year), caption, str(int(row[0])) if isinstance(row[0], (int, float)) else None,
                                region(row[1]) or normalized(row[1]), number(row[years[year]]), *locator(f, sheet, i)))

    with TemporaryDirectory() as directory:
        arch = original("universities", 2021)
        subprocess.run(["tar", "-xf", str(arch), "-C", directory], check=True, capture_output=True)
        university_2021 = next(Path(directory).rglob("*.xlsx"))
        for year in PERIOD:
            for level, caption in LEVEL.items():
                file = original(level, year)
                actual = university_2021 if (level, year) == ("universities", 2021) else file
                member = actual.relative_to(directory).as_posix() if actual != file else ""
                tab = ("3.11" if year == 2021 else "1.11" if year == 2022 else "3.12") if level == "colleges" else ("3" if year == 2022 else "19")
                for row, code, name, *values in parse(rows_from(actual, tab), level)[1]:
                    if detail_code(level, code):
                        budgets.append((year_label(year), caption, code, name, values[0], values[1], *locator(file, tab, row, member)))

                supplement = SOURCE / "supplementary"
                use = (supplement / ("college_students_2022-2023.xls" if level == "colleges" else "university_network_2022-2023.xlsx")) if year == 2022 else actual
                source_file, source_member = (file, member) if use == actual else (use, "")
                age_tab = ("3" if year == 2022 else "7") if level == "colleges" else ("11" if year == 2022 else "9")
                age_rows = rows_from(use, age_tab)
                age_header = next(r for r in age_rows[:15] if "до 14 лет" in str(r).replace("\\n", " "))
                columns = [(i, g) for i, value in enumerate(age_header) if (g := age_group(normalized(value).replace("_x000D_", " ")))]
                for name, (index, row) in geographic_rows(age_rows).items():
                    for i, (label, bucket) in columns:
                        ages.append((year_label(year), caption, name, label, bucket, number(row[i]), number(row[1]),
                                     *locator(source_file, age_tab, index, source_member)))

                if level == "colleges":
                    locality_tabs = ("1.1", "1.2") if year == 2022 else ("3.1", "3.2")
                    for locality, tab in zip(("Город", "Село"), locality_tabs):
                        for row, code, name, *values in parse(rows_from(actual, tab), level)[1]:
                            if detail_code(level, code) and technical(level, code):
                                places.append((year_label(year), locality, code, name, *values,
                                               *locator(file, tab, row)))
                    base_tabs = ("2.11", "2.12", "2.13") if year == 2021 else ("1.11", "1.12", "1.13") if year == 2022 else ("2.12", "2.13", "2.14")
                    for base, tab in zip(CAPTIONS, base_tabs):
                        for name, (i, row) in geographic_rows(rows_from(use, tab)).items():
                            bases.append((year_label(year), caption, name, base, number(row[3]), *locator(source_file, tab, i, source_member)))
                    tab = "8" if year == 2022 else "14"
                    for name, (i, row) in geographic_rows(rows_from(use, tab)).items():
                        outcomes.append((year_label(year), name, *(number(v) for v in row[1:7]), *locator(source_file, tab, i, source_member)))
                else:
                    tab = "6." if year == 2022 else "5"
                    for name, (i, row) in geographic_rows(rows_from(use, tab)).items():
                        columns = (2, 4, 6) if year in (2021, 2022) else (2, 3, 4)
                        for base, column in zip(("После школы", "После ТиПО", "После высшего образования"), columns):
                            bases.append((year_label(year), caption, name, base, number(row[column]), *locator(source_file, tab, i, source_member)))

    origin = ("Файл БНС", "Книга в архиве", "Лист БНС", "Строка БНС")
    profile_key = ("Учебный год", "Уровень", "Код", "Название")
    return [
        ("01 специальности", [(profile_key + ADMISSION + origin, profiles),
                               (profile_key + ("По госзаказу всего", "По госзаказу женщин") + origin, budgets)]),
        ("02 регионы", [(("Учебный год", "Уровень", "КАТО", "Регион", "Принято всего") + origin, regions)]),
        ("03 город и село", [(("Учебный год", "Местность", "Код", "Название") + ADMISSION + origin, places)]),
        ("04 возраст", [(("Учебный год", "Уровень", "Регион", "Возраст в БНС", "Возрастная группа", "Численность", "Контингент") + origin, ages)]),
        ("05 база поступления", [(("Учебный год", "Уровень", "Регион", "База поступления", "Принято всего") + origin, bases)]),
        ("06 выпуск и работа", [(("Учебный год", "Регион", "Выпуск всего", "Продолжили обучение", "Продолжили в вузах", "Продолжили в СУЗах", "Трудоустроены", "Трудоустроены с очной формы") + origin, outcomes)]),
    ]


if __name__ == "__main__":
    import sys
    books = prepare()
    for name, tables in books:
        for header, rows in tables:
            counts = Counter(row[0] for row in rows)
            assert set(counts) == {year_label(y) for y in PERIOD}, (name, counts)
            assert all(len(row) == len(header) for row in rows)
            print(name, dict(sorted(counts.items())))
    Path(sys.argv[1]).write_text(json.dumps(books, ensure_ascii=False), encoding="utf-8")
