"""Extract national admission tables from the local BNS annual bulletins."""

import csv
import re
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory

import xlrd
from openpyxl import load_workbook


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
SOURCE = ROOT / "sources/data/admissions"
YEARS = range(2011, 2026)
COLLEGE_ARCHIVE_BOOKS = (
    "t3_2_0", "t3_1_00", "t3_2_0", "t3_2_0", "t3_2_0",
    "t3_1_00", "3.0", "3.0", "Том1",
)
UNIVERSITY_ARCHIVE_BOOKS = (
    "t04_00_0__000_00", "t04_00_0__000_00", "t04_00_0__000_00",
    "t04_00_0__000_00", "t04_00_0__000 _00", "1.", "1.", "1", "2_том",
)
CODE_PATTERN = re.compile(
    r"^(6B\d{2,3}|5B\d{6}|\d{8}|\d{7}|\d{6}|\d{2})(?=\s|$)",
    re.IGNORECASE,
)


def normalized(value):
    return " ".join(str(value or "").split())


def rows_from(path, sheet):
    if path.suffix.lower() == ".xlsx":
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            return list(book[sheet].values) if sheet in book.sheetnames else []
        finally:
            book.close()
    book = xlrd.open_workbook(file_contents=path.read_bytes())
    if sheet not in book.sheet_names():
        return []
    worksheet = book.sheet_by_name(sheet)
    return [worksheet.row_values(index) for index in range(worksheet.nrows)]


def number(value):
    if isinstance(value, (int, float)) and value == int(value):
        return int(value)
    if normalized(value) in {"-", "–", "—"}:
        return 0
    return ""


def official_totals(level):
    file = SOURCE / level / ("college_admissions_series.xls" if level == "colleges" else "university_admissions_series.xls")
    sheet = "Показатель" if level == "colleges" else "Прием всего"
    rows = rows_from(file, sheet)
    return dict(zip(map(int, rows[2][2:28]), map(int, rows[3][2:28])))


def sheet_and_book(level, year):
    if level == "colleges":
        sheet = "Лист1" if year <= 2016 else "1" if year == 2022 else "3"
        books = COLLEGE_ARCHIVE_BOOKS
    else:
        sheet = (
            "Лист1" if year <= 2015 else "3" if year <= 2018 else
            "2" if year == 2019 else "1" if year == 2022 else "17"
        )
        books = UNIVERSITY_ARCHIVE_BOOKS
    return sheet, books[year - 2011] if year < 2020 else None


def parse(rows, level):
    for index, row in enumerate(rows[:16]):
        heading = next(
            (column for column, value in enumerate(row[:5])
             if re.search(r"при[её]м (учащихся|обучающихся|студентов)|принято в отчетном году",
                          normalized(value), re.I)),
            None,
        )
        if heading is not None and heading > 0:
            break
    else:
        return None, []

    header = " ".join(normalized(value) for row in rows[:index + 1] for value in row)
    full_profile = "Численность" in header and "Выпуск" in header
    total = None
    records = []
    for line, row in enumerate(rows[index + 1:], index + 2):
        if len(row) <= heading:
            continue
        if total is None and any(
            re.fullmatch(r"(?:0+\s+)?(?:Всего|Барлығы)\s*[*¹⁾)]*", normalized(value), re.I)
            for value in row
        ):
            value = number(row[heading])
            if isinstance(value, int):
                total = value
        code_cell = row[heading - 1]
        if heading == 2 and isinstance(code_cell, (int, float)) and code_cell == int(code_cell):
            code_cell = str(int(code_cell)).zfill(7 if level == "colleges" else 6)
        labels = [normalized(code_cell), normalized(row[0])]
        if len(row) > heading + (6 if full_profile else 2):
            labels.insert(0, normalized(row[-1]))
        match = next(((value, CODE_PATTERN.match(value.replace("В", "B", 1)))
                      for value in labels if CODE_PATTERN.match(value.replace("В", "B", 1))), None)
        if match is None:
            continue
        label, code_match = match
        code = code_match[1].upper()
        if not int(re.sub(r"\D", "", code)):
            continue
        name = label[len(code):].strip()
        if not name:
            name = next((normalized(v) for v in reversed(row[heading + (6 if full_profile else 2):])
                         if isinstance(v, str) and re.search(r"[А-Яа-я]{3}", v)), normalized(row[0]))
        values = [number(v) for v in row[heading: heading + (6 if full_profile else 2)]]
        if not any(isinstance(v, int) for v in values):
            continue
        records.append((line, code, name, *(values + [""] * (6 - len(values)))))
    return total, records


def annual_table(level, year, expected):
    file = next((SOURCE / level).glob(f"*_{year}-{year + 1}.*"))
    sheet, stem = sheet_and_book(level, year)
    with TemporaryDirectory() as directory:
        if file.suffix.lower() in {".rar", ".zip"}:
            subprocess.run(["tar", "-xf", str(file), "-C", directory], capture_output=True)
            candidates = sorted(
                path for path in Path(directory).rglob("*")
                if path.suffix.lower() in {".xls", ".xlsx"}
                and (stem is None or re.sub(r"\s", "", path.stem) == re.sub(r"\s", "", stem))
            )
        else:
            candidates = [file]
        alternatives = []
        for candidate in candidates:
            rows = rows_from(candidate, sheet)
            total, records = parse(rows, level)
            if records:
                member = candidate.relative_to(directory).as_posix() if candidate != file else ""
                found = file.relative_to(ROOT).as_posix(), member, sheet, total, records
                if total == expected:
                    return found
                alternatives.append(found)
        if len(alternatives) == 1:
            return alternatives[0]
        raise ValueError(f"No unambiguous national table for {level} {year} ({expected}): "
                         f"{[(item[1], item[3]) for item in alternatives]}")


def extract():
    detail, registry = [], []
    for level in ("colleges", "universities"):
        totals = official_totals(level)
        for year in YEARS:
            file, member, sheet, total, records = annual_table(level, year, totals[year])
            detail.extend((level, year, file, member, sheet, *record) for record in records)
            registry.append((level, year, file, member, sheet, totals[year], total, len(records)))
    return detail, registry


def save(name, columns, records):
    with (HERE / name).open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, delimiter=";")
        writer.writerow(columns.split(";"))
        writer.writerows(records)


def selection(level, year, code):
    if level == "colleges" and year >= 2021:
        if code in {"06", "07"}:
            return "published_area"
        if len(code) == 8 and code[:2] in {"06", "07"}:
            return "modern_detail"
        return ""
    if level == "universities" and year >= 2019:
        if code in {"6B06", "6B07"}:
            return "published_area"
        if len(code) == 5 and code[:4] in {"6B06", "6B07"}:
            return "modern_detail"
        return ""
    if level == "colleges":
        if len(code) != 7 or code.endswith("00000"):
            return ""
        if code in {"1122000", "1202000", "1203000", "1204000", "1205000", "1206000", "1226000",
                    "1227000", "1231000", "1417000", "1501000", "1510000", "1511000", "1518000"}:
            return "legacy_review"
        if code[:2] in {"06", "07", "08", "09", "10", "11", "12", "13", "14"}:
            return "legacy_included"
        return ""
    if code in {"050420", "5B042000", "051002", "5B100200"}:
        return "legacy_included"
    if code in {"050700", "5B070000"}:
        return ""
    if code.startswith(("0507", "5B07")):
        if code in {"050701", "050705", "050706", "050711", "050722", "050723", "050731",
                    "5B070100", "5B070500", "5B070600", "5B071100", "5B072200", "5B072300", "5B073100"}:
            return "legacy_review"
        return "legacy_included"
    return ""


def technical_series(detail, registry):
    selected = []
    for record in detail:
        level, year, _, _, _, _, code, *_ = record
        rule = selection(level, year, code)
        if rule:
            selected.append((*record, rule))
    annual = []
    for level, year, file, member, sheet, official, reported, _ in registry:
        subset = [record for record in selected if record[:2] == (level, year)]
        rule = "published_area" if any(record[-1] == "published_area" for record in subset) else "legacy_included"
        values = [record for record in subset if record[-1] == rule]
        reviews = [record for record in subset if record[-1] == "legacy_review"]
        if not values or (rule == "published_area" and len(values) != 2):
            raise ValueError(f"Unexpected technical scope for {level} {year}")
        codes = [record[6] for record in values if record[8]]
        if len(codes) != len(set(codes)):
            raise ValueError(f"Duplicate codes with positive admissions for {level} {year}")
        annual.append((
            level, year, f"{year}/{(year + 1) % 100:02d}", rule,
            sum(record[8] for record in values if isinstance(record[8], int)),
            sum(record[9] for record in values if isinstance(record[9], int)),
            official, reported, len(values),
            sum(record[8] for record in reviews if isinstance(record[8], int)),
            sum(record[8] == "" for record in values), file, member, sheet,
        ))
    return selected, annual


if __name__ == "__main__":
    detail, registry = extract()
    detail_columns = "level;year;file;member;sheet;row;code;name;admissions;women_admissions;students;women_students;graduates;women_graduates"
    selected, annual = technical_series(detail, registry)
    save("specialty-source.csv", detail_columns, detail)
    save("source-tables.csv", "level;year;file;member;sheet;official_all_admissions;table_all_admissions;source_rows", registry)
    save("technical-codes.csv", f"{detail_columns};selection", selected)
    save("admissions-by-year.csv", "level;year;academic_year;status;technical_admissions;technical_women_admissions;official_all_admissions;table_all_admissions;rows_in_total;review_admissions;unavailable_admissions_rows;file;member;sheet", annual)
