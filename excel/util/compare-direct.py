"""Compare refreshed workbook outputs with the verified five-year extracts."""

from collections import Counter
from pathlib import Path
import sys

from openpyxl import load_workbook


def sheet(filename, title):
    with open(filename, "rb") as file:
        book = load_workbook(file, read_only=True, data_only=True)
        result = list(book[title].values)
        book.close()
        return result


def unused(value):
    return value is None or value == ""


def rounded(value):
    return None if unused(value) else round(value, 12) if isinstance(value, float) else value


for candidate in sorted(Path(sys.argv[1]).glob("*.xlsx")):
    original = Path("excel/bns-5-let") / candidate.name
    old, new = sheet(original, "итоги"), sheet(candidate, "итоги")
    if old[0] != new[0]:
        raise ValueError(f"Columns mismatch in {candidate.name}")
    a = Counter(tuple(map(rounded, row)) for row in old[1:])
    b = Counter(tuple(map(rounded, row)) for row in new[1:])
    if a != b:
        print(candidate.name, "mismatch", sum((a-b).values()), sum((b-a).values()), flush=True)
        for row in list((a-b))[:3]:
            print("expected", row)
        for row in list((b-a))[:3]:
            print("actual  ", row)
    else:
        print(candidate.name, "OK", sum(a.values()), flush=True)
