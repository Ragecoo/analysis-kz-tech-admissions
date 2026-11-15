from collections import Counter
from itertools import combinations
from math import ceil
from pathlib import Path

import pandas as pd


CHECK_YEAR = "2025/26"
FIRST_COMPARABLE_YEAR = "2022/23"
MIN_SUPPORT = 0.08
MIN_CONFIDENCE = 0.55
MIN_ANTECEDENT = 10
AREA = "Техническая область"
IN = "Принято всего"
OUT = "Выпуск всего"


def prepare(path):
    data = pd.read_excel(path, sheet_name="итоги", dtype={"Код": str})
    keys = ["Учебный год", "Местность", "Код"]
    if data.duplicated(keys).any():
        raise ValueError("Есть дубликаты специальностей по году и местности")
    # За 2021/22 в выгрузке по используемым техническим кодам выпуск равен нулю
    # по всем строкам. Этот год не используем для правил, основанных на выпуске.
    data = data.loc[data["Учебный год"].between(FIRST_COMPARABLE_YEAR, CHECK_YEAR)]
    complete = data.dropna(subset=[AREA, IN, OUT]).copy()
    if complete[[IN, OUT]].lt(0).any().any():
        raise ValueError("Приём и выпуск должны быть неотрицательными")
    learning = complete.loc[complete["Учебный год"] < CHECK_YEAR]
    if learning.empty or not complete["Учебный год"].eq(CHECK_YEAR).any():
        raise ValueError("Нужны данные до контрольного года и за сам контрольный год")
    cutoffs = {column: ceil(learning[column].median()) for column in (IN, OUT)}
    complete["Приём"] = complete[IN].ge(cutoffs[IN]).map({True: "Приём: высокий", False: "Приём: низкий"})
    complete["Выпуск"] = complete[OUT].ge(cutoffs[OUT]).map({True: "Выпуск: высокий", False: "Выпуск: низкий"})
    complete["Область"] = complete[AREA].map({
        "ИКТ": "Область: ИКТ",
        "Инженерные, обрабатывающие и строительные отрасли": "Область: инженерия, производство и строительство",
    })
    if complete["Область"].isna().any():
        raise ValueError("Найдена неизвестная техническая область")
    complete["Транзакция"] = [
        frozenset(row) for row in complete[["Область", "Приём", "Выпуск"]].itertuples(index=False, name=None)
    ]
    return complete, cutoffs, len(data) - len(complete)


def apriori(transactions):
    """Набор включается в кандидаты, только если все его поднаборы частые."""
    minimum = ceil(MIN_SUPPORT * len(transactions))
    single_counts = Counter(frozenset([item]) for transaction in transactions for item in transaction)
    level = {items: count for items, count in single_counts.items() if count >= minimum}
    frequent = level.copy()
    for size in (2, 3):
        items = sorted(set().union(*level))
        candidates = [
            frozenset(candidate)
            for candidate in combinations(items, size)
            if all(frozenset(subset) in level for subset in combinations(candidate, size - 1))
        ]
        level = {
            candidate: count
            for candidate in candidates
            if (count := sum(candidate <= transaction for transaction in transactions)) >= minimum
        }
        frequent.update(level)
    return frequent


def rule_stats(transactions, premise, conclusion):
    n = len(transactions)
    a = sum(premise <= row for row in transactions)
    b = sum(conclusion <= row for row in transactions)
    both = sum((premise | conclusion) <= row for row in transactions)
    confidence = both / a if a else None
    return {
        "Строк": n,
        "Условие": a,
        "Следствие": b,
        "Совпадений": both,
        "Support": both / n if n else None,
        "Confidence": confidence,
        "Lift": confidence / (b / n) if a and b else None,
    }


def discover(transactions):
    rules = []
    for items, count in apriori(transactions).items():
        if len(items) < 2:
            continue
        for item in sorted(items):
            conclusion = frozenset([item])
            premise = items - conclusion
            stats = rule_stats(transactions, premise, conclusion)
            if stats["Условие"] >= MIN_ANTECEDENT and stats["Confidence"] >= MIN_CONFIDENCE and stats["Lift"] > 1:
                assert stats["Совпадений"] == count
                rules.append((premise, conclusion, stats))
    return sorted(rules, key=lambda row: (-row[2]["Lift"], -row[2]["Support"], tuple(sorted(row[0])), tuple(row[1])))


def main():
    folder = Path(__file__).resolve().parent
    source = folder.parents[1] / "excel" / "03_urban_rural.xlsx"
    data, cutoffs, excluded = prepare(source)
    train = data.loc[data["Учебный год"] < CHECK_YEAR]
    test = data.loc[data["Учебный год"] == CHECK_YEAR]
    rules = discover(train["Транзакция"].tolist())
    if len(rules) < 5:
        raise ValueError(f"Найдено только {len(rules)} правил при заданных порогах")

    master, locality, years = [], [], []
    for number, (premise, conclusion, stats) in enumerate(rules, 1):
        key = {
            "ID правила": f"R{number:02}",
            "Если": " + ".join(sorted(premise)),
            "То": next(iter(conclusion)),
        }
        master.append({**key, **stats})
        for period, period_data in [("Поиск", train), ("Проверка", test)]:
            for place, part in period_data.groupby("Местность", sort=True):
                locality.append({
                    **key, "Период": period, "Местность": place,
                    **rule_stats(part["Транзакция"].tolist(), premise, conclusion),
                })
        for year, part in data.groupby("Учебный год", sort=True):
            years.append({
                **key, "Учебный год": year,
                **rule_stats(part["Транзакция"].tolist(), premise, conclusion),
            })

    settings = pd.DataFrame([
        ("Высокий приём, от человек", cutoffs[IN]),
        ("Высокий выпуск, от человек", cutoffs[OUT]),
        ("Минимальный Support при поиске", MIN_SUPPORT),
        ("Минимальный Confidence при поиске", MIN_CONFIDENCE),
        ("Минимум случаев условия при поиске", MIN_ANTECEDENT),
        ("Строк при поиске", len(train)),
        ("Строк при проверке", len(test)),
        ("Пропущено неполных строк", excluded),
        ("Начало периода поиска", FIRST_COMPARABLE_YEAR),
        ("Контрольный год", CHECK_YEAR),
        ("Ограничение 2021/22", "Выпуск нулевой во всех строках по используемым кодам"),
    ], columns=["Настройка", "Значение"])

    folder = folder / "result"
    folder.mkdir(exist_ok=True)
    for name, table in {
        "rules.csv": pd.DataFrame(master),
        "rules_by_locality.csv": pd.DataFrame(locality),
        "rules_by_year.csv": pd.DataFrame(years),
        "settings.csv": settings,
    }.items():
        table.to_csv(folder / name, index=False, encoding="utf-8-sig")

    print(f"Поиск: {len(train)}, проверка: {len(test)}, исключено строк: {excluded}")
    print(f"Высокий приём от {cutoffs[IN]}, высокий выпуск от {cutoffs[OUT]}")
    print(f"Найдено правил: {len(master)}")
    print(pd.DataFrame(master)[["ID правила", "Если", "То", "Support", "Confidence", "Lift"]].round(3).to_string(index=False))


if __name__ == "__main__":
    main()
