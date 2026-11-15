from pathlib import Path

import pandas as pd
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score


AGE_GROUPS = ["до 18 лет", "18–19 лет", "20–24 года", "25 лет и старше"]
AGE_COLUMNS = [f"Доля {group}" for group in AGE_GROUPS]
KEYS = ["Учебный год", "Регион"]
REFERENCE_YEAR = "2025/26"
FIRST_COMPARABLE_YEAR = "2022/23"


def age_profiles(project):
    ages = pd.read_excel(project / "excel" / "04_age_groups.xlsx", sheet_name="итоги")
    rows = ages.loc[
        ages["Уровень"].eq("СУЗы")
        & ages["Регион"].ne("Республика Казахстан")
        & ages["Учебный год"].between(FIRST_COMPARABLE_YEAR, REFERENCE_YEAR)
    ].copy()
    if set(rows["Возрастная группа"]) != set(AGE_GROUPS):
        raise ValueError("Изменились возрастные группы в источнике")

    totals = rows.groupby(KEYS)["Контингент"]
    if totals.nunique(dropna=False).ne(1).any():
        raise ValueError("Для одного региона и года получены разные итоги")
    counts = rows.pivot(index=KEYS, columns="Возрастная группа", values="Численность")
    counts = counts.reindex(columns=AGE_GROUPS)
    population = totals.first().reindex(counts.index)
    complete = counts.notna().all(axis=1) & population.gt(0)
    if counts.loc[complete].lt(0).any().any():
        raise ValueError("В профилях обнаружена отрицательная численность")
    if not counts.loc[complete].sum(axis=1).eq(population.loc[complete]).all():
        raise ValueError("Сумма возрастных групп не совпадает с контингентом")

    shares = counts.div(population, axis=0)
    shares.columns = AGE_COLUMNS
    shares["Контингент"] = population
    shares["Полный профиль"] = complete
    shares = shares.reset_index()

    region_map = pd.read_csv(project / "excel" / "region_coordinates.csv", dtype={"КАТО": str})
    region_map["Регион"] = region_map["Регион"].str.rstrip("*").str.strip()
    shares = shares.merge(region_map, on="Регион", how="left", validate="many_to_one")
    if shares["КАТО"].isna().any():
        raise ValueError("У части регионов нет кода КАТО")
    return shares


def main():
    analysis_dir = Path(__file__).resolve().parent
    project = analysis_dir.parents[1]
    regions = age_profiles(project)
    reference = regions.loc[
        regions["Учебный год"].eq(REFERENCE_YEAR) & regions["Полный профиль"]
    ].sort_values("КАТО")
    if len(reference) <= 4:
        raise ValueError("Слишком мало полных профилей для сравнения K-Means")

    # Все признаки — доли в одной шкале. Сравниваем различия в процентных пунктах.
    X = reference[AGE_COLUMNS]
    scores = []
    models = {}
    for k in (2, 3, 4):
        model = KMeans(n_clusters=k, n_init=30, random_state=42).fit(X)
        models[k] = model
        counts = pd.Series(model.labels_).value_counts()
        scores.append({
            "Учебный год": REFERENCE_YEAR,
            "Число кластеров": k,
            "Silhouette Score": silhouette_score(X, model.labels_),
            "Регионов": len(X),
            "Минимальный размер кластера": int(counts.min()),
            "Максимальный размер кластера": int(counts.max()),
        })
    scores = pd.DataFrame(scores)
    k = int(scores.loc[scores["Silhouette Score"].idxmax(), "Число кластеров"])
    scores["Выбрана модель"] = scores["Число кластеров"].eq(k)
    model = models[k]

    # Кластер 1 имеет наибольшую среднюю долю учащихся младше 18 лет.
    order = sorted(range(k), key=lambda label: -model.cluster_centers_[label, 0])
    labels = {label: i for i, label in enumerate(order, start=1)}
    complete = regions["Полный профиль"]
    regions["Кластер"] = pd.Series(pd.NA, index=regions.index, dtype="Int64")
    regions.loc[complete, "Кластер"] = [
        labels[label] for label in model.predict(regions.loc[complete, AGE_COLUMNS])
    ]
    regions["Статус"] = regions["Полный профиль"].map({
        True: "Классифицирован",
        False: "Пропуск в возрастных данных",
    })

    # Во все годы используем центры 2025/26, чтобы номера кластеров были сопоставимы.
    latest = regions.loc[
        regions["Учебный год"].eq(REFERENCE_YEAR) & complete,
        ["КАТО", "Кластер"],
    ].rename(columns={"Кластер": "Кластер в 2025/26"})
    regions = regions.merge(latest, on="КАТО", how="left", validate="many_to_one")
    regions["Тот же кластер, что в 2025/26"] = (
        regions["Кластер"].eq(regions["Кластер в 2025/26"])
    )

    groups = regions.loc[regions["Полный профиль"]].groupby(["Учебный год", "Кластер"])
    profiles = groups[AGE_COLUMNS].mean().join(groups.size().rename("Регионов")).reset_index()
    stability = (
        regions.loc[regions["Учебный год"].ne(REFERENCE_YEAR)]
        .groupby("Учебный год")["Тот же кластер, что в 2025/26"]
        .agg(["count", "sum"])
        .rename(columns={"count": "Сопоставлено регионов", "sum": "В том же кластере"})
        .reset_index()
    )
    stability["Доля в том же кластере"] = (
        stability["В том же кластере"] / stability["Сопоставлено регионов"]
    )

    result_dir = analysis_dir / "result"
    result_dir.mkdir(exist_ok=True)
    tables = {
        "model_scores.csv": scores,
        "region_clusters.csv": regions[[
            "Учебный год", "КАТО", "Регион", "Широта", "Долгота", "Кластер",
            "Контингент", *AGE_COLUMNS, "Статус", "Кластер в 2025/26",
            "Тот же кластер, что в 2025/26",
        ]].sort_values(["Учебный год", "КАТО"]),
        "cluster_profiles.csv": profiles,
        "stability.csv": stability,
    }
    for filename, table in tables.items():
        table.to_csv(result_dir / filename, index=False, encoding="utf-8-sig")

    print(scores.to_string(index=False))
    print("\nПрофили 2025/26 (доли в процентах)")
    last_profiles = profiles.loc[profiles["Учебный год"].eq(REFERENCE_YEAR)].copy()
    last_profiles[AGE_COLUMNS] = (last_profiles[AGE_COLUMNS] * 100).round(1)
    print(last_profiles.to_string(index=False))
    print("\nПереходы между кластерами")
    print(stability.to_string(index=False))
    excluded = regions.loc[regions["Учебный год"].eq(REFERENCE_YEAR) & ~regions["Полный профиль"], "Регион"]
    if not excluded.empty:
        print("\nПропуск возраста в 2025/26:", ", ".join(excluded))


if __name__ == "__main__":
    main()
