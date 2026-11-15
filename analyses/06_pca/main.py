from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler


YEAR = "2025/26"
PREVIOUS_YEAR = "2024/25"
FEATURES = [
    "Принято всего",
    "Численность всего",
    "Выпуск всего",
    "Изменение приёма",
    "Доля госзаказа",
]


def prepare(data, level):
    technical = ["ИКТ", "Инженерные, обрабатывающие и строительные отрасли"]
    profiles = data.loc[
        data["Уровень"].eq(level) & data["Техническая область"].isin(technical)
    ]
    current = profiles.loc[profiles["Учебный год"].eq(YEAR)].copy()
    previous = profiles.loc[
        profiles["Учебный год"].eq(PREVIOUS_YEAR), ["Код", "Принято всего"]
    ].rename(columns={"Принято всего": "Прошлый приём"})
    current = current.merge(previous, on="Код", how="left", validate="one_to_one")
    current["Изменение приёма"] = current["Принято всего"].div(
        current["Прошлый приём"].where(current["Прошлый приём"].gt(0))
    ) - 1
    if not current["Доля госзаказа"].dropna().between(0, 1).all():
        raise ValueError("В источнике есть доля госзаказа вне диапазона от 0 до 1")
    complete = current[FEATURES].notna().all(axis=1)
    detail = ["Код", "Название", "Учебный год", "Техническая область", "Прошлый приём"]
    excluded = current.loc[~complete, detail].copy()
    missing = current.loc[~complete, FEATURES].isna().to_numpy()
    excluded["Недоступные признаки"] = [
        ", ".join(feature for feature, absent in zip(FEATURES, row, strict=True) if absent)
        for row in missing
    ]
    ready = current.loc[complete, [*detail, *FEATURES]].sort_values(
        ["Техническая область", "Код"]
    )
    return ready, excluded


def fit_pca(data):
    if len(data) < len(FEATURES) or data[FEATURES].nunique().le(1).any():
        raise ValueError("Недостаточно наблюдений или есть постоянный признак")
    standardized = StandardScaler().fit_transform(data[FEATURES])
    pca = PCA().fit(standardized)
    scores = pca.transform(standardized)
    weights = pca.components_.copy()

    # Знак компоненты произволен. Ориентируем PC1 по приёму, PC2 по доле госзаказа.
    for component, feature in [(0, 0), (1, 4)]:
        if weights[component, feature] < 0:
            weights[component] *= -1
            scores[:, component] *= -1
    variance = pd.DataFrame({
        "Компонента": [f"PC{i}" for i in range(1, 6)],
        "Объяснённая дисперсия": pca.explained_variance_ratio_,
        "Накопленная дисперсия": pca.explained_variance_ratio_.cumsum(),
        "Собственное значение": pca.explained_variance_,
    })
    corr = np.corrcoef(standardized.T, scores[:, :2].T)[:len(FEATURES), len(FEATURES):]
    loadings = pd.DataFrame([
        {
            "Компонента": f"PC{i + 1}",
            "Признак": name,
            "Вес": weights[i, j],
            "Нагрузка (корреляция)": corr[j, i],
            "Вклад в компоненту": weights[i, j] ** 2,
        }
        for i in range(2) for j, name in enumerate(FEATURES)
    ])
    return variance, loadings, data.assign(PC1=scores[:, 0], PC2=scores[:, 1])


def sensitivity_table(data, variance, loadings):
    sensitivity = []
    for minimum, subset, explained, weights in [
        (1, data, variance, loadings),
        (10, data.loc[data["Прошлый приём"].ge(10)], None, None),
    ]:
        if explained is None:
            explained, weights, _ = fit_pca(subset)
        sensitivity.append({
            "Минимум прошлого приёма": minimum,
            "Наблюдений": len(subset),
            "Дисперсия PC1": explained.loc[0, "Объяснённая дисперсия"],
            "Дисперсия PC2": explained.loc[1, "Объяснённая дисперсия"],
            "Вклад масштаба в PC1": weights.loc[
                weights["Компонента"].eq("PC1") & weights["Признак"].isin(FEATURES[:3]),
                "Вклад в компоненту",
            ].sum(),
            "Вклад изменения приёма в PC2": weights.loc[
                weights["Компонента"].eq("PC2") & weights["Признак"].eq("Изменение приёма"),
                "Вклад в компоненту",
            ].item(),
            "Вклад госзаказа в PC2": weights.loc[
                weights["Компонента"].eq("PC2") & weights["Признак"].eq("Доля госзаказа"),
                "Вклад в компоненту",
            ].item(),
        })

    return pd.DataFrame(sensitivity)


def main():
    folder = Path(__file__).resolve().parent
    source = folder.parents[1] / "excel" / "01_specialties.xlsx"
    data = pd.read_excel(source, sheet_name="итоги", dtype={"Код": str})
    output = {name: [] for name in (
        "variance.csv", "loadings.csv", "specialty_scores.csv", "excluded.csv", "sensitivity.csv"
    )}
    for level in ("СУЗы", "Вузы"):
        ready, excluded = prepare(data, level)
        variance, loadings, scores = fit_pca(ready)
        tables = (variance, loadings, scores, excluded, sensitivity_table(ready, variance, loadings))
        for name, table in zip(output, tables, strict=True):
            output[name].append(table.assign(Уровень=level))

        print(f"{YEAR}, {level}: наблюдений {len(scores)}, исключено {len(excluded)}.")
        print(variance.round(3).to_string(index=False))
        print(loadings.round(3).to_string(index=False))

    result = folder / "result"
    result.mkdir(exist_ok=True)
    for name, tables in output.items():
        pd.concat(tables, ignore_index=True).to_csv(result / name, index=False, encoding="utf-8-sig")


if __name__ == "__main__":
    main()
