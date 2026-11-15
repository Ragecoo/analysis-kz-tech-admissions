from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.dummy import DummyClassifier
from sklearn.metrics import (
    accuracy_score,
    confusion_matrix,
    f1_score,
    precision_recall_fscore_support,
)
from sklearn.tree import DecisionTreeClassifier


CLASSES = ["Спад", "Стабильно", "Рост"]
FEATURES = ["Принято всего", "Численность всего", "Выпуск всего", "Инженерная область"]
TARGET = "Фактический класс"
TEST_YEAR = 2025
MIN_CHANGE = 10
MIN_PERCENT = 15
BASELINE = "Наиболее частый класс"
TREE = "Дерево решений"


def make_pairs(source):
    data = pd.read_excel(source, sheet_name="итоги", dtype={"Код": str})
    technical = ["ИКТ", "Инженерные, обрабатывающие и строительные отрасли"]
    data = data.loc[
        data["Уровень"].eq("СУЗы") & data["Техническая область"].isin(technical)
    ].copy()
    data["Год прогноза"] = data["Учебный год"].str[:4].astype(int)
    actual = data[["Код", "Год прогноза", "Принято всего"]].rename(
        columns={"Принято всего": "Приём следующего года"}
    )
    data["Год прогноза"] += 1
    data["Инженерная область"] = data["Техническая область"].eq(technical[1]).astype(int)
    pairs = data.merge(
        actual, on=["Код", "Год прогноза"], how="inner", validate="one_to_one"
    ).dropna(subset=[*FEATURES, "Приём следующего года"])
    pairs["Изменение, чел."] = pairs["Приём следующего года"] - pairs["Принято всего"]
    # Крупный набор должен измениться минимум на 15%, малый — минимум на 10 человек.
    pairs["Порог, чел."] = np.maximum(
        MIN_CHANGE, np.ceil(pairs["Принято всего"] * MIN_PERCENT / 100)
    ).astype(int)
    delta, threshold = pairs["Изменение, чел."], pairs["Порог, чел."]
    pairs[TARGET] = np.select([delta <= -threshold, delta >= threshold], ["Спад", "Рост"], default="Стабильно")
    return pairs.sort_values(["Год прогноза", "Код"])


def make_tree(depth, balance):
    return DecisionTreeClassifier(
        max_depth=depth,
        min_samples_leaf=8,
        class_weight="balanced" if balance else None,
        random_state=42,
    )


def select_tree(train):
    years = sorted(train["Год прогноза"].unique())
    folds = [(train[train["Год прогноза"] < y], train[train["Год прогноза"] == y]) for y in years[1:]]
    if not folds or any(set(a[TARGET]) != set(CLASSES) or set(b[TARGET]) != set(CLASSES) for a, b in folds):
        raise ValueError("В хронологической проверке должны быть представлены все три класса")

    results = []
    for depth in (2, 3, 4):
        for balance in (False, True):
            scores = {}
            for past, validation in folds:
                model = make_tree(depth, balance).fit(past[FEATURES], past[TARGET])
                predictions = model.predict(validation[FEATURES])
                year = int(validation["Год прогноза"].iloc[0])
                scores[f"Macro F1 {year}"] = f1_score(
                    validation[TARGET], predictions, labels=CLASSES, average="macro", zero_division=0
                )
            results.append({
                "Глубина": depth,
                "Взвешивание классов": balance,
                **scores,
                "Среднее Macro F1": float(np.mean(list(scores.values()))),
            })
    scores = pd.DataFrame(results)
    best = scores.loc[scores["Среднее Macro F1"].idxmax()]
    scores["Выбрана модель"] = scores.index == best.name
    return make_tree(int(best["Глубина"]), bool(best["Взвешивание классов"])), scores


def evaluate(actual, predictions):
    metrics, by_class, matrix = [], [], []
    for model, estimates in predictions.items():
        precision, recall, f1, support = precision_recall_fscore_support(
            actual, estimates, labels=CLASSES, zero_division=0
        )
        metrics.append({
            "Модель": model,
            "Accuracy": accuracy_score(actual, estimates),
            "Macro F1": f1.mean(),
            "Наблюдений": len(actual),
        })
        for label, p, r, f, n in zip(CLASSES, precision, recall, f1, support, strict=True):
            by_class.append({"Модель": model, "Класс": label, "Precision": p, "Recall": r, "F1": f, "Наблюдений": n})
        confusion = confusion_matrix(actual, estimates, labels=CLASSES)
        for i, actual_label in enumerate(CLASSES):
            for j, predicted_label in enumerate(CLASSES):
                matrix.append({
                    "Модель": model,
                    "Фактический класс": actual_label,
                    "Прогнозный класс": predicted_label,
                    "Число специальностей": confusion[i, j],
                })
    return pd.DataFrame(metrics), pd.DataFrame(by_class), pd.DataFrame(matrix)


def main():
    folder = Path(__file__).resolve().parent
    pairs = make_pairs(folder.parents[1] / "excel" / "01_specialties.xlsx")
    train = pairs.loc[pairs["Год прогноза"] < TEST_YEAR]
    test = pairs.loc[pairs["Год прогноза"] == TEST_YEAR]
    if test.empty:
        raise ValueError("Нет пар для контрольного учебного года")

    tree, scores = select_tree(train)
    models = {BASELINE: DummyClassifier(strategy="most_frequent"), TREE: tree}
    predictions = {
        name: model.fit(train[FEATURES], train[TARGET]).predict(test[FEATURES])
        for name, model in models.items()
    }
    metrics, class_metrics, confusion = evaluate(test[TARGET], predictions)
    forecasts = test[[
        "Код", "Название", "Учебный год", "Год прогноза", "Техническая область",
        "Принято всего", "Приём следующего года", "Изменение, чел.", "Порог, чел.", TARGET,
    ]].copy()
    for name, prediction in predictions.items():
        forecasts[name] = prediction
    class_balance = (
        pairs.loc[pairs["Год прогноза"] <= TEST_YEAR]
        .groupby(["Год прогноза", TARGET]).size()
        .rename("Число специальностей").reset_index()
    )

    result = folder / "result"
    result.mkdir(exist_ok=True)
    for name, table in {
        "metrics.csv": metrics,
        "class_metrics.csv": class_metrics,
        "confusion_matrix.csv": confusion,
        "predictions.csv": forecasts,
        "validation_scores.csv": scores,
        "class_balance.csv": class_balance,
    }.items():
        table.to_csv(result / name, index=False, encoding="utf-8-sig")

    print(f"Обучение: {len(train)}. Проверка: {len(test)}.")
    print("Порог: минимум 10 человек и 15% текущего приёма.")
    print(scores.loc[scores["Выбрана модель"]].to_string(index=False))
    print(metrics.round(3).to_string(index=False))
    print(class_metrics.round(3).to_string(index=False))


if __name__ == "__main__":
    main()
