from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.ensemble import IsolationForest


KEYS = ["Уровень", "КАТО"]
CHANGE = "Темп изменения"
BASELINE = "Типичный темп региона"
DEVIATION = "Отклонение темпа, п.п."
PEOPLE = "Отклонение от типичного приёма, чел."
CONTAMINATION = 0.10


def prepare(project):
    raw = pd.read_excel(project / "excel" / "02_regions.xlsx", sheet_name="итоги", dtype={"КАТО": str})
    rows = raw.loc[raw["Регион"].ne("Республика Казахстан")].copy()
    if rows["КАТО"].isna().any() or rows.duplicated([*KEYS, "Учебный год"]).any():
        raise ValueError("Неполные или повторяющиеся ключи региона")
    rows["Причина исключения"] = rows["Сопоставимость"].where(rows["Сопоставимость"].notna(), "")
    missing = rows["Принято годом ранее"].isna() | rows["Принято всего"].isna() | rows["Принято годом ранее"].le(0)
    rows.loc[missing & rows["Причина исключения"].eq(""), "Причина исключения"] = "Нет сопоставимой базы для расчёта"
    good = rows["Причина исключения"].eq("")
    rows.loc[good, "Изменение в людях"] = rows.loc[good, "Принято всего"] - rows.loc[good, "Принято годом ранее"]
    rows.loc[good, CHANGE] = rows.loc[good, "Изменение в людях"] / rows.loc[good, "Принято годом ранее"]
    used = rows.loc[good].sort_values([*KEYS, "Учебный год"]).copy()
    if used.groupby(KEYS).size().lt(3).any():
        raise ValueError("Для региона нужно минимум три сопоставимых изменения")

    # Для каждого наблюдения сравнение строится с остальными изменениями этого региона.
    typical = {}
    for _, group in used.groupby(KEYS):
        changes = group[CHANGE]
        typical.update({idx: changes.drop(idx).median() for idx in changes.index})
    used[BASELINE] = pd.Series(typical)
    gap = used[CHANGE] - used[BASELINE]
    used[DEVIATION] = gap * 100
    used[PEOPLE] = gap * used["Принято годом ранее"]
    return used, rows.loc[~good]


def detect(group):
    group = group.copy()
    q1, q3 = group[DEVIATION].quantile([0.25, 0.75])
    lower, upper = q1 - 1.5 * (q3 - q1), q3 + 1.5 * (q3 - q1)
    group["IQR"] = ~group[DEVIATION].between(lower, upper)
    model = IsolationForest(
        n_estimators=300, contamination=CONTAMINATION, max_samples=len(group), random_state=42
    )
    X = group[[DEVIATION, PEOPLE]]
    group["Isolation Forest"] = model.fit_predict(X) == -1
    group["Оценка Isolation Forest"] = -model.decision_function(X)
    group["Совпали оба метода"] = group["IQR"] & group["Isolation Forest"]
    group["Отметка"] = np.select(
        [group["Совпали оба метода"], group["IQR"], group["Isolation Forest"]],
        ["Оба метода", "Только IQR", "Только Isolation Forest"],
        default="Без отметки",
    )
    summary = {
        "Уровень": group["Уровень"].iloc[0],
        "Проверено": len(group),
        "IQR": int(group["IQR"].sum()),
        "Isolation Forest": int(group["Isolation Forest"].sum()),
        "Совпали оба метода": int(group["Совпали оба метода"].sum()),
        "Нижняя граница IQR, п.п.": lower,
        "Верхняя граница IQR, п.п.": upper,
        "Доля аномалий для Isolation Forest": CONTAMINATION,
    }
    return group, summary


def main():
    folder = Path(__file__).resolve().parent
    used, excluded = prepare(folder.parents[1])
    detections, summaries = zip(*(detect(group) for _, group in used.groupby("Уровень")))
    detections = pd.concat(detections, ignore_index=True)
    fields = [
        "Уровень", "КАТО", "Регион", "Учебный год", "Принято всего", "Принято годом ранее",
        "Изменение в людях", CHANGE, BASELINE, DEVIATION, PEOPLE,
        "IQR", "Isolation Forest", "Совпали оба метода", "Отметка", "Оценка Isolation Forest",
    ]
    detections = detections[fields].sort_values(["Уровень", "Учебный год", "КАТО"])
    anomalous = detections.loc[detections["Отметка"].ne("Без отметки")].copy()
    summary = pd.DataFrame(summaries)
    result_dir = folder / "result"
    result_dir.mkdir(exist_ok=True)
    for name, table in {
        "observations.csv": detections,
        "anomalies.csv": anomalous,
        "metrics.csv": summary,
        "excluded.csv": excluded[[
            "Уровень", "КАТО", "Регион", "Учебный год", "Принято всего",
            "Принято годом ранее", "Изменение в людях", CHANGE, "Причина исключения",
        ]],
    }.items():
        table.to_csv(result_dir / name, index=False, encoding="utf-8-sig")
    print(summary.round(3).to_string(index=False))
    print("\nСовпали оба метода:")
    print(anomalous.loc[anomalous["Совпали оба метода"], [
        "Уровень", "Регион", "Учебный год", "Изменение в людях", CHANGE
    ]].round(3).to_string(index=False))


if __name__ == "__main__":
    main()
