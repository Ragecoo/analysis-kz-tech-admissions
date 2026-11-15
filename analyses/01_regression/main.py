from pathlib import Path
from math import sqrt
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import pandas as pd

analysis_dir = Path(__file__).resolve().parent
project = analysis_dir.parents[1]
data = pd.read_excel(
    project / "excel" / "01_specialties.xlsx",
    sheet_name="итоги",
    dtype={"Код": str},
)

technical_areas = [
    "ИКТ",
    "Инженерные, обрабатывающие и строительные отрасли",
]

profiles = data.loc[
    data["Уровень"].eq("СУЗы")
    & data["Техническая область"].isin(technical_areas)
].copy()

features = ["Принято всего", "Численность всего", "Выпуск всего"]
target = "Принято в следующем году"

profiles["Год"] = profiles["Учебный год"].str[:4].astype(int)

current = profiles[["Код", "Название", "Год", *features]].copy()
current["Год прогноза"] = current["Год"] + 1

actual = profiles[["Код", "Год", "Принято всего"]].rename(
    columns={"Год": "Год прогноза", "Принято всего": target}
)

pairs = current.merge(
    actual,
    on=["Код", "Год прогноза"],
    how="inner",
    validate="one_to_one",
).dropna(subset=[*features, target])

train = pairs.loc[pairs["Год прогноза"] < 2025].copy()
test = pairs.loc[pairs["Год прогноза"] == 2025].copy()

print(f"Обучение: {len(train)} строк")
print(f"Проверка: {len(test)} строк")
print(test[["Код", "Название", *features, target]].head())

predictions = {
    "Без изменений": test["Принято всего"],
}

model_features = {
    "Регрессия по приёму": ["Принято всего"],
    "Регрессия с контингентом и выпуском": features,
}

coefficients = []
for name, columns in model_features.items():
    model = LinearRegression()
    model.fit(train[columns], train[target])
    predictions[name] = model.predict(test[columns])
    terms = ["Свободный член", *columns]
    estimates = [model.intercept_, *model.coef_]
    coefficients.extend(
        {"Модель": name, "Параметр": term, "Коэффициент": estimate}
        for term, estimate in zip(terms, estimates, strict=True)
    )

metrics = pd.DataFrame([
    {
        "Модель": name,
        "MAE": mean_absolute_error(test[target], forecast),
        "RMSE": sqrt(mean_squared_error(test[target], forecast)),
        "R²": r2_score(test[target], forecast),
    }
    for name, forecast in predictions.items()
]).set_index("Модель")

print(metrics.round(3))

improvement = (
    metrics.loc["Регрессия по приёму", "MAE"]
    - metrics.loc["Регрессия с контингентом и выпуском", "MAE"]
)
print(f"\nСнижение MAE от добавления контингента и выпуска: {improvement:.2f} чел.")

result_dir = analysis_dir / "result"
result_dir.mkdir(exist_ok=True)

metrics.to_csv(result_dir / "metrics.csv", encoding="utf-8-sig")
pd.DataFrame(coefficients).to_csv(
    result_dir / "coefficients.csv", index=False, encoding="utf-8-sig"
)

forecasts = test[["Код", "Название", "Год", "Год прогноза", target]].copy()
for name, forecast in predictions.items():
    forecasts[name] = forecast

forecasts.to_csv(result_dir / "predictions.csv", index=False, encoding="utf-8-sig")
print(f"\nМетрики, коэффициенты и прогнозы сохранены в {result_dir}")
