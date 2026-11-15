import csv
from pathlib import Path

from openpyxl import load_workbook


def read_table(workbook, sheet_name, max_col=None):
    sheet = workbook[sheet_name]
    rows = []
    for row in sheet.iter_rows(max_col=max_col):
        errors = [cell.coordinate for cell in row if cell.data_type == "e"]
        if errors:
            raise ValueError(f"{sheet_name}: ошибка формулы в {', '.join(errors)}")
        values = tuple(cell.value for cell in row)
        if any(value is not None for value in values):
            rows.append(values)
    return rows


def main():
    analysis_dir = Path(__file__).resolve().parent
    workbook_path = analysis_dir / "forecast.xlsx"
    result_dir = analysis_dir / "result"

    # Читаем результаты, сохранённые Excel после пересчёта книги.
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        forecast = read_table(workbook, "Динамика")
        metrics = read_table(workbook, "Метрики")
        backtest = read_table(workbook, "Проверка")
        college_metrics = read_table(workbook, "СУЗы модели", max_col=7)
        college_backtest = read_table(workbook, "СУЗы горизонты")
    finally:
        workbook.close()

    for level in ("Вузы", "СУЗы"):
        rows = [row for row in forecast[1:] if row[2] == level]
        historical = [row for row in rows if isinstance(row[3], (int, float))]
        future = [row for row in rows if row[3] is None]
        if len(historical) != 15 or len(future) != 6 or any(
            not isinstance(row[6], (float, int)) or not isinstance(row[7], (float, int))
            for row in future
        ):
            raise ValueError(f"{level}: пересчитай и сохрани forecast.xlsx в Excel")

    if len(metrics) != 13 or any(
        row[3] != 3 or not all(isinstance(value, (int, float)) for value in row[4:6])
        for row in metrics[1:]
    ):
        raise ValueError("Метрики рассчитаны не для всех моделей и периодов")

    if len(college_backtest) != 37 or len(college_metrics) != 5 or any(
        row[1] != 9 or not all(isinstance(value, (int, float)) for value in row[2:])
        for row in college_metrics[1:]
    ):
        raise ValueError("Проверка прогноза СУЗов на горизонтах 1–3 года не завершена")

    result_dir.mkdir(exist_ok=True)
    tables = {
        "forecast.csv": forecast,
        "metrics.csv": metrics,
        "backtest.csv": backtest,
        "college_metrics_by_horizon.csv": college_metrics,
        "college_backtest_by_horizon.csv": college_backtest,
    }
    for filename, rows in tables.items():
        with (result_dir / filename).open("w", encoding="utf-8-sig", newline="") as file:
            csv.writer(file).writerows(rows)
        print(f"{filename} — {len(rows) - 1} строк")


if __name__ == "__main__":
    main()
