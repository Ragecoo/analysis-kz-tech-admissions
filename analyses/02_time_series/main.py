import os
os.environ.setdefault('OMP_NUM_THREADS', '1')
from pathlib import Path
import pandas as pd
from itertools import product
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
YEAR = 'Год прогноза'
LEVEL = 'Уровень'

def time_series():
    source = pd.read_excel(ROOT / 'excel' / '07_admissions.xlsx', sheet_name='Приём').dropna(subset=['Год начала'])
    rows, future, history = ([], [], [])
    for level, suffix in (('СУЗы', 'СУЗы'), ('Вузы', 'вузы')):
        facts = source.set_index('Год начала')[f'Технический приём — {suffix}'].sort_index()
        labels = source.set_index('Год начала')[f'Метод — {suffix}']

        def estimate(past, horizon):
            drift = past.diff().dropna().tail(5).median()
            return past.iloc[-1] + drift * sum((0.8 ** step for step in range(horizon)))
        for year, horizon in product(labels[labels.eq('Итог БНС')].index, range(1, 7)):
            past = facts.loc[facts.index <= year - horizon]
            if len(past) >= 5:
                prediction = estimate(past, horizon)
                rows.append({LEVEL: level, YEAR: int(year), 'Горизонт': horizon, 'Факт': facts.loc[year], 'Прогноз': prediction, 'Абсолютная ошибка': abs(prediction - facts.loc[year])})
        for year, actual in facts.items():
            history.append({LEVEL: level, 'Год': int(year), 'Факт': actual, 'Источник': labels.loc[year]})
        for horizon in range(1, 7):
            year = int(facts.index.max()) + horizon
            future.append({LEVEL: level, YEAR: year, 'Горизонт': horizon, 'Прогноз': estimate(facts, horizon)})
    cases = pd.DataFrame(rows)
    stats = cases.groupby([LEVEL, 'Горизонт'], as_index=False).agg(MAE=('Абсолютная ошибка', 'mean'), Проверок=('Абсолютная ошибка', 'size'))
    return dict(metrics=stats, backtest=cases, forecast=pd.DataFrame(future), history=pd.DataFrame(history))

def main():
    tables = time_series()
    result = HERE / 'result'
    result.mkdir(exist_ok=True)
    for name, data in tables.items():
        path = result / f'{name}.csv'
        data.to_csv(path, index=False, encoding='utf-8-sig')
        print(f'{path.name}: {len(data)} строк', flush=True)
if __name__ == '__main__':
    main()
