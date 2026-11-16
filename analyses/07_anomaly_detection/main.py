import os
os.environ.setdefault('OMP_NUM_THREADS', '1')
from pathlib import Path
import pandas as pd
import numpy as np
from sklearn.ensemble import IsolationForest
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
FEATURES = ['Принято всего', 'Численность всего', 'Выпуск всего']
LEVEL = 'Уровень'
INTAKE, COHORT, GRADUATES = FEATURES

def anomalies(path):
    history = pd.read_csv(path, dtype={'КАТО': str})
    history = history.loc[history['Год'].between(2000, 2025)].copy()
    keys = [LEVEL, 'КАТО', 'Год']
    if history.duplicated(keys).any():
        raise ValueError('Исторические ряды содержат повторяющиеся КАТО и годы')
    current = history.sort_values(keys)
    changed_2022 = {'100000000', '190000000', '330000000', '350000000', '620000000', '630000000'}
    changed_2018 = {'610000000', '790000000'}
    current['Границы'] = np.select([current['КАТО'].isin(changed_2022) & current['Год'].ge(2022), current['КАТО'].isin(changed_2018) & current['Год'].ge(2018)], ['с 2022', 'с 2018'], default='ранние')
    parts = current.groupby([LEVEL, 'КАТО', 'Границы'])
    prior_year = parts['Год'].shift()
    previous = parts[INTAKE].shift()
    current['Темп'] = current[INTAKE].div(previous.where(previous.gt(0))) - 1
    current = current.loc[current['Год'].eq(prior_year + 1)].dropna(subset=['Темп']).copy()
    parts = current.groupby([LEVEL, 'КАТО', 'Границы'])
    current['Сколько сопоставимых изменений'] = parts['Темп'].transform('size')
    current['Типичный темп'] = parts['Темп'].transform('median')
    current = current.loc[current['Сколько сопоставимых изменений'].ge(3)].copy()
    current['Отклонение в п.п.'] = 100 * (current['Темп'] - current['Типичный темп'])
    all_rows = []
    for _, sample in current.groupby(LEVEL):
        x = sample[['Отклонение в п.п.']].copy()
        q1, q3 = x.iloc[:, 0].quantile([0.25, 0.75])
        span = q3 - q1
        result = sample.copy()
        result['IQR'] = ~x.iloc[:, 0].between(q1 - 1.5 * span, q3 + 1.5 * span)
        result['Isolation Forest'] = IsolationForest(n_estimators=300, contamination=0.1, random_state=42).fit_predict(x).astype(int) == -1
        result['Оба метода'] = result['IQR'] & result['Isolation Forest']
        all_rows.append(result)
    result = pd.concat(all_rows, ignore_index=True)
    result['Учебный год'] = result['Год'].map(lambda y: f'{y}/{(y + 1) % 100:02d}')
    summaries = result.groupby(LEVEL, as_index=False).agg(Наблюдений=('Темп', 'size'), IQR=('IQR', 'sum'), **{'Isolation Forest': ('Isolation Forest', 'sum'), 'Оба метода': ('Оба метода', 'sum')})
    return (summaries, result)

def main():
    tables = dict(zip(('metrics', 'cases'), anomalies(HERE / 'regional_source.csv'), strict=True))
    result = HERE / 'result'
    result.mkdir(exist_ok=True)
    for name, data in tables.items():
        path = result / f'{name}.csv'
        data.to_csv(path, index=False, encoding='utf-8-sig')
        print(f'{path.name}: {len(data)} строк', flush=True)
if __name__ == '__main__':
    main()
