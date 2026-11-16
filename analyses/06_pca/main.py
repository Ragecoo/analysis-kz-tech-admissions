import os
os.environ.setdefault('OMP_NUM_THREADS', '1')
from pathlib import Path
import pandas as pd
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
FEATURES = ['Принято всего', 'Численность всего', 'Выпуск всего']
LEVEL = 'Уровень'
INTAKE, COHORT, GRADUATES = FEATURES
BUDGET = 'Доля госзаказа'
TRANSITION = 'Переход классификатора'
TECHNICAL = ['ИКТ', 'Инженерные, обрабатывающие и строительные отрасли']

def read_book(file):
    return pd.read_excel(ROOT / 'excel' / file, sheet_name='итоги', dtype={'Код': str, 'КАТО': str})

def technical_profiles():
    data = read_book('01_specialties.xlsx')
    data = data.loc[data['Техническая область'].isin(TECHNICAL)].copy()
    data['Год'] = data['Учебный год'].str[:4].astype(int)
    data[TRANSITION] = data[LEVEL].eq('СУЗы') & data['Год'].eq(2021)
    return data

def pca():
    data = technical_profiles()
    keys = [LEVEL, 'Код', 'Год']
    lag = data[[*keys, INTAKE]].rename(columns={INTAKE: 'Предыдущий приём'}).copy()
    lag['Год'] += 1
    data = data.merge(lag, on=keys, how='left', validate='one_to_one')
    data = data.loc[~data[TRANSITION] & data['Предыдущий приём'].gt(0)].dropna(subset=[INTAKE, COHORT, GRADUATES, 'Предыдущий приём', BUDGET])
    rows, loadings, scores = ([], [], [])
    for (level, year), group in data.groupby([LEVEL, 'Учебный год']):
        features = group[[INTAKE, COHORT, GRADUATES, BUDGET]].copy()
        features['Динамика приёма'] = group[INTAKE] / group['Предыдущий приём'] - 1
        if len(group) < len(features.columns) or features.nunique().min() <= 1:
            continue
        x = features.copy()
        scaled = StandardScaler().fit_transform(x)
        model = PCA().fit(scaled)
        key = {LEVEL: level, 'Учебный год': year, 'Наблюдений': len(x)}
        rows.append({**key, 'Доля PC1': model.explained_variance_ratio_[0], 'Доля PC1 и PC2': model.explained_variance_ratio_[:2].sum()})
        values = model.transform(scaled)
        scores.append(group[[LEVEL, 'Учебный год', 'Код', 'Название']].assign(PC1=values[:, 0], PC2=values[:, 1]))
        for i, weights in enumerate(model.components_[:2], 1):
            for feature, weight in zip(x, weights, strict=True):
                loadings.append({**key, 'Компонента': f'PC{i}', 'Признак': feature, 'Вклад': weight ** 2})
    return (pd.DataFrame(rows), pd.DataFrame(loadings), pd.concat(scores, ignore_index=True))

def main():
    tables = dict(zip(('metrics', 'loadings', 'scores'), pca(), strict=True))
    result = HERE / 'result'
    result.mkdir(exist_ok=True)
    for name, data in tables.items():
        path = result / f'{name}.csv'
        data.to_csv(path, index=False, encoding='utf-8-sig')
        print(f'{path.name}: {len(data)} строк', flush=True)
if __name__ == '__main__':
    main()
