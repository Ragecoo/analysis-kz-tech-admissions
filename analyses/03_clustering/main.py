import os
os.environ.setdefault('OMP_NUM_THREADS', '1')
from pathlib import Path
import pandas as pd
import numpy as np
from itertools import product
from sklearn.cluster import KMeans
from sklearn.metrics import adjusted_rand_score, silhouette_score
from sklearn.neighbors import KNeighborsRegressor
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
LEVEL = 'Уровень'

def read_book(file):
    return pd.read_excel(ROOT / 'excel' / file, sheet_name='итоги', dtype={'Код': str, 'КАТО': str})

def clustering():
    data = read_book('04_age_groups.xlsx')
    data = data[data['Регион'].ne('Республика Казахстан')]
    keys = [LEVEL, 'Учебный год', 'Регион']
    counts = data.pivot(index=keys, columns='Возрастная группа', values='Численность')
    totals = data.groupby(keys)['Контингент'].first().reindex(counts.index)
    shares = counts.div(totals, axis=0)
    shares = shares.loc[shares.notna().all(axis=1) & totals.gt(0)].reset_index()
    columns = list(counts.columns)
    rows, labels = ([], [])
    for level, panel in shares.groupby(LEVEL):
        variant = 'Корень из долей' if level == 'Вузы' else 'Доли'
        choices = {}
        for year, sample in panel.groupby('Учебный год', sort=True):
            x = sample[columns]
            if variant == 'Корень из долей':
                x = np.sqrt(x)
            candidates = [(k, KMeans(n_clusters=k, n_init=30, random_state=42).fit(x)) for k in (2, 3, 4)]
            rank = [(k, model, silhouette_score(x, model.labels_)) for k, model in candidates]
            k, chosen, quality = max(rank, key=lambda entry: entry[2])
            u18 = columns.index('до 18 лет')
            order = np.argsort(-chosen.cluster_centers_[:, u18])
            label = {i: number for number, i in enumerate(order, 1)}
            mapped = sample.assign(Кластер=[label[i] for i in chosen.labels_])
            labels.append(mapped)
            based = mapped.set_index('Регион')['Кластер']
            choices[year] = based
            rows.append({LEVEL: level, 'Учебный год': year, 'Кластеров': k, 'Silhouette': quality, 'Регионов': len(x)})
        ref_year = max(choices)
        ref = choices[ref_year]
        for row in rows:
            if row[LEVEL] != level:
                continue
            year = row['Учебный год']
            common = choices[year].index.intersection(ref.index)
            row['Совпадение с последним годом, ARI'] = adjusted_rand_score(choices[year].loc[common], ref.loc[common])
    members = pd.concat(labels, ignore_index=True)
    groups = members.groupby([LEVEL, 'Учебный год', 'Кластер'])
    profiles = groups[columns].mean().join(groups.size().rename('Регионов')).reset_index()
    return (pd.DataFrame(rows), members, profiles)

def candidates():
    features = 'Масштаб, структура и динамика'
    return [(KNeighborsRegressor(n_neighbors=k, weights='distance', algorithm='kd_tree'), features, weight) for k, weight in product((5, 3, 7), (0.5, 0.25, 1.0))]

def main():
    tables = dict(zip(('metrics', 'clusters', 'profiles'), clustering(), strict=True))
    result = HERE / 'result'
    result.mkdir(exist_ok=True)
    for name, data in tables.items():
        path = result / f'{name}.csv'
        data.to_csv(path, index=False, encoding='utf-8-sig')
        print(f'{path.name}: {len(data)} строк', flush=True)
if __name__ == '__main__':
    main()
