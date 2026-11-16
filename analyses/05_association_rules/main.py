import os
os.environ.setdefault('OMP_NUM_THREADS', '1')
from pathlib import Path
import pandas as pd
from collections import Counter
from itertools import combinations
from math import ceil
from itertools import product
from sklearn.neighbors import KNeighborsRegressor
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
FEATURES = ['Принято всего', 'Численность всего', 'Выпуск всего']
LEVEL = 'Уровень'
INTAKE, COHORT, GRADUATES = FEATURES
TRANSITION = 'Переход классификатора'
TECHNICAL = ['ИКТ', 'Инженерные, обрабатывающие и строительные отрасли']
MIN_SUPPORT = 0.08
MIN_CONFIDENCE = 0.55
MIN_ANTECEDENT = 10

def read_book(file):
    return pd.read_excel(ROOT / 'excel' / file, sheet_name='итоги', dtype={'Код': str, 'КАТО': str})

def technical_profiles():
    data = read_book('01_specialties.xlsx')
    data = data.loc[data['Техническая область'].isin(TECHNICAL)].copy()
    data['Год'] = data['Учебный год'].str[:4].astype(int)
    data[TRANSITION] = data[LEVEL].eq('СУЗы') & data['Год'].eq(2021)
    return data

def apriori(transactions):
    """Набор включается в кандидаты, только если все его поднаборы частые."""
    minimum = ceil(MIN_SUPPORT * len(transactions))
    single_counts = Counter((frozenset([item]) for transaction in transactions for item in transaction))
    level = {items: count for items, count in single_counts.items() if count >= minimum}
    frequent = level.copy()
    for size in (2, 3):
        items = sorted(set().union(*level))
        candidates = [frozenset(candidate) for candidate in combinations(items, size) if all((frozenset(subset) in level for subset in combinations(candidate, size - 1)))]
        level = {candidate: count for candidate in candidates if (count := sum((candidate <= transaction for transaction in transactions))) >= minimum}
        frequent.update(level)
    return frequent

def rule_stats(transactions, premise, conclusion):
    n = len(transactions)
    a = sum((premise <= row for row in transactions))
    b = sum((conclusion <= row for row in transactions))
    both = sum((premise | conclusion <= row for row in transactions))
    confidence = both / a if a else None
    return {'Строк': n, 'Условие': a, 'Следствие': b, 'Совпадений': both, 'Support': both / n if n else None, 'Confidence': confidence, 'Lift': confidence / (b / n) if a and b else None}

def discover(transactions):
    rules = []
    for items, count in apriori(transactions).items():
        if len(items) < 2:
            continue
        for item in sorted(items):
            conclusion = frozenset([item])
            premise = items - conclusion
            stats = rule_stats(transactions, premise, conclusion)
            if stats['Условие'] >= MIN_ANTECEDENT and stats['Confidence'] >= MIN_CONFIDENCE and (stats['Lift'] > 1):
                assert stats['Совпадений'] == count
                rules.append((premise, conclusion, stats))
    return sorted(rules, key=lambda row: (-row[2]['Lift'], -row[2]['Support'], tuple(sorted(row[0])), tuple(row[1])))

def association():
    title = 'Учебный год'
    college = read_book('03_urban_rural.xlsx').assign(Уровень='СУЗы')
    college = college.loc[college[title].ge('2022/23')].dropna(subset=[INTAKE, GRADUATES])
    uni = technical_profiles()
    uni = uni.loc[uni[LEVEL].eq('Вузы')].dropna(subset=[INTAKE, GRADUATES]).assign(Местность='По стране')
    rows = []
    for level, data in [('СУЗы', college), ('Вузы', uni)]:
        for holdout_year, test in data.groupby(title):
            train = data.loc[data[title].lt(holdout_year)].copy()
            if train.empty:
                continue
            cutoffs = train[[INTAKE, GRADUATES]].median()
            if level == 'СУЗы':
                pairs = test.groupby('Код')['Местность'].nunique()
                used = test.loc[test['Код'].isin(pairs.loc[pairs.eq(2)].index)]
            else:
                used = test

            def transactions(sample):
                return [frozenset(('Приём высокий' if i >= cutoffs[INTAKE] else 'Приём низкий', 'Выпуск высокий' if o >= cutoffs[GRADUATES] else 'Выпуск низкий', str(a))) for i, o, a in sample[[INTAKE, GRADUATES, 'Техническая область']].itertuples(index=False, name=None)]
            rules = discover(transactions(train))
            for premise, conclusion, _ in rules:
                for locality, sub in used.groupby('Местность'):
                    rows.append({LEVEL: level, title: holdout_year, 'Местность': locality, 'Если': ' + '.join(sorted(premise)), 'То': next(iter(conclusion)), **rule_stats(transactions(sub), premise, conclusion)})
    return pd.DataFrame(rows)

def candidates():
    features = 'Масштаб, структура и динамика'
    return [(KNeighborsRegressor(n_neighbors=k, weights='distance', algorithm='kd_tree'), features, weight) for k, weight in product((5, 3, 7), (0.5, 0.25, 1.0))]

def main():
    tables = {'rules': association()}
    result = HERE / 'result'
    result.mkdir(exist_ok=True)
    for name, data in tables.items():
        path = result / f'{name}.csv'
        data.to_csv(path, index=False, encoding='utf-8-sig')
        print(f'{path.name}: {len(data)} строк', flush=True)
if __name__ == '__main__':
    main()
