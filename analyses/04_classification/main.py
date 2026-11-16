import os
os.environ.setdefault('OMP_NUM_THREADS', '1')
from pathlib import Path
import pandas as pd
import numpy as np
from itertools import product
from sklearn.preprocessing import StandardScaler
from sklearn.base import clone
from sklearn.impute import SimpleImputer
from sklearn.metrics import f1_score
from sklearn.pipeline import make_pipeline
from sklearn.tree import DecisionTreeClassifier
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
LEVELS = ('СУЗы', 'Вузы')
FEATURES = ['Принято всего', 'Численность всего', 'Выпуск всего']
TARGET = 'Принято в следующем году'
YEAR = 'Год прогноза'
LEVEL = 'Уровень'
INTAKE, COHORT, GRADUATES = FEATURES
GROWTH = 'Изменение приёма'
BUDGET = 'Доля госзаказа'
TRANSITION = 'Переход классификатора'
TECHNICAL = ['ИКТ', 'Инженерные, обрабатывающие и строительные отрасли']
LABELS = ['Спад', 'Стабильно', 'Рост']
PERIOD = 'Период'
ALL_YEARS = 'Все проверочные годы'

def read_pairs():
    source = ROOT / 'excel' / '01_specialties.xlsx'
    data = pd.read_excel(source, sheet_name='итоги', dtype={'Код': str})
    technical = ['ИКТ', 'Инженерные, обрабатывающие и строительные отрасли']
    profiles = data.loc[data['Уровень'].isin(LEVELS) & data['Техническая область'].isin(technical)].copy()
    profiles['Год'] = profiles['Учебный год'].str[:4].astype(int)
    keys = ['Уровень', 'Код']
    current = profiles[[*keys, 'Название', 'Год', *FEATURES]].copy()
    current[YEAR] = current['Год'] + 1
    actual = profiles[[*keys, 'Год', FEATURES[0]]].rename(columns={'Год': YEAR, FEATURES[0]: TARGET})
    return current.merge(actual, on=[*keys, YEAR], how='inner', validate='one_to_one').dropna(subset=[*FEATURES, TARGET])

def read_university_archive(pairs):
    source = ROOT / 'excel' / 'util' / 'technical-codes.csv'
    archive = pd.read_csv(source, sep=';', dtype={'code': str})
    first = pairs.loc[pairs['Уровень'].eq('Вузы'), 'Год'].min()
    history = archive.loc[archive['level'].eq('universities') & archive['selection'].eq('modern_detail') & archive['year'].between(first - 2, first)].rename(columns={'code': 'Код', 'name': 'Название', 'year': 'Год', 'admissions': FEATURES[0], 'students': FEATURES[1], 'graduates': FEATURES[2]}).assign(Уровень='Вузы')
    keys = ['Уровень', 'Код']
    actual = history[[*keys, 'Год', FEATURES[0]]].rename(columns={'Год': YEAR, FEATURES[0]: TARGET})
    current = history.loc[history['Год'].lt(first), [*keys, 'Название', 'Год', *FEATURES]].copy()
    current[YEAR] = current['Год'] + 1
    early = current.merge(actual, on=[*keys, YEAR], validate='one_to_one').dropna(subset=[*FEATURES, TARGET])
    return (early, history.loc[history['Год'].lt(first), [*keys, 'Год', FEATURES[0]]])

def read_book(file):
    return pd.read_excel(ROOT / 'excel' / file, sheet_name='итоги', dtype={'Код': str, 'КАТО': str})

def technical_profiles():
    data = read_book('01_specialties.xlsx')
    data = data.loc[data['Техническая область'].isin(TECHNICAL)].copy()
    data['Год'] = data['Учебный год'].str[:4].astype(int)
    data[TRANSITION] = data[LEVEL].eq('СУЗы') & data['Год'].eq(2021)
    return data

def prepare_pairs():
    published = read_pairs()
    archive, history = read_university_archive(published)
    pairs = pd.concat([published, archive], ignore_index=True)
    keys = [LEVEL, 'Код', 'Год']
    profiles = technical_profiles()
    pairs = pairs.merge(profiles[[*keys, BUDGET, TRANSITION]], on=keys, how='left', validate='one_to_one')
    previous = pd.concat([pairs[[*keys, INTAKE]], history], ignore_index=True)
    if previous.groupby(keys)[INTAKE].nunique().gt(1).any():
        raise ValueError('Различаются значения приёма для одного кода и года')
    previous = previous.drop_duplicates(keys).rename(columns={INTAKE: 'Приём годом ранее'})
    previous['Год'] += 1
    pairs = pairs.merge(previous, on=keys, how='left', validate='one_to_one')
    pairs[TRANSITION] = pairs[TRANSITION].eq(True)
    pairs['Инженерная область'] = pairs['Код'].str.startswith(('07', '6B07')).astype(int)
    pairs[GROWTH] = np.log1p(pairs[INTAKE]) - np.log1p(pairs['Приём годом ранее'])
    pairs['Учебный год'] = pairs[YEAR].map(lambda y: f'{y}/{(y + 1) % 100:02d}')
    pairs['Порог спада'] = np.maximum(10, np.ceil(pairs[INTAKE] * 0.15))
    delta = pairs[TARGET] - pairs[INTAKE]
    pairs['Класс'] = np.select([delta.le(-pairs['Порог спада']), delta.ge(pairs['Порог спада'])], ['Спад', 'Рост'], default='Стабильно')
    pairs['Спад'] = pairs['Класс'].eq('Спад').astype(int)
    return pairs.sort_values([LEVEL, YEAR, 'Код'])

def clean_inputs(frame):
    cohort = frame[COHORT].mask(frame[TRANSITION])
    graduates = frame[GRADUATES].mask(frame[TRANSITION])
    return pd.DataFrame({'Логарифм приёма': np.log1p(frame[INTAKE]), 'Логарифм контингента': np.log1p(cohort), 'Логарифм выпуска': np.log1p(graduates), 'Приём к контингенту': np.log1p(frame[INTAKE] / (cohort + 1)), 'Выпуск к контингенту': np.log1p(graduates / (cohort + 1)), GROWTH: frame[GROWTH], BUDGET: frame[BUDGET].where(frame[BUDGET].between(0, 1))}, index=frame.index)

def design(frame, feature_set):
    x = clean_inputs(frame)
    if feature_set == 'Масштаб':
        return x.iloc[:, :1]
    if feature_set == 'Масштаб и структура':
        return x.iloc[:, :5]
    return x

def predict(estimator, train, test, feature_set, weight):
    features = design(train, feature_set)
    test_features = design(test, feature_set)
    fitted = clone(estimator)
    model = make_pipeline(SimpleImputer(strategy='median', keep_empty_features=True), StandardScaler(), fitted)
    return model.fit(features, train['Класс']).predict(test_features)

def available(train):
    if len(train) < 7:
        return False
    return train['Класс'].nunique() >= 2

def score(test, prediction):
    return f1_score(test['Класс'], prediction, labels=LABELS, average='macro', zero_division=0)

def candidates():
    features = 'Масштаб, структура и динамика'
    return [(DecisionTreeClassifier(max_depth=d, min_samples_leaf=leaf, class_weight=w, random_state=42), features, 1.0) for d, leaf, w in product((2, 3), (2, 8), (None, 'balanced'))]

def fit_early(train, test):
    folds = []
    for year, holdout in train.groupby(YEAR):
        past = train.loc[train[YEAR].lt(year)]
        if available(past):
            folds.append((past, holdout))
    grid = candidates()
    best = grid[0]
    quality = float('-inf')
    if folds:
        validation = pd.concat([holdout for _, holdout in folds])
        for estimator, feature_set, weight in grid:
            predicted = np.concatenate([predict(estimator, past, holdout, feature_set, weight) for past, holdout in folds])
            value = score(validation, predicted)
            if value > quality:
                best = (estimator, feature_set, weight)
                quality = value
    estimator, feature_set, weight = best
    return (predict(estimator, train, test, feature_set, weight), {'Модель': estimator.__class__.__name__, 'Параметры': str(estimator), 'Признаки': feature_set, 'Порог или вес': weight, 'Годы подбора': ', '.join((str(v[YEAR].iloc[0]) for _, v in folds))})

def class_metrics(y, p, positive):
    actual, alert = (y.eq(positive), p.eq(positive))
    tp, fp = (int((actual & alert).sum()), int((~actual & alert).sum()))
    fn, tn = (int((actual & ~alert).sum()), int((~actual & ~alert).sum()))
    return {'Число случаев': tp + fn, 'Верные': tp, 'Ложные': fp, 'Пропущенные': fn, 'Верно без класса': tn, 'Precision': tp / (tp + fp) if tp + fp else None, 'Recall': tp / (tp + fn) if tp + fn else None, 'F1': 2 * tp / (2 * tp + fp + fn) if 2 * tp + fp + fn else None}

def summarize(predictions):
    rows, classes = ([], [])
    for level, group in predictions.groupby(LEVEL):
        for period, data in [(ALL_YEARS, group), *group.groupby(PERIOD, sort=True)]:
            y, p = (data['Факт'], data['Прогноз'])
            row = {LEVEL: level, PERIOD: period, 'Наблюдений': len(data)}
            positive = 'Спад'
            event = class_metrics(y, p, positive)
            row.update({'Accuracy': y.eq(p).mean(), 'F1': f1_score(y, p, labels=LABELS, average='macro', zero_division=0), 'Спадов': event['Число случаев'], 'Найдено спадов': event['Верные'], 'Ложные предупреждения': event['Ложные'], 'Пропущено спадов': event['Пропущенные'], 'Precision спада': event['Precision'], 'Recall спада': event['Recall']})
            labels = LABELS
            for label in labels:
                classes.append({LEVEL: level, PERIOD: period, 'Класс': label, **class_metrics(y, p, label)})
            rows.append(row)
    return (pd.DataFrame(rows), pd.DataFrame(classes))

def evaluate(pairs):
    rows, settings = ([], [])
    label = 'Класс'
    for (level, year), test in pairs.groupby([LEVEL, YEAR]):
        train = pairs.loc[pairs[LEVEL].eq(level) & pairs[YEAR].lt(year)]
        if not available(train):
            continue
        prediction, params = fit_early(train, test)
        settings.append({LEVEL: level, YEAR: year, 'Строк обучения': len(train), 'Последний год обучения': train[YEAR].max(), **params})
        result = test[[LEVEL, 'Код', 'Название', YEAR, INTAKE]].copy()
        result[PERIOD] = test['Учебный год']
        result['Факт'] = test[label]
        result['Прогноз'] = prediction
        rows.append(result)
    predictions = pd.concat(rows, ignore_index=True)
    metrics, classes = summarize(predictions)
    tables = {'metrics': metrics, 'predictions': predictions, 'settings': pd.DataFrame(settings)}
    if not classes.empty:
        tables['classes'] = classes
    return tables

def main():
    tables = evaluate(prepare_pairs())
    result = HERE / 'result'
    result.mkdir(exist_ok=True)
    for name, data in tables.items():
        path = result / f'{name}.csv'
        data.to_csv(path, index=False, encoding='utf-8-sig')
        print(f'{path.name}: {len(data)} строк', flush=True)
if __name__ == '__main__':
    main()
