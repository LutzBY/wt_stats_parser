import json
import re
import pandas as pd

# --- Загрузка данных ---
path = r'C:\Users\lutsevich\Desktop\py\wt_stats\wt_stats_parser\vehicles_rus.json'
with open(path, encoding='UTF-8') as file:
    vehicles_rus_raw = json.load(file)

# --- Таблица конвертации rb → BR ---
RB_TO_BR = {
    0: 1.0, 1: 1.3, 2: 1.7, 3: 2.0, 4: 2.3, 5: 2.7, 6: 3.0, 7: 3.3,
    8: 3.7, 9: 4.0, 10: 4.3, 11: 4.7, 12: 5.0, 13: 5.3, 14: 5.7,
    15: 6.0, 16: 6.3, 17: 6.7, 18: 7.0, 19: 7.3, 20: 7.7, 21: 8.0,
    22: 8.3, 23: 8.7, 24: 9.0, 25: 9.3, 26: 9.7, 27: 10.0,
    28: 10.3, 29: 10.7, 30: 11.0, 31: 11.3, 32: 11.7, 33: 12.0,
    34: 12.3, 35: 12.7, 36: 13.0, 37: 13.3, 38: 13.7, 39: 14.0, 40: 14.3
}

# Таблица замен для стран
COUNTRY_TO_RUSSIAN = {
    'ussr': 'СССР',
    'germany': 'Германия',
    'usa': 'США',
    'britain': 'Великобритания',
    'france': 'Франция',
    'japan': 'Япония',
    'china': 'Китай',
    'czech': 'Чехословакия',
    'sweden': 'Швеция',
    'poland': 'Польша',
    'italy': 'Италия',
    'israel': 'Израиль'
}

COUNTRY_SYMBOLS = {
    '🇺🇸': 'США',
    '🇩🇪': 'Германия',
    '🇬🇧': 'Великобритания',
    '🇫🇷': 'Франция',
    '🇯🇵': 'Япония',
    '🇨🇳': 'Китай',
    '🇮🇹': 'Италия',
    '🇨🇿': 'Чехословакия',
    '🇸🇪': 'Швеция',
    '🇵🇱': 'Польша'
}

# Таблица замен для HTML-символов
HTML_REPLACEMENTS = {
    '&#039;': "'",  # апостроф
    '&amp;': '&',   # &
    '<': '<',    # <
    '>': '>',     # >
    '&quot;': '"' # ""
}

# Заменяет HTML-символы, заменяет символы стран на текст
def normalize_name(name, country_operator=None): 

    if not isinstance(name, str):
        name = str(name)

    # 1. Заменяем HTML-символы
    for html, text in HTML_REPLACEMENTS.items():
        name = name.replace(html, text)

    # 2. Заменяем эмодзи-флаги на текст (🇺🇸 → США)
    for symbol, country in COUNTRY_SYMBOLS.items():
        name = name.replace(symbol, country)

    # 3. Обработка ␗Имя → Имя (Страна)
    if '␗' or '▄' in name and country_operator:
        clean_name = name.replace('␗', '').strip()
        clean_name = name.replace('▄', '').strip()
        country_rus = COUNTRY_TO_RUSSIAN.get(country_operator.lower())
        if country_rus:
            name = f"{clean_name}({country_rus})"
        else:
            name = clean_name  # на всякий случай

    # 4. Финальная очистка
    name = re.sub(r'\s+', ' ', name.strip())

    return name

# Чистим базу: нормализуем display_name (индекс [1])
vehicles_rus = []
for item in vehicles_rus_raw:
    if len(item) < 8:
        continue
    # Копируем, чтобы не портить оригинал
    clean_item = item.copy()
    original_name = item[1]
    clean_name = normalize_name(original_name, item[2])
    clean_item[1] = clean_name  # заменяем на чистое имя
    vehicles_rus.append(clean_item)

# --- Функция: получить данные по одной машине ---
def get_vehicle_info(vehicle_name, vehicles_db):
    # Нормализуем входное имя
    norm_query = normalize_name(vehicle_name, vehicles_db[2])

    # Ищем по нормализованному имени
    for item in vehicles_db:
        if len(item) < 8:
            continue
        display_name = item[1]
        norm_db_name = normalize_name(display_name, item[2])

        if norm_db_name == norm_query:
            br_rb = item[4]['rb']
            real_br = RB_TO_BR.get(br_rb, None)
            type_rus = item[7][0][1] if item[7] and item[7][0] else "Неизвестно"
            return {
                'type': type_rus,
                'br': real_br,
                'country': item[2].title()
            }
    return None

# --- Функция: получить список по строке техники ---
def get_vehicles_info_list(vehicles_str, vehicles_db):
    names = [v.strip() for v in vehicles_str.split(',') if v.strip()]
    result = []
    for name in names:
        info = get_vehicle_info(name, vehicles_db)
        if info:
            result.append({
                'name': name,
                'type': info['type'],
                'br': info['br'],
                'country': info['country']
            })
        else:
            result.append({
                'name': name,
                'type': 'Неизвестно',
                'br': None,
                'country': 'Неизвестно'
            })
    return result

# --- Функция: определить тип боя ---
def classify_battle(info_list):
    if not info_list:
        return "Unknown"

    AIR_TYPES = {'Истребитель', 'Бомбардировщик', 'Ударный самолёт'}
    HELICOPTER_TYPES = {'Ударный вертолёт', 'Многоцелевой вертолёт'}
    GROUND_TYPES = {'Средний танк', 'Лёгкий танк', 'Тяжёлый танк', 'САУ', 'ЗСУ'}
    NON_AIR_TYPES = GROUND_TYPES | HELICOPTER_TYPES
    try:
        types = {v['type'] for v in info_list}
        countries = {v['country'] for v in info_list}
        max_br = max(v['br'] for v in info_list if v['br'] is not None)
        num_vehicles = len(info_list)
    except ValueError:
        return "Unknown"

    # Air AB: только самолёты, 2+, одна страна
    if types.issubset(AIR_TYPES) and num_vehicles >= 2 and len(countries) == 1:
        return "Air AB"

    # Air RB: один самолёт
    if types.issubset(AIR_TYPES) and num_vehicles == 1:
        return "Air RB"

    # Tank RB: есть наземная техника, не Air, BR < 10.7 или мало машин
    if types & NON_AIR_TYPES and not types.issubset(AIR_TYPES):
        if max_br < 10.7 or num_vehicles <= 2:
            return "Tank RB"

    # Tank SB: наземная техника, высокий BR
    if types & GROUND_TYPES and max_br >= 10.7 and num_vehicles <= 2:
        return "Tank SB"

    return "Unknown"

# --- ОСНОВНАЯ ФУНКЦИЯ ---
def analyze_battle(vehicles_str, vehicles_db=vehicles_rus):
    """
    Принимает строку с техникой (через запятую), возвращает:
        (тип_боя, максимальный_BR, страна_с_макс_BR)
    """
    info_list = get_vehicles_info_list(vehicles_str, vehicles_db)
    
    if not info_list:
        return "Unknown", None, "Неизвестно"

    # Находим запись с максимальным BR
    valid_vehicles = [v for v in info_list if v['br'] is not None]
    if not valid_vehicles:
        highest = info_list[0]
        max_br = None
        br_country = highest['country']
    else:
        highest = max(valid_vehicles, key=lambda x: x['br'])
        max_br = highest['br']
        br_country = highest['country']

    battle_type = classify_battle(info_list)
    
    return battle_type, max_br, br_country

battles = pd.read_excel(r'C:\Users\lutsevich\Downloads\Telegram Desktop\data.xlsx')

result_series = battles['vehicles'].apply(analyze_battle) # используем apply чтобы применить функцию к столбцу, на выхоже кортеж с индексом и тремя значениями
# преобразуем в df
new_columns = battles['vehicles'].apply(analyze_battle).apply(pd.Series)
new_columns.columns = ['battle_type', 'max_br', 'br_country']
# Добавляем новые столбцы к исходному df
battles_with_results = pd.concat([battles, new_columns], axis=1)

# Сохраняем в эксель
battles_with_results.to_excel('data_с_результатами.xlsx', index=False)