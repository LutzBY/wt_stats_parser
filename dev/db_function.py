import json
import re
import pandas as pd

# --- Пути и данные ---
# path = r'C:\Users\lutsevich\Desktop\py\wt_stats\wt_stats_parser\vehicles_rus.json'
path = r"C:\Users\lutsevich\Desktop\py\wt_stats\wt_stats_parser\res\vehicles_rus.json"

# --- Таблица rb → BR ---
RB_TO_BR = {
    0: 1.0, 1: 1.3, 2: 1.7, 3: 2.0, 4: 2.3, 5: 2.7, 6: 3.0, 7: 3.3,
    8: 3.7, 9: 4.0, 10: 4.3, 11: 4.7, 12: 5.0, 13: 5.3, 14: 5.7,
    15: 6.0, 16: 6.3, 17: 6.7, 18: 7.0, 19: 7.3, 20: 7.7, 21: 8.0,
    22: 8.3, 23: 8.7, 24: 9.0, 25: 9.3, 26: 9.7, 27: 10.0,
    28: 10.3, 29: 10.7, 30: 11.0, 31: 11.3, 32: 11.7, 33: 12.0,
    34: 12.3, 35: 12.7, 36: 13.0, 37: 13.3, 38: 13.7, 39: 14.0, 40: 14.3
}

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

# Таблица значков
SHITTY_SYMBOLS = [
    '▃', 
    '␗', 
    '▄',
    '▀', 
    '◔', 
    '▅',
    "▂", # ▂МК-II "Матильда"
    '◄', # ◄CL-13A Mk.5
    '◗', # ◗Fokker D.XXI
    '◡', # ◡Kfir C.10
    '◊', # ◊Lim-5P
    '◌', # ◌Mirage IIIS C.70
    '', # P-51D-20-NA
    '◘' # ◘SB-25J
    # '␙' = кубок,␠ = лапка
] 

# --- Нормализация ---
def normalize_name(name, country_code=None):
    """Нормализует имя техники при загрузке базы"""
    if not isinstance(name, str):
        return str(name)

    # 1. HTML-сущности
    for html, text in HTML_REPLACEMENTS.items():
        name = name.replace(html, text)

    # 2. Эмодзи-флаги
    for symbol, country in COUNTRY_SYMBOLS.items():
        name = name.replace(symbol, country)

    # 3. Значки вроде ␗Имя → Имя (Страна)
    if any(symbol in name for symbol in SHITTY_SYMBOLS) and country_code:
        clean_name = re.sub('|'.join(re.escape(s) for s in SHITTY_SYMBOLS), '', name)
        clean_name = re.sub(r'\s+', ' ', clean_name.strip())
        country_rus = COUNTRY_TO_RUSSIAN.get(country_code.lower(), "")
        if country_rus:
            name = f"{clean_name} ({country_rus})"
        else:
            name = clean_name

    # 4. Финальная очистка
    name = re.sub(r'\s+', ' ', name.strip())
    return name

# --- Загрузка и очистка базы ---
with open(path, encoding='UTF-8') as file:
    vehicles_rus_raw = json.load(file)

vehicles_rus = {}
for item in vehicles_rus_raw:
    if len(item) < 8:
        continue
    country_code = item[2].lower()
    original_name = item[1]
    clean_name = normalize_name(original_name, country_code)
    br_rb = item[4]['rb'] if item[4] else None
    real_br = RB_TO_BR.get(br_rb)
    vehicle_type = item[7][0][1] if item[7] and item[7][0] else "Неизвестно"

    vehicles_rus[clean_name] = {
        'type': vehicle_type,
        'br': real_br,
        'country': item[2].title()
    }

# --- Основная функция ---
def analyze_battle(vehicles_str):
    """
    Анализирует строку техники и возвращает:
        (battle_type, max_br, br_country)
    """
    if not vehicles_str or not isinstance(vehicles_str, str):
        return "Unknown", None, "Неизвестно"

    names = [v.strip() for v in vehicles_str.split(',') if v.strip()]
    info_list = []

    for name in names:
        info = vehicles_rus.get(name)
        if info:
            info_list.append({
                'name': name,
                'type': info['type'],
                'br': info['br'],
                'country': info['country']
            })
        else:
            info_list.append({
                'name': name,
                'type': 'Неизвестно',
                'br': None,
                'country': 'Неизвестно'
            })

    if not info_list:
        return "Unknown", None, "Неизвестно"

    # Определение типа боя
    AIR_TYPES = {'Истребитель', 'Бомбардировщик', 'Ударный самолёт'}
    GROUND_TYPES = {'Средний танк', 'Лёгкий танк', 'Тяжёлый танк', 'САУ', 'ЗСУ'}
    NON_AIR_TYPES = GROUND_TYPES | {'Ударный вертолёт', 'Многоцелевой вертолёт'}

    types = {v['type'] for v in info_list}
    countries = {v['country'] for v in info_list}
    valid_brs = [v['br'] for v in info_list if v['br'] is not None]
    max_br = max(valid_brs) if valid_brs else None
    num_vehicles = len(info_list)

    if not max_br:
        return "Unknown", None, "Неизвестно"

    # Air AB
    if types.issubset(AIR_TYPES) and num_vehicles >= 2 and len(countries) == 1:
        battle_type = "Air AB"
    # Air RB
    elif types.issubset(AIR_TYPES) and num_vehicles == 1:
        battle_type = "Air RB"
    # Tank RB
    elif types & NON_AIR_TYPES and not types.issubset(AIR_TYPES) and max_br < 10.7:
        battle_type = "Tank RB"
    # Tank SB
    elif types & GROUND_TYPES and max_br >= 10.7 and num_vehicles <= 2:
        battle_type = "Tank SB"
    else:
        battle_type = "Unknown"

    # Страна с макс BR
    highest = max(info_list, key=lambda x: x['br'] if x['br'] is not None else -1)
    br_country = highest['country']

    return battle_type, max_br, br_country

# --- Применение к Excel ---
battles = pd.read_excel(r'C:\Users\lutsevich\Downloads\Telegram Desktop\data.xlsx')
result_series = battles['vehicles'].apply(analyze_battle) # используем apply чтобы применить функцию к столбцу, на выхоже кортеж с индексом и тремя значениями
# преобразуем в df
new_columns = battles['vehicles'].apply(analyze_battle).apply(pd.Series)
new_columns.columns = ['battle_type', 'max_br', 'br_country']
# Добавляем новые столбцы к исходному df
battles_with_results = pd.concat([battles, new_columns], axis=1)

# Сохраняем в эксель
battles_with_results.to_excel('data_с_результатами.xlsx', index=False)