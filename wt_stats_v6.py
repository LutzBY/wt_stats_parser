import tkinter as tk
from tkinter import scrolledtext, messagebox
import pyperclip
import re
import pandas as pd
import sys
from datetime import timedelta, datetime
from threading import Thread
import keyboard
import pygetwindow as gw
import time
import json
from PIL import Image, ImageTk

# cd E:\PY\wt_stats_parser
# pyinstaller --onefile --windowed wt_stats_v6.py

import getpass # для определения текущего пользователя позже убрать
env = getpass.getuser()

# 0 Вводные
# 0.1 Куда сохранять эксель
xlsx_path = r"C:\Users\lutzb\Desktop\wt_stats\data.xlsx" if env == 'lutzb' else r"D:\py\wt_stats\data.xlsx"
# 0.2 Где лежит база техники
bd_path = r"E:\PY\wt_stats_parser\res\vehicles_rus.json" if env == 'lutzb' else r"D:\py\wt_stats\wt_stats_parser\res\vehicles_rus.json"
# 0.3 Параметры расположения окна tkinter
tkinter_geometry = (500, 300, 3965, 1050) if env == 'lutzb' else (500, 300, 1400, 725) # размер - ш, в, положение - ш, в (3520 + 1080 )
# 0.4 Где лежат флажки
res_loc = r"E:\PY\wt_stats_parser\res" if env == 'lutzb' else r'D:\py\wt_stats\wt_stats_parser\res'
# 0.5 Время запуска программы
session_start_time = datetime.now()
# 0.6 Датасет для SessionSummaryWindow
df_for_session = pd.DataFrame()

##### временная функция дампа (см строку 43)
def save_raw_report(text, file_path='report_dump.txt'):
    with open(file_path, 'a', encoding='utf-8') as f:
        f.write(f"\n{'='*50}\n")
        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"{'='*50}\n")
        f.write(text.strip() + '\n')
        f.write(f"{'-'*50}\n")

# 1 Функция парсинга результатов
def parse_battle_stats():
    imported_game_log = pyperclip.paste()
    if not imported_game_log.strip():
        print("❌ Буфер обмена пуст. Скопируй статистику боя и запусти скрипт снова.")
        return None
    # дополнить список репортов
    save_raw_report(imported_game_log)

    # --- Результат: Победа / Поражение ---
    result_match = re.search(r'(Победа|Поражение) в миссии', imported_game_log)
    result = result_match.group(1) if result_match else "Неизвестно"

    # --- Название миссии ---
    mission_match = re.search(r'миссии\s+"([^"]+)"', imported_game_log)
    mission = mission_match.group(1) if mission_match else "Неизвестно"

    # --- Итого: СЛ, СОИ (FRP), ОИ (RP) — только последнее вхождение ---
    total_matches = re.findall(r'Итого:\s*(\d+)\s*СЛ,\s*(\d+)\s*СОИ,\s*(\d+)\s*ОИ', imported_game_log)
    if not total_matches:
        print("❌ Не удалось найти ни одного вхождения 'Итого'.")
        return None

    # Берём ПОСЛЕДНЕЕ вхождение (финальные итоги)
    last_match = total_matches[-1]
    total_sl = int(last_match[0])   # Silver Lions
    total_frp = int(last_match[1])  # Free Research Points
    total_rp = int(last_match[2])   # Research Points

    # --- Очки миссии ---
    mission_points = re.findall(r'(\d+)\s*очк(?:о|а|ов)\s*миссии', imported_game_log)
    total_mission_points = sum(int(x) for x in mission_points)

    # --- Сессия ---
    session_match = re.search(r'Сессия:\s*([a-f0-9]+)', imported_game_log)
    session_id = session_match.group(1) if session_match else None
    if not session_id:
        print("❌ Не удалось найти session_id.")
        return None

    # --- Активность (%) ---
    activity_match = re.search(r'Активность:\s*(\d+)%', imported_game_log)
    activity_percent = int(activity_match.group(1)) if activity_match else None

    # --- Использованная техника ---
    vehicles_set = set()

    # Паттерн 1: "Время активности" — ищем текст до "Цифры + (ПА)"
    pattern_active = r'^\s*(.+?)\s+\d+\s*\+\s*$$ПА$$'
    # Паттерн 2: "Время игры" — ищем текст до "95% ... 4:51"
    pattern_game = r'^\s*(.+?)\s+\d+%.*?\d+:\d+'

    active_time_matches = re.findall(pattern_active, imported_game_log, re.MULTILINE)
    game_time_matches = re.findall(pattern_game, imported_game_log, re.MULTILINE)
    
    all_vehicles = active_time_matches + game_time_matches

    for v in all_vehicles:
        cleaned = re.sub(r'\s+', ' ', v.strip())
        # Исключаем ложные срабатывания (например, "Заработано", "Итого")
        if cleaned and not re.match(r'^[\[\]"]', cleaned) and len(cleaned) > 1:
            vehicles_set.add(cleaned)

    vehicles = ", ".join(sorted(vehicles_set)) if vehicles_set else "Неизвестно"

    # --- Время миссии ---
    mission_time_match = re.search(r'Время игры\s*(\d+:\d+)', imported_game_log)
    mission_time = mission_time_match.group(1) if mission_time_match else "Неизвестно"
    minutes, seconds = map(int, mission_time.split(':'))
    mission_time = timedelta(minutes=minutes, seconds=seconds)

    # --- Бустеры ---
    boosters_sl_match = re.search(r'Активные усилители на СЛ:[^.]*?Общий:\s*\+\s*(\d+)%СЛ', imported_game_log)
    boosters_rp_match = re.search(r'Активные усилители на ОИ:[^.]*?Общий:\s*\+\s*(\d+)%ОИ', imported_game_log)

    boosters_sl_percent = int(boosters_sl_match.group(1)) if boosters_sl_match else None
    boosters_rp_percent = int(boosters_rp_match.group(1)) if boosters_rp_match else None

    # --- Запуск анализатора по строке vehicles - Получение бр, типа боя и страны ---
    battle_type, max_br, br_country = analyzer.analyze_battle(vehicles)

    # --- Запуск анализатора по vehicles_set - Получение индекса был ли прем техника и сколько ---
    is_prem_veh_used = analyzer.is_prem_veh_used(vehicles_set)

    # --- Запуск анализатора для получения листа vehicles по бою ---
    analyzer.save_vehicle_stats(
        imported_game_log,
        vehicles_set,
        boosters_sl_percent,
        boosters_rp_percent,
        session_id,
        result,
        xlsx_path
    )

    return {
        'session_id': session_id,
        'vehicles': vehicles,
        'total_sl': total_sl,
        'total_frp': total_frp,
        'total_rp': total_rp,
        'total_mission_points': total_mission_points,
        'result': result,
        'mission': mission,
        'activity_percent': activity_percent,
        'mission_time': mission_time,
        'battle_type': battle_type,
        'max_br': max_br,
        'br_country': br_country,
        'boosters_sl_percent': boosters_sl_percent,
        'boosters_rp_percent': boosters_rp_percent,
        'is_prem_veh_used': is_prem_veh_used
    }

# 2 Функция сохранения в эксель
def save_to_excel(data, xlsx_path):
    global columns, df_for_session

    columns = [
        'session_id', 'vehicles', 'total_sl', 'total_frp', 'total_rp',
        'total_mission_points', 'result', 'mission', 'activity_percent', 
        'mission_time', 'battle_type', 'max_br', 'br_country', 
        'boosters_sl_percent', 'boosters_rp_percent', 'is_prem_veh_used'
    ]

    try:
        with pd.ExcelFile(xlsx_path, engine='openpyxl') as xls:
            # Пытаемся прочитать лист 'battles'
            if 'battles' in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name='battles', engine='openpyxl')
            else:
                df = pd.DataFrame(columns=columns)
    except (FileNotFoundError, ValueError):
        df = pd.DataFrame(columns=columns)

    # Удаляем строку с таким session_id, если есть
    df = df[df['session_id'] != data['session_id']]

    # Добавляем новую
    new_row = pd.DataFrame([data], columns=columns)
    df = pd.concat([df, new_row], ignore_index=True)

    # Дополняем второй датафрейм для finish_window
    df_for_session = pd.concat([df_for_session, new_row], ignore_index=True) 

    # Сохраняем обратно
    with pd.ExcelWriter(xlsx_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='battles', index=False)

    print(f"\n ✅ Обновлено: {data['session_id']}")

# 3 Работа с БД бр-ов и видов техники, возврат страны, бр и вида боя
class BattleAnalyzer:
    def __init__(self, bd_path):
        """
        Загружает и подготавливает базу техники.
        :param bd_path: путь к vehicles_rus.json
        """

        # 3.1 Таблицы замен
        self.RB_TO_BR = {
            0: 1.0, 1: 1.3, 2: 1.7, 3: 2.0, 4: 2.3, 5: 2.7, 6: 3.0, 7: 3.3,
            8: 3.7, 9: 4.0, 10: 4.3, 11: 4.7, 12: 5.0, 13: 5.3, 14: 5.7,
            15: 6.0, 16: 6.3, 17: 6.7, 18: 7.0, 19: 7.3, 20: 7.7, 21: 8.0,
            22: 8.3, 23: 8.7, 24: 9.0, 25: 9.3, 26: 9.7, 27: 10.0,
            28: 10.3, 29: 10.7, 30: 11.0, 31: 11.3, 32: 11.7, 33: 12.0,
            34: 12.3, 35: 12.7, 36: 13.0, 37: 13.3, 38: 13.7, 39: 14.0, 40: 14.3
        } # конвертация rb → BR

        self.COUNTRY_TO_RUSSIAN = {
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
        } # таблица замен для стран

        self.COUNTRY_SYMBOLS = {
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
        } # таблица замен для символов (вроде не нужно)

        self.COUNTRY_TO_FLAG_FILE = {
            'Usa': 'country_usa',
            'Germany': 'country_germany',
            'Ussr': 'country_ussr',
            'Britain': 'country_britain',
            'Japan': 'country_japan',
            'China': 'country_china',
            'Italy': 'country_italy',
            'France': 'country_france',
            'Sweden': 'country_sweden',
            'Israel': 'country_israel'
        } # таблица соответствия страна из базы - флаг

        self.HTML_REPLACEMENTS = {
            '&#039;': "'",  # апостроф
            '&amp;': '&',   # &
            '<': '<',    # <
            '>': '>',     # >
            '&quot;': '"' # ""
        } # таблица замен для HTML-символов

        self.SHITTY_SYMBOLS = [
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
            # '◘' # ◘SB-25J в игре именно sbj со значком
            # '␙' = кубок,␠ = лапка
        ] # таблица значков

        # 3.2 Загрузка базы и очистка ---

        with open(bd_path, encoding='UTF-8') as file:
            vehicles_rus_raw = json.load(file)

        self.vehicles_rus = []
        for item in vehicles_rus_raw:
            if len(item) < 8:
                continue
            # Копируем, чтобы не портить оригинал
            clean_item = item.copy()
            original_name = item[1]
            clean_name = self.normalize_name(original_name, item[2])
            clean_item[1] = clean_name  # заменяем на чистое имя

            self.vehicles_rus.append(clean_item)
        
    # 3.3 - Нормализация имен - заменяет HTML-символы, заменяет символы стран на текст
    def normalize_name(self, name, country_operator=None): 

        if not isinstance(name, str):
            name = str(name)

        # 1. Заменяем HTML-символы
        for html, text in self.HTML_REPLACEMENTS.items():
            name = name.replace(html, text)

        # 2. Заменяем эмодзи-флаги на текст (🇺🇸 → США)
        for symbol, country in self.COUNTRY_SYMBOLS.items():
            name = name.replace(symbol, country)

        # 3. Обработка ␗Имя → Имя (Страна)
        if any(symbol in name for symbol in self.SHITTY_SYMBOLS) and country_operator:
            clean_name = name
            for symbol in self.SHITTY_SYMBOLS:
                clean_name = clean_name.replace(symbol, '')
            clean_name = clean_name.strip()
            country_rus = self.COUNTRY_TO_RUSSIAN.get(country_operator.lower())
            if country_rus:
                name = f"{clean_name}({country_rus})"
            else:
                name = clean_name  # на всякий случай

        # 4. Финальная очистка
        name = re.sub(r'\s+', ' ', name.strip())

        return name
    
    # 3.4 Получить данные по одной машине ---
    def get_vehicle_info(self, vehicle_name, vehicles_db):
        # Нормализуем входное имя
        norm_query = self.normalize_name(vehicle_name)

        # Ищем по нормализованному имени
        for item in vehicles_db:
            if len(item) < 8:
                continue
            display_name = item[1]
            norm_db_name = self.normalize_name(display_name, item[2])

            if norm_db_name == norm_query:
                br_rb = item[4]['rb']
                real_br = self.RB_TO_BR.get(br_rb, None)
                type_rus = item[7][0][1] if item[7] and item[7][0] else "Неизвестно"
                return {
                    'type': type_rus,
                    'br': real_br,
                    'country': item[2].title()
                }
        return None
    
    # 3.5 Получить имя, тип, бр и страну по всей строке техники ---
    def get_vehicles_info_list(self, vehicles_str, vehicles_db):
        names = [v.strip() for v in vehicles_str.split(',') if v.strip()]
        result = []
        for name in names:
            info = self.get_vehicle_info(name, vehicles_db)
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
    
    # 3.6 определить тип боя ---
    def classify_battle(self, info_list):
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

        #### !!!!!!!!!Изменить
        # Tank RB: есть наземная техника, не Air, BR < 10.7 или мало машин
        if types & NON_AIR_TYPES and not types.issubset(AIR_TYPES):
            if max_br < 10.7 or num_vehicles <= 2:
                return "Tank RB"

        #### !!!!!!!!!Изменить
        # Tank SB: наземная техника, высокий BR
        if types & GROUND_TYPES and max_br >= 10.7 and num_vehicles <= 2:
            return "Tank SB"

        return "Unknown"
    
    # 3.7 Основная функция формирования battle_type, max_br, br_country
    def analyze_battle(self, vehicles_str):
        """
        Основной метод
        Принимает строку с техникой (через запятую), возвращает:
            (тип_боя, максимальный_BR, страна_с_макс_BR)
        """
        info_list = self.get_vehicles_info_list(vehicles_str, self.vehicles_rus)
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

        battle_type = self.classify_battle(info_list)
        
        return battle_type, max_br, br_country
    
    # 3.8 Получение флага на основе текста нации 
    def load_img(self, img_type, img_name, img_size):
        """
        Загружает картинку из папки res и возвращает PhotoImage.
        :param img_type: соответствует имени папки (flags или res)
        :param img_name: имя файла без расширения (rp или country_britain)
        :param size: размер (ширина, высота) в пикселях
        :return: ImageTk.PhotoImage
        """ 
        path = f'{res_loc}\\{img_type}\\{img_name}.png'
            
        # Конвертируем SVG → PNG (через PIL)
        try:
            # Открываем SVG как изображение
            img = Image.open(path)
            # Изменяем размер
            img = img.resize(img_size, Image.Resampling.LANCZOS)
            # Преобразуем в PhotoImage
            photo = ImageTk.PhotoImage(img)
            return photo
        except Exception as e:
            print(f"❌ Ошибка при загрузке {path}: {e}")
            return None
    
    # 3.9 Расчет средних из xlsx на основании данных миссии
    def get_averages_from_xlsx(self, battle_type, max_br, br_country):
        # Подгружаем эксель
        df = pd.read_excel(xlsx_path, engine='openpyxl', sheet_name='battles')
        
        # 3.9.1 Для поля "Тип+БР"
        # Создаем фильтрованный дф и получаем нужные поля
        filtered_df = df[(df['battle_type'] == battle_type) & (df['max_br'] == max_br)]
        if filtered_df.empty == False:
            avg_mp = int(filtered_df['total_mission_points'].mean()) # cannot convert float NaN to integer
            avg_sl = int(filtered_df['total_sl'].mean())
            avg_rp = int(filtered_df['total_rp'].mean())
            avg_act = int(filtered_df['activity_percent'].mean())
            avg_time = filtered_df['mission_time'].mean()
            td = pd.to_timedelta(avg_time, unit='D')
            hours = td.components.hours
            minutes = td.components.minutes
            seconds = td.components.seconds
            formatted_time = f"{hours}:{minutes:02d}:{seconds:02d}"
        else: 
           avg_mp = avg_sl = avg_rp = avg_act = avg_time = formatted_time = None

        # 3.9.2 Для поля "Нация"
        filtered_df = df[(df['battle_type'] == battle_type) & (df['br_country'] == br_country)]
        if filtered_df.empty == False:
            avg_mp_country = int(filtered_df['total_mission_points'].mean())
            avg_sl_country = int(filtered_df['total_sl'].mean())
            avg_rp_country = int(filtered_df['total_rp'].mean())
            avg_act_country = int(filtered_df['activity_percent'].mean())
            avg_time = filtered_df['mission_time'].mean()
            td = pd.to_timedelta(avg_time, unit='D')
            hours = td.components.hours
            minutes = td.components.minutes
            seconds = td.components.seconds
            formatted_time_country = f"{hours}:{minutes:02d}:{seconds:02d}"
        else: 
           avg_mp_country = avg_sl_country = avg_rp_country = avg_act_country = avg_time = formatted_time_country = None

        # 3.9.3 Для поля "По типу всего без бустеров"
        
        filtered_df = df[(df['battle_type'] == battle_type) ]
        avg_mp_no_boosters= int(filtered_df['total_mission_points'].mean())
        avg_sl_no_boosters = int((filtered_df['total_sl'] / (1 + filtered_df['boosters_sl_percent'].fillna(0) / 100)).mean())
        avg_rp_no_boosters = int((filtered_df['total_rp'] / (1 + filtered_df['boosters_rp_percent'].fillna(0) / 100)).mean())
        avg_act_no_boosters = int(filtered_df['activity_percent'].mean())
        avg_time = filtered_df['mission_time'].mean()
        td = pd.to_timedelta(avg_time, unit='D')
        hours = td.components.hours
        minutes = td.components.minutes
        seconds = td.components.seconds
        formatted_time_no_boosters = f"{hours}:{minutes:02d}:{seconds:02d}"

        return {
        'avg_mp': avg_mp,
        'avg_sl': avg_sl,
        'avg_rp': avg_rp,
        'avg_time': formatted_time,
        'avg_act': avg_act,
        'avg_mp_country' : avg_mp_country,
        'avg_sl_country': avg_sl_country,
        'avg_rp_country': avg_rp_country,
        'avg_act_country': avg_act_country,
        'formatted_time_country': formatted_time_country,
        'avg_mp_no_boosters' : avg_mp_no_boosters,
        'avg_sl_no_boosters' : avg_sl_no_boosters,
        'avg_rp_no_boosters' : avg_rp_no_boosters,
        'avg_act_no_boosters' : avg_act_no_boosters,
        'formatted_time_no_boosters' : formatted_time_no_boosters
        }
    
    # 3.10 Проверка использовалась ли прем техника в бою
    # Возвращает 0 если нет, 1 и более, если было 1 и более машин в списке
    def is_prem_veh_used(self, vehicles):
        is_prem = 0
        for row in self.vehicles_rus:
            name = row[1]
            value = row[5]
            if name in vehicles:
                is_prem += value
        return is_prem

    # 3.11 - Анализатор по технике
    def save_vehicle_stats(self, imported_game_log, vehicles_set, boosters_sl_percent, boosters_rp_percent, session_id, result, xlsx_path):
        
        # Создаем хранилку для доступа к результатам по технике
        global battle_data_vehicles
        battle_data_vehicles = None

        normalized_log = imported_game_log.replace('\r\n', '\n').replace('\r', '\n')
        def extract_block(text, keywords):
            """
            Извлекает блок, начиная со строки, содержащей все ключевые слова,
            и до первой пустой строки или конца текста.
            """
            lines = text.split('\n')
            in_block = False
            block_lines = []
            # Список возможных заголовков других блоков для раннего выхода (можно расширить)
            # Это поможет, если блоки не всегда отделены пустой строкой, но начинаются новый заголовок
            other_block_starters = [
                'Уничтожение авиации',
                'Уничтожение наземной техники', # Повтор для других вариантов
                'Помощь в уничтожении противника',
                'Критические повреждения противника',
                'Фатальные повреждения противника', # На случай, если это другой тип
                'Повреждения противника',
                'Захват зон',
                'Разведка противника',
                'Награды',
                'Время активности',
                'Время игры',
                'Награда за победу',
                'Награда за участие в миссии',
                'Бонус за мастерство',
                'Заработано:',
                'Активность:',
                'Повреждённая техника:',
                'Потраченных машин-дублёров:',
                'Автоматический ремонт',
                'Автоматическая закупка',
                'Исследуемая техника:',
                'Прогресс исследований:',
                'Сессия:',
                'Итого:'
                # Добавьте сюда другие, если нужно
            ]
            
            for i, line in enumerate(lines):
                stripped_line = line.strip()
                
                # --- Начало блока ---
                if not in_block and all(kw in stripped_line for kw in keywords):
                    in_block = True
                    continue # Пропускаем саму строку-заголовок, если не нужна
                
                # --- Внутри блока ---
                if in_block:
                    # --- Условие выхода ---
                    # 1. Конец текста
                    if not stripped_line:
                        break # Обычно пустая строка означает конец блока
                    
                    # 2. Начало другого блока
                    is_other_block_start = any(starter in stripped_line for starter in other_block_starters if starter != ' '.join(keywords))
                    # Также проверим, если сама строка является заголовком
                    if is_other_block_start and not all(kw in stripped_line for kw in keywords):
                        break

                    block_lines.append(stripped_line)

            result = '\n'.join(block_lines) if block_lines else None
            return result

        # --- Извлечение блоков ---
        kill_block   = extract_block(normalized_log, ['Уничтожение', 'наземной', 'техники'])
        kill_air_block   = extract_block(normalized_log, ['Уничтожение', 'авиации'])
        assist_block = extract_block(normalized_log, ['Помощь', 'уничтожении', 'противника'])
        crit_block    = extract_block(normalized_log, ['Критические', 'повреждения'])
        cap_block     = extract_block(normalized_log, ['Захват', 'зон'])
        game_block    = extract_block(normalized_log, ['Время', 'игры'])

        rows = []

        for vehicle in vehicles_set:
            total_sl = 0
            total_rp = 0
            total_mp = 0

            # Суммируем SL/RP/MP по всему логу
            pattern = rf'.*{re.escape(vehicle)}.*(?:СЛ|ОИ|очков миссии)'
            matches = re.findall(pattern, imported_game_log, re.IGNORECASE)

            for match in matches:
                sl_match = re.search(r'=\s*(\d+)\s*СЛ', match)
                rp_match = re.search(r'=\s*(\d+)\s*ОИ', match)
                mp_match = re.search(r'(\d+)\s*очк(?:о|а|ов)\s*миссии', match)

                if sl_match:
                    total_sl += int(sl_match.group(1))
                if rp_match:
                    total_rp += int(rp_match.group(1))
                if mp_match:
                    total_mp += int(mp_match.group(1))

            # --- Активность и время — только из "Время игры" ---
            activity_percent = None
            mission_time = None
            if game_block:
                for line in game_block.split('\n'):
                    if vehicle in line:
                        act_match = re.search(r'(\d+)%', line)
                        time_match = re.search(r'(\d{1,2}:\d{2}(?::\d{2})?)', line)
                        if act_match:
                            activity_percent = int(act_match.group(1))
                        if time_match:
                            mission_time = time_match.group(1)
                            minutes, seconds = map(int, mission_time.split(':'))
                            mission_time = timedelta(minutes=minutes, seconds=seconds)
                        break

            # --- Если данных нет — пропускаем? или пишем нули? ---
            if total_sl == 0 and total_rp == 0 and total_mp == 0 and activity_percent is None:
                print(f"⚠️ Нет данных по: {vehicle}")
                continue

            # --- Коррекция на бустеры ---
            sl_boost = boosters_sl_percent or 0
            rp_boost = boosters_rp_percent or 0
            corrected_sl = total_sl / (1 + sl_boost / 100) if sl_boost > 0 else total_sl
            corrected_rp = total_rp / (1 + rp_boost / 100) if rp_boost > 0 else total_rp

            # --- Премиум? ---
            premium = False
            for row in self.vehicles_rus:
                if row[1] == vehicle:
                    premium = row[5]
                    break

            # --- Подсчёт действий ---
            kills = len(re.findall(rf'{re.escape(vehicle)}', kill_block, re.IGNORECASE)) if kill_block else 0
            kills_air = len(re.findall(rf'{re.escape(vehicle)}', kill_air_block, re.IGNORECASE)) if kill_air_block else 0
            assists = len(re.findall(rf'{re.escape(vehicle)}', assist_block, re.IGNORECASE)) if assist_block else 0
            crits = len(re.findall(rf'{re.escape(vehicle)}', crit_block, re.IGNORECASE)) if crit_block else 0
            base_caps = len(re.findall(rf'{re.escape(vehicle)}.*?\d+%', cap_block, re.IGNORECASE)) if cap_block else 0

            rows.append({
                'battle_id': session_id,
                'result': result,
                'vehicle': vehicle,
                'premium': premium,
                'sl_corrected': round(corrected_sl),
                'rp_corrected': round(corrected_rp),
                'mp': total_mp,
                'activity_percent': activity_percent,
                'mission_time': mission_time,
                'kills': kills,
                'kills_air': kills_air,
                'assists': assists,
                'crits': crits,
                'base_caps': base_caps
            })
        # Заполняем хранилку
        battle_data_vehicles = rows

        # --- Запись в Excel ---
        if not rows:
            print("❌ Нечего записывать: ни одна машина не имеет данных.")
            return

        df_vehicles = pd.DataFrame(rows)

        try:
            # Пытаемся прочитать существующий лист vehicles
            try:
                existing_df = pd.read_excel(xlsx_path, sheet_name='vehicles', engine='openpyxl')
                # Фильтруем старые строки с этим battle_id
                existing_df = existing_df[existing_df['battle_id'] != session_id]
                # Объединяем
                combined_df = pd.concat([existing_df, df_vehicles], ignore_index=True)
            except (ValueError, KeyError) as e:
                # Листа нет или он пуст — просто используем новые данные
                combined_df = df_vehicles
                print(f'❌ Ошибка в save_vehicle_stats - {e}')

            # Записываем обратно
            with pd.ExcelWriter(xlsx_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                combined_df.to_excel(writer, sheet_name='vehicles', index=False)

        except Exception as e:
            print(f'❌ Ошибка в save_vehicle_stats - {e}')
    
# 4 Основное рабочее окно Tkinter
class WTApp:
    def __init__(self, root, tkinter_geometry):
        self.root = root
        self.root.title("WT Parser")
        root.geometry('%dx%d+%d+%d' % tkinter_geometry)
        self.root.resizable(True, True)
        self.root.attributes('-topmost', True)
        self.root.attributes('-alpha', 0.90)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing) # запуск окна SessionSummaryWindow по закрытию

        # Настройка сетки
        root.grid_rowconfigure(5, weight=1)  # растяжение для text_area
        root.grid_columnconfigure(0, weight=1)

        # --- 0. Заголовок: флаг + результат + миссия ---
        self.header_frame = tk.Frame(root)
        self.header_frame.grid(row=0, column=0, sticky='w', padx=10, pady=(10, 2))

        self.flag_label = tk.Label(
            self.header_frame,
            font=("Segoe UI", 11),
            anchor="w"
        )
        self.flag_label.pack(side='left')

        self.mission_label = tk.Label(
            self.header_frame,
            text="Последняя миссия: неизвестно",
            font=("Segoe UI", 11),
            fg="gray",
            anchor="w"
        )
        self.mission_label.pack(side='left', padx=(5, 0))

        # --- 1. Инфо-строка: тип, БР, бустеры ---
        self.info_label = tk.Label(
            root,
            text="",
            font=("Segoe UI", 9),
            fg="black",
            anchor="w"
        )
        self.info_label.grid(row=1, column=0, sticky='w', padx=10, pady=2)
        
        # --- 2. Техника ---
        self.vehicles_frame = tk.Frame(
            root, 
            bg=root.cget('bg'),
            )
        self.vehicles_frame.grid(row=2, column=0, sticky='w', padx=10, pady=2)

        self.vehicles_text = tk.Text(
            root,
            font=("Courier New", 10),
            fg="gray",
            wrap=tk.WORD,
            height=3,
            state="normal",
            bg=root.cget('bg'),
            borderwidth=0,
            highlightthickness=0
        )
        self.vehicles_text.grid(row=2, column=0, sticky='w', padx=10, pady=2)

        # --- 3. Таблица: текущие и средние значения ---
        self.stats_frame = tk.Frame(root, bd=1, relief="solid")
        self.stats_frame.grid(row=3, column=0, sticky='ew', padx=10, pady=5)

        # Настройка колонок
        for col in range(5):
            self.stats_frame.grid_columnconfigure(col, weight=1, uniform="col")

        # Заголовки (Row 0)
        headers = ['По типу боя:','🌐', '🐱', '💡', '⏲️', '🏃']
        for col, text in enumerate(headers):
            tk.Label(
                self.stats_frame,
                text=text,
                font=("Courier New", 9, "bold"),
                fg="black",
                anchor="center"
            ).grid(row=0, column=col, sticky='ew')

        # Создаём строки данных (без вложенных Frame!)
        self.current_row = self.create_stat_row(1, "Текущий бой")
        self.avg_type_br_row = self.create_stat_row(2, "AVG (БР)")
        self.avg_nation_row = self.create_stat_row(3, "AVG (нация)")
        self.avg_no_boosters_row = self.create_stat_row(4, "Всех, без бустеров")

        # --- 4. Кнопка ---
        self.button = tk.Button(
            root,
            text="📝 Записать",
            font=("Arial", 11),
            command=self.on_button_click
        )
        self.button.grid(row=4, column=0, pady=5)

        # --- 5. Лог (терминал) ---
        self.text_area = scrolledtext.ScrolledText(
            root,
            wrap=tk.WORD,
            font=("Consolas", 9),
            state='disabled',
            bg="white",
            fg="black",
            height=6
        )
        self.text_area.grid(row=5, column=0, padx=10, pady=(0, 10))

        sys.stdout = TextRedirector(self.text_area)

    # Столбцы таблички
    def create_stat_row(self, row, label_text):
        """Возвращает список из 5 Label, размещённых в stats_frame"""
        # Метка слева
        tk.Label(
            self.stats_frame,
            text=label_text,
            font=("Courier New", 8),
            fg="gray",
            width=12,
            anchor="w",
            wraplength='120'
        ).grid(row=row, column=0, sticky='w', padx=(0, 5))

        # Пять ячеек данных
        labels = []
        for col in range(1, 6):  # колонки 1–5
            lbl = tk.Label(
                self.stats_frame,
                text="—",
                font=("Courier New", 9, "bold"),
                anchor="center"
            )
            lbl.grid(row=row, column=col, sticky='ew')
            labels.append(lbl)
        return labels
    
    # Действия по кнопке "Записать" - наполнение заголовка и таблички, вызов save_to_excel
    def on_button_click(self):
        print("🔄 Обработка буфера обмена...")
        data = parse_battle_stats()
        if not data:
            print("❌ Обработка не удалась.")
            return

        # --- 0. Заголовок: флаг + миссия ---
        mission = data['mission']
        result = data['result']
        br_country = data['br_country']

        img_name = analyzer.COUNTRY_TO_FLAG_FILE.get(br_country, None)
        flag_image = analyzer.load_img('flags', img_name, img_size=(20, 14)) if img_name else None

        self.flag_label.config(image=flag_image)
        self.flag_label.image = flag_image

        self.mission_label.config(text=f"{result}: {mission}", fg="black")

        # --- 1. Инфо-строка: тип, БР, бустеры ---
        battle_type = data['battle_type']
        max_br = data['max_br']
        boosters_rp_percent = data['boosters_rp_percent']
        boosters_sl_percent = data['boosters_sl_percent']
        if boosters_rp_percent and boosters_sl_percent:
            boosters_percent_formatted = f'RP +{boosters_rp_percent}%, SL +{boosters_sl_percent}%'
        elif boosters_rp_percent and boosters_sl_percent is None:
            boosters_percent_formatted = f'RP +{boosters_rp_percent}%'
        elif boosters_sl_percent and boosters_rp_percent is None:
            boosters_percent_formatted = f'SL +{boosters_sl_percent}%'
        else:
            boosters_percent_formatted = 'Без бустеров'

        self.info_label.config(
            text=f"{battle_type} | BR {max_br} | {boosters_percent_formatted}"
        )
        
        # --- 2. Техника ---

        self.vehicles_text.delete(1.0, tk.END)
        
        # задаем конфиги
        self.vehicles_text.tag_configure("gray", foreground="gray")
        self.vehicles_text.tag_configure("orange", foreground="orange")
        self.vehicles_text.tag_configure("green", foreground="green")
        self.vehicles_text.tag_configure("orange", foreground="orange")
        self.vehicles_text.tag_configure("accent", underline=True)
        
        # Задаем условия по значкам
        top_sl = max(battle_data_vehicles, key=lambda x: x['sl_corrected'])
        top_rp = max(battle_data_vehicles, key=lambda x: x['rp_corrected'])
        top_mp = max(battle_data_vehicles, key=lambda x: x['mp'])
        # сумма убийств, ассистов, баз и т.д.
        top_usefulness = max(battle_data_vehicles, key=lambda x: x['kills_air'] + x['kills'] + x['assists'] + x['crits'] + x['base_caps'])
        
        for i, item in enumerate(battle_data_vehicles):
            vehicle = item['vehicle']
            
            # Задаем цвет
            if item['premium'] == 0:
                color = 'gray'
            elif item['premium'] == 1:
                color = 'orange'
            elif item['premium'] == 2:
                color = 'green'
            else:
                color = 'gold'
            
            # Задаем правила для значков
            # Если один собрал все критерии
            if item == top_sl == top_rp == top_mp == top_usefulness:
                self.vehicles_text.insert(tk.END, '🌟')
            else: 
                # Лучий по SL
                if top_sl == item:
                    self.vehicles_text.insert(tk.END, '🐱')
                # Лучий по RP
                if top_rp == item:
                    self.vehicles_text.insert(tk.END, '💡')
                # Лучий по MP
                if top_mp == item:
                    self.vehicles_text.insert(tk.END, '🌐')
                if top_usefulness == item:
                    self.vehicles_text.insert(tk.END, '💀')

            # Записываем тестовое имя с цветом и добавляем запятую
            self.vehicles_text.insert(tk.END, vehicle, color)
            if i < len(battle_data_vehicles) - 1:
                self.vehicles_text.insert(tk.END, " | ", "gray")
        

        # --- 3. Таблица значений ---
        mp = data['total_mission_points']
        sl = data['total_sl']
        rp = data['total_frp']
        time_str = str(data['mission_time'])
        act = data['activity_percent']

        self.current_row[0].config(text=mp)
        self.current_row[1].config(text=f"{sl:,}".replace(',', ' '))
        self.current_row[2].config(text=f"{rp:,}".replace(',', ' '))
        self.current_row[3].config(text=time_str)
        self.current_row[4].config(text=f"{act}%")

        # --- Расчет средних значений ---
        averages = analyzer.get_averages_from_xlsx(battle_type, max_br, br_country)

        avg_mp, avg_sl, avg_rp, avg_time, avg_act, avg_mp_country, avg_sl_country, avg_rp_country, avg_act_country, formatted_time_country, avg_mp_no_boosters, avg_sl_no_boosters, avg_rp_no_boosters, avg_act_no_boosters, formatted_time_no_boosters = averages.values()
        if avg_mp:
            self.avg_type_br_row[0].config(text=avg_mp)
            self.avg_type_br_row[1].config(text=f"{avg_sl:,}".replace(',', ' '))
            self.avg_type_br_row[2].config(text=f"{avg_rp:,}".replace(',', ' '))
            self.avg_type_br_row[3].config(text=avg_time)
            self.avg_type_br_row[4].config(text=f"{avg_act}%")
        if avg_mp_country:
            self.avg_nation_row[0].config(text=avg_mp_country)
            self.avg_nation_row[1].config(text=f"{avg_sl_country:,}".replace(',', ' '))
            self.avg_nation_row[2].config(text=f"{avg_rp_country:,}".replace(',', ' '))
            self.avg_nation_row[3].config(text=formatted_time_country)
            self.avg_nation_row[4].config(text=f"{avg_act_country}%")
        if avg_mp_no_boosters:
            self.avg_no_boosters_row[0].config(text=avg_mp_no_boosters)
            self.avg_no_boosters_row[1].config(text=f"{avg_sl_no_boosters:,}".replace(',', ' '))
            self.avg_no_boosters_row[2].config(text=f"{avg_rp_no_boosters:,}".replace(',', ' '))
            self.avg_no_boosters_row[3].config(text=formatted_time_no_boosters)
            self.avg_no_boosters_row[4].config(text=f"{avg_act_no_boosters}%")
        
        # --- Сохранение ---
        save_to_excel(data, xlsx_path)
        print("✅ Данные сохранены")
    
    # Создает наполнение и логику окна статистики
    def on_closing(self):
        if 'df_for_session' in globals() and not df_for_session.empty:

            # Длительность сессии
            session_end_time = datetime.now()
            session_total_time = session_end_time - session_start_time
            hours = int(session_total_time.total_seconds() // 3600)
            minutes = int((session_total_time.total_seconds() % 3600) // 60)
            session_total_time_str = f'{hours} ч, {minutes} мин'

            # Среднее время боя за сессию
            mission_avg_time = df_for_session['mission_time'].mean()
            td = pd.to_timedelta(mission_avg_time, unit='D')
            minutes_avg = td.components.minutes
            seconds_avg = td.components.seconds
            mission_avg_time_str = f"{minutes_avg:02d} мин, {seconds_avg:02d} сек"

            # Суммы по sl, rp, mp
            session_total_sl = f"{sum(df_for_session['total_sl']):_}".replace("_", " ")
            session_total_rp = f"{sum(df_for_session['total_frp']):_}".replace("_", " ")
            session_total_mp = f"{sum(df_for_session['total_mission_points']):_}".replace("_", " ")

            # Средние по sl, rp, mp
            session_average_sl = f"{int(df_for_session['total_sl'].mean()):_}".replace("_", " ")
            session_average_rp = f"{int(df_for_session['total_frp'].mean()):_}".replace("_", " ")
            session_average_mp = f"{int(df_for_session['total_mission_points'].mean()):_}".replace("_", " ")

            # Винрейт
            winrate = df_for_session['result'].value_counts()
            winrate = round(winrate.get('Победа', 1) / winrate.sum() * 100, 1)

            session_data = {
                'session_total_time': session_total_time_str,
                'battles_count': len(df_for_session),
                'winrate': winrate,
                'mission_avg_time': mission_avg_time_str,
                'session_total_sl': session_total_sl,
                'session_total_rp': session_total_rp,
                'session_total_mp': session_total_mp,
                'session_average_sl': session_average_sl,
                'session_average_rp': session_average_rp,
                'session_average_mp': session_average_mp
            }
            # Закрываем основное окно, открываем окно SessionSummaryWindow
            self.root.destroy()
            SessionSummaryWindow(tk.Tk(), session_data)
        else:
            self.root.destroy()

# 4.1 Окно статистики по окончанию игровой сессии
class SessionSummaryWindow:
    def __init__(self, parent, session_data):
        self.session_data = session_data
        self.window = parent
        self.window.title("Игровая сессия завершена")
        self.window.geometry('%dx%d+%d+%d' % tkinter_geometry)
        self.window.resizable(False, False)
        self.window.attributes('-topmost', True)
        self.window.attributes('-alpha', 0.75)
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        self.create_widgets()

    # Вид окошка
    def create_widgets(self):
        data = self.session_data

        text = f"""
Продлилась {data['session_total_time']}, боев - {data['battles_count']}, побед - {data['winrate']} %
Средняя продолжительность миссии - {data['mission_avg_time']}

Заработано всего:
🐱 {data['session_total_sl']} SL
💡 {data['session_total_rp']} RP
🌐 {data['session_total_mp']} MP

Заработано в среднем:
🐱 {data['session_average_sl']} SL
💡 {data['session_average_rp']} RP
🌐 {data['session_average_mp']} MP
        """.strip()

        label = tk.Label(self.window, text=text, font=("Consolas", 11), justify="left")
        label.pack(pady=20)

    # Действия по крестику
    def close(self):
        self.window.destroy()
        sys.exit()

# 5 Забираем текст из print() для размещения его в окне ткинтер
class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        if text.strip():  # чтобы не вставлять пустые строки
            self.widget.configure(state='normal')
            self.widget.insert(tk.END, text)
            self.widget.see(tk.END)
            self.widget.configure(state='disabled')
            self.widget.update_idletasks()  # обновление интерфейса

    def flush(self):
        pass  # требуется для совместимости с stdout

# 6 Листенер для проверки запускать ли логику парсера - Проверяет: нажат ли Ctrl+C и активно ли окно War Thunder.
def start_global_listener(app_instance):
    def is_wt_active():
        try:
            w = gw.getActiveWindow()
            if not w:
                return False
            title = w.title.lower()
            keywords = ['war thunder', 'wt', 'aces']
            return any(kw in title for kw in keywords)
        except:
            return False

    print("🟢 Перехват Ctrl+C активирован...")

    while True:
        # Проверяем, что оба нажаты
        if keyboard.is_pressed('ctrl') and keyboard.is_pressed('c'):
            if is_wt_active():
                print("\n✅ Ctrl+C в War Thunder — запускаем парсинг...")
                time.sleep(0.4)
                # Имитируем нажатие кнопки
                app_instance.on_button_click()
                # Ждём, пока клавиши отпущены
                while keyboard.is_pressed('c'):
                    time.sleep(0.1)
        time.sleep(0.1)  # не грузим CPU

# 7 === ЗАПУСК ===
if __name__ == "__main__":
    root = tk.Tk()
    app = WTApp(root, tkinter_geometry)
    analyzer = BattleAnalyzer(bd_path=bd_path)

    # Запускаем перехват в фоне, передаём экземпляр app
    listener_thread = Thread(target=start_global_listener, args=(app,), daemon=True)
    listener_thread.start()

    root.mainloop()
    