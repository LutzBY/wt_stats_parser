import tkinter as tk
from tkinter import scrolledtext
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
# pyinstaller --onefile --windowed wt_stats_v3.py

import getpass # для определения текущего пользователя позже убрать
env = getpass.getuser()

# 0 Вводные
# 0.1 Куда сохранять эксель
xlsx_path = r"C:\Users\lutzb\Desktop\wt_stats\data.xlsx" if env == 'lutzb' else r"D:\data.xlsx"
# 0.2 Где лежит база техники
bd_path = r"E:\PY\wt_stats_parser\res\vehicles_rus.json" if env == 'lutzb' else r"C:\Users\lutsevich\Desktop\py\wt_stats\wt_stats_parser\res\vehicles_rus.json"
# 0.3 Параметры расположения окна tkinter
tkinter_geometry = (400, 350, 4065, 1000) if env == 'lutzb' else (400, 350, 1500, 675) # размер - ш, в, положение - ш, в (3520 + 1080 )
# 0.4 Где лежат флажки
res_loc = r"E:\PY\wt_stats_parser\res" if env == 'lutzb' else r'C:\Users\lutsevich\Desktop\py\wt_stats\wt_stats_parser\res'

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
    ########################################### save_raw_report(imported_game_log)

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
    vehicles = set()

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
        if cleaned and not re.match(r'^[0-9\[\]"]', cleaned) and len(cleaned) > 1:
            vehicles.add(cleaned)

    vehicles = ", ".join(sorted(vehicles)) if vehicles else "Неизвестно"

    # --- Запуск анализатора по строке vehicles ---
    battle_type, max_br, br_country = analyzer.analyze_battle(vehicles)

    # --- Время миссии ---
    mission_time_match = re.search(r'Время игры\s*(\d+:\d+)', imported_game_log)
    mission_time = mission_time_match.group(1) if mission_time_match else "Неизвестно"
    minutes, seconds = map(int, mission_time.split(':'))
    mission_time = timedelta(minutes=minutes, seconds=seconds)

    # --- Бустеры ---
    boosters_sl_match = re.search(r'Активные усилители на СЛ:[^.]*?Общий:\s*\+\s*(\d+)%СЛ', imported_game_log)
    boosters_rp_match = re.search(r'Активные усилители на ОИ:[^.]*?Общий:\s*\+\s*(\d+)%ОИ', imported_game_log)

    boosters_sl_percent = int(boosters_sl_match.group(1)) if boosters_sl_match else None
    booster_rp_percent = int(boosters_rp_match.group(1)) if boosters_rp_match else None

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
        'booster_rp_percent': booster_rp_percent
    }

# 2 Функция сохранения в эксель
def save_to_excel(data, xlsx_path):
    
    columns = [
        'session_id', 'vehicles', 'total_sl', 'total_frp', 'total_rp',
        'total_mission_points', 'result', 'mission', 'activity_percent', 
        'mission_time', 'battle_type', 'max_br', 'br_country', 
        'boosters_sl_percent', 'booster_rp_percent'
    ]

    try:
        df = pd.read_excel(xlsx_path, engine='openpyxl')
    except (FileNotFoundError, ValueError):
        df = pd.DataFrame(columns=columns)

    # Удаляем строку с таким session_id, если есть
    df = df[df['session_id'] != data['session_id']]

    # Добавляем новую
    new_row = pd.DataFrame([data], columns=columns)
    df = pd.concat([df, new_row], ignore_index=True)

    # Сохраняем
    df.to_excel(xlsx_path, index=False, engine='openpyxl')
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
            '◘' # ◘SB-25J
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
        df = pd.read_excel(xlsx_path, engine='openpyxl')
        # Создаем фильтрованный дф и получаем нужные поля
        filtered_df = df[(df['battle_type'] == battle_type) & (df['max_br'] == max_br)]
        avg_mp = int(filtered_df['total_mission_points'].mean())
        avg_sl = int(filtered_df['total_sl'].mean())
        avg_rp = int(filtered_df['total_rp'].mean())
        avg_act= int(filtered_df['activity_percent'].mean())
        avg_time =  timedelta(filtered_df['mission_time'].mean())
        
        return {
        'avg_mp': avg_mp,
        'avg_sl': avg_sl,
        'avg_rp': avg_rp,
        'avg_act': avg_act,
        'avg_time': avg_time
        }

# 4 Окно Tkinter
class WTApp:
    def __init__(self, root, tkinter_geometry):
        self.root = root
        self.root.title("WT Parser")
        root.geometry('%dx%d+%d+%d' % (tkinter_geometry))
        self.root.resizable(True, True)
        self.root.attributes('-topmost', True) # поверх
        self.root.attributes('-alpha', 0.75) # прозрачность
        
        # Метка: последняя миссия
        self.last_mission_label = tk.Label(
            root,
            text="Последняя миссия: неизвестно",
            font=("Segoe UI", 11),
            fg="gray",
            anchor="w",
            justify="left"
        )
        self.last_mission_label.pack(pady=(10, 5), padx=10, fill='x')
        
        # Метка: результат посл. боя
        self.last_mission_earnings_label = tk.Label(
            root,
            text="-",
            font=("Segoe UI", 14),
            fg="gray",
            anchor="w",
            justify="left"
        )
        self.last_mission_earnings_label.pack(pady=(2, 1), padx=10, fill='x')

        # Метка: статистика по фильтру
        self.filtered_averages_label = tk.Label(
            root,
            text="-",
            font=("Segoe UI", 11),
            fg="gray",
            anchor="w",
            justify="left"
        )
        self.filtered_averages_label.pack(pady=(2, 1), padx=10, fill='x')

        # Кнопка
        self.button = tk.Button(
            root,
            text="📝 Записать",
            font=("Arial", 12),
            command=self.on_button_click
        )
        self.button.pack(pady=3)

        # Текстовое поле с выводом
        self.text_area = scrolledtext.ScrolledText(
            root,
            wrap=tk.WORD,
            font=("Consolas", 10),
            state='disabled',
            bg="white",
            fg="black",
            padx=10,
            pady=10
        )
        self.text_area.pack(expand=True, fill='both', padx=10, pady=1)

        # Перенаправление print в текстовое поле
        sys.stdout = TextRedirector(self.text_area)

    def on_button_click(self):
        self.text_area.configure(state='normal')
        self.text_area.delete(1.0, tk.END)
        self.text_area.configure(state='disabled')
        print("🔄 Обработка буфера обмена...")
        
        data = parse_battle_stats()
        if data:
            print("\n📋 Извлечено:")

            # Обновляем заголовок информацией из миссии
            mission = data['mission']
            result = data['result']
            br_country = data['br_country']
            # и получаем флаг из словаря
            img_name = analyzer.COUNTRY_TO_FLAG_FILE.get(br_country, None)
            if img_name:
                flag_image = analyzer.load_img('flags', img_name, img_size=(20, 14))
            else:
                flag_image = None

            self.last_mission_label.config(
                text=f"{result}: {mission}",
                image=flag_image, 
                compound='left',
                fg="black"
            ) # Подставляем новый текст
            self.last_mission_label.image = flag_image # сохраняем флажок чтобы ткинтер его не удалил после выполнения
            
            # Обновляем строку результатов
            mission_points = data['total_mission_points']
            earnings_sl = data['total_sl']
            earnings_frp = data['total_frp']
            mission_time = data['mission_time']
            activity_percent = data['activity_percent']
            
            self.last_mission_earnings_label.config(
                text=f"🌐 {mission_points} 🐱 {earnings_sl} 💡{earnings_frp} ⏲️ {mission_time} 🏃 {activity_percent}%",
                compound='left',
                fg="black"
            )

            # Обновляем строку статистики
            battle_type = data['battle_type']
            max_br = data['max_br']
            br_country = data['br_country']

            # Получем средние из экселя
            averages = analyzer.get_averages_from_xlsx(battle_type, max_br, br_country)
            
            self.filtered_averages_label.config(
                text=f"""Средние значения для {battle_type} на БР {max_br}:
                \n🌐 {averages['avg_mp']} 🐱 {averages['avg_sl']} 💡{averages['avg_rp']} ⏲️ {averages['avg_time']} 🏃 {averages['avg_act']}%
                """,
                compound='left',
                fg="black"
            )

            # Выводим в окошко WORD распаршенные строки
            for k, v in data.items():
                print(f"\n{k}: {v}")
            
            # Вызываем запись в эксель
            save_to_excel(data, xlsx_path)

        else:
            print("❌ Обработка не удалась.")

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