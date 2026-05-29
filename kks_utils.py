# kks_utils.py
'''
Все функции для работы с KKS (кодировка оборудования и помещений)
и координатами (парсинг, нормализация, проверка границ зданий)
'''

import re
import json
from pathlib import Path
from openpyxl import load_workbook
import config


# ----- базовые утилиты для KKS -----

def cleanCyrFromLat(lineIn):
    '''
    замена русских букв на латинские
    '''
    try:
        return str(lineIn).replace('А', 'A').replace('В', 'B').replace('С', 'C').replace('Е', 'E').replace('Н', 'H').replace('К', 'K').replace('М', 'M').replace('О', 'O').replace('Р', 'P').replace('Т', 'T').replace('Х', 'X').replace(' ', '').replace(',', '.')
    except Exception as e:
        raise

def parse_kks_room_and_equip(list_of_KKS, building_bounds):
    """
    Разделяет список KKS на помещение и оборудование.
    
    Args:
        list_of_KKS: список строк с KKS
        building_bounds: словарь с границами зданий из KKS_building_bounds.json
    
    Returns:
        tuple: (KKS_room, KKS_equipment) — помещение и оборудование
    """
    KKS_room = ''
    KKS_equipment = ''
    all_buildings = set(building_bounds.keys()) if building_bounds else set()
    
    if list_of_KKS:
        list_of_KKS = list(set(list_of_KKS))
        
        # Сначала ищем помещение
        for kks in list_of_KKS:
            if config.regular_KKS_room.search(kks):
                building_match = config.regular_KKS_building.search(kks)
                if building_match:
                    building = building_match.group().strip()
                    if building in all_buildings:
                        KKS_room = kks
                        break
            elif config.regular_KKS_building.search(kks) and not KKS_room:
                building = kks.strip()
                if building in all_buildings:
                    KKS_room = building
        
        # Затем ищем оборудование (отличное от помещения)
        for kks in list_of_KKS:
            if config.regular_KKS_equipment.search(kks) and config.regular_KKS_equipment.search(kks).group().strip() != KKS_room:
                KKS_equipment = config.regular_KKS_equipment.search(kks).group().strip()
                break
    
    return KKS_room, KKS_equipment

# ----- функции для работы с координатами и границами зданий -----

def parse_route_to_bounds(route_string):
    """
    Извлекает из строки Route координаты квадрата и возвращает расширенные границы.
    Используется при создании JSON-файла границ зданий.
    
    Формат строки: "14E; 20N" или "18E; 12N"
    Возвращает словарь с границами x_min, x_max, y_min, y_max.
    """
    pattern = re.compile(r'(\d{1,2})\s*([EN])')
    matches = pattern.findall(route_string.upper())
    
    if len(matches) < 2:
        return {
            'is_valid': False,
            'original': route_string,
            'x_min': None,
            'x_max': None,
            'y_min': None,
            'y_max': None
        }
    
    e_val = None
    n_val = None
    
    for num, axis in matches:
        value = int(num) * 100
        if axis == 'E':
            e_val = value
        elif axis == 'N':
            n_val = value
    
    if e_val is None or n_val is None:
        return {
            'is_valid': False,
            'original': route_string,
            'x_min': None,
            'x_max': None,
            'y_min': None,
            'y_max': None
        }
    
    # Расширяем диапазон на ±100 метров
    return {
        'is_valid': True,
        'original': route_string,
        'x_min': n_val,
        'x_max': n_val + 100,
        'y_min': e_val,
        'y_max': e_val + 100
    }

def export_kks_to_json(excel_path, json_path='KKS_building_bounds.json'):
    """
    Читает KKS.xlsx и сохраняет building_bounds в JSON.
    Эту функцию достаточно запустить один раз при обновлении справочника зданий.
    
    Args:
        excel_path: путь к файлу KKS.xlsx
        json_path: путь для сохранения JSON-файла
    """
    if not Path(excel_path).exists():
        print(f"❌ Файл не найден: {excel_path}")
        return
    
    building_bounds = {}
    
    wb = load_workbook(excel_path, data_only=True)
    sheet = wb.active
    
    row_count = 0
    saved_count = 0
    
    for row in range(2, sheet.max_row + 1):
        kks = str(sheet.cell(row=row, column=1).value).strip()
        route = str(sheet.cell(row=row, column=4).value).strip()
        row_count += 1
        
        if not kks or not config.regular_KKS_building.search(kks):
            continue
        
        kks_clean = cleanCyrFromLat(kks)
        bounds = parse_route_to_bounds(route)
        building_bounds[kks_clean] = bounds
        saved_count += 1
    
    # Сохраняем в JSON
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(building_bounds, f, ensure_ascii=False, indent=2)
    
    print(f"\n📊 Статистика:")
    print(f"   Всего строк в Excel: {row_count}")
    print(f"   Сохранено зданий: {saved_count}")
    print(f"   Пропущено (без KKS здания): {row_count - saved_count}")
    print(f"\n✅ Файл сохранён: {json_path}")

def parse_coordinate_pair(x_str, y_str):
    """
    Парсит пару координат X, Y из строк, поддерживая разные форматы.
    
    Возвращает:
        tuple: (x_val, y_val, scale) или (None, None, 1) при ошибке
        scale показывает множитель (1 или 1000) для нормализации
    """
    x_clean = x_str.strip().replace(',', '.')
    y_clean = y_str.strip().replace(',', '.')
    
    # Формат: 4001100-4002200, 547700-549200 (большие координаты)
    try:
        x = float(x_clean)
        y = float(y_clean)
        if (4001100 < x < 4002200) and (547700 < y < 549200):
            return (x - 4000000), (y - 547000), 1
    except:
        pass
    
    # Формат: 1100-2200, 700-2200 (нормальные координаты)
    try:
        x = float(x_clean)
        y = float(y_clean)
        if (1100 < x < 2200) and (700 < y < 2200):
            return x, y, 1
    except:
        pass
    
    # Формат: 1100000-2200000, 700000-2200000 (координаты с масштабом 1000)
    try:
        x = float(x_clean.replace('.', ''))
        y = float(y_clean.replace('.', ''))
        if (1100000 < x < 2200000) and (700000 < y < 2200000):
            return x / 1000.0, y / 1000.0, 1000
    except:
        pass
    
    return None, None, 1

def check_and_swap_axes(x, y, bounds, tolerance=100):
    """
    Проверяет, не перепутаны ли оси X и Y, используя границы здания как эталон.
    
    Логика:
        1. Если здание найдено и границы валидны:
           - Если x_min < y_min → ожидаем, что x < y. Если x > y → меняем местами
           - Если x_min > y_min → ожидаем, что x > y. Если x < y → меняем местами
        2. Если здание не найдено или границы невалидны:
           - По умолчанию ожидаем, что x < y. Если x > y → меняем местами
    
    Args:
        x, y: координаты для проверки
        bounds: словарь с границами здания (x_min, x_max, y_min, y_max)
        tolerance: допуск отклонения от границ (не используется в этой версии, 
                   оставлен для совместимости)
    
    Returns:
        tuple: (x, y) — возможно, исправленные координаты
    """
    
    # Случай 1: здание найдено и границы валидны
    if bounds and bounds.get('is_valid'):
        x_min = bounds['x_min']
        y_min = bounds['y_min']
        
        # Определяем ожидаемое соотношение осей по эталону здания
        if x_min < y_min:
            # В этом здании X должен быть меньше Y
            if x > y:
                # print(f"Обнаружены перепутанные координаты: X({x}) > Y({y}), а должно быть X < Y (по эталону здания). Меняем местами.")
                return y, x
        elif x_min > y_min:
            # В этом здании X должен быть больше Y
            if x < y:
                # print(f"Обнаружены перепутанные координаты: X({x}) < Y({y}), а должно быть X > Y (по эталону здания). Меняем местами.")
                return y, x
        # Если x_min == y_min — не меняем
        
        return x, y
    
    # Случай 2: здание не найдено или границы невалидны
    # По умолчанию ожидаем, что X < Y
    if x > y:
        # print(f"Эвристика: X({x}) > Y({y}) — меняем местами (здание не найдено или нет границ)")
        return y, x
    
    return x, y

def format_coordinate(value):
    """
    Форматирует координату для вывода в Excel:
    - округляет до 3 знаков
    - убирает лишние нули
    - целые числа выводит без десятичной части
    """
    if value is None:
        return ''
    rounded = round(value, 3)
    if rounded.is_integer():
        return str(int(rounded))
    return str(rounded).rstrip('0').rstrip('.')

def normalize_coordinates(kks_room, coords, building_bounds):
    """
    Нормализует координаты кабеля: парсит, масштабирует, проверяет и меняет оси при необходимости.
    
    Args:
        kks_room: KKS помещения (для определения здания и его границ)
        coords: список из трёх строк [x, y, z]
        building_bounds: словарь с границами зданий
    
    Returns:
        list: [x_str, y_str, z_str] — отформатированные координаты или None
    """
    if not coords or len(coords) != 3:
        return None
    
    x_raw, y_raw, z_raw = coords[0].strip(), coords[1].strip(), coords[2].strip()
    
    # Определяем здание по KKS помещения
    building_kks = None
    if kks_room and building_bounds:
        match = config.regular_KKS_building.search(kks_room)
        if match:
            building_kks = match.group()
            if building_kks not in building_bounds:
                building_kks = kks_room if kks_room in building_bounds else None
        else:
            if kks_room in building_bounds:
                building_kks = kks_room
    
    bounds = building_bounds.get(building_kks) if building_kks else None
    
    # Парсим координаты
    x_val, y_val, scale = parse_coordinate_pair(x_raw, y_raw)
    if x_val is None or y_val is None:
        return None
    
    # Парсим Z-координату
    try:
        z_val = float(z_raw.replace(',', '.'))
        if scale == 1000:
            z_val = z_val / 1000.0
    except:
        z_val = 0.0
    
    # Проверяем и меняем оси при необходимости
    x_final, y_final = check_and_swap_axes(x_val, y_val, bounds)
    
    return [format_coordinate(x_final), format_coordinate(y_final), format_coordinate(z_val)]