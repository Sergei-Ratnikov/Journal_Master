# cable_parser.py
"""
Парсер строки кабельного журнала.
Основан на явном разборе блоков "Откуда" и "Куда" по заданным правилам.

Входные данные: input_row — список, полученный из таблицы Word после обработки doc_utils.
Структура input_row:
    [0] = имя журнала (строка)
    [1] = список с номером кабеля, например ['8.0001']
    [2] = список с KKS кабеля, маркой, сечением и т.п.
    [3] = список с группой раскладки и классом безопасности
    [4..n-3] = блоки "Откуда" и "Куда" (каждый блок может состоять из 1 или 3 ячеек-списков)
    [-2] = список с длиной кабеля (может содержать звёздочки)
    [-1] = список с трассой (может быть многострочным)

Выходные данные: array_row — список из 20 полей (индексы 0..19), который позже дополняется датой, версией и т.д.
"""

import re
from kks_utils import cleanCyrFromLat, normalize_coordinates, extract_kks_from_list, parse_kks_room_and_equip
import config


# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (общие) ==========
# Эти функции используются для работы с вложенными списками и проверок.

def make_hashable(item):
    """Рекурсивно преобразует списки в кортежи для возможности хеширования.
       Нужна для is_subset_with_lists."""
    if isinstance(item, list):
        return tuple(make_hashable(x) for x in item)
    return item

def is_subset_with_lists(list1, list2):
    """Проверяет, является ли list1 подмножеством list2 с учётом вложенных списков.
       Используется в doc_utils для склейки разорванных строк."""
    set1 = {make_hashable(x) for x in list1}
    set2 = {make_hashable(x) for x in list2}
    return set1.issubset(set2)

def all_deep_empty(lst):
    """Рекурсивно проверяет, пусты ли все вложенные списки."""
    if not isinstance(lst, list):
        return False
    if not lst:
        return True
    return all(all_deep_empty(item) for item in lst)

def get_mismatch_indices(list1, list2):
    """Возвращает индексы элементов, которые различаются в двух списках.
       Списки должны быть одинаковой длины."""
    if len(list1) != len(list2):
        raise ValueError("Списки должны быть одинаковой длины")
    mismatch_indices = []
    for i, (a, b) in enumerate(zip(list1, list2)):
        if a != b:
            mismatch_indices.append(i)
    return mismatch_indices

def is_number(s):
    """Проверяет, является ли строка числом (целым или дробным, с необязательным знаком + или -).
       ВНИМАНИЕ: здесь сначала удаляется ведущий '+', затем проверяется число.
       Это может привести к тому, что строка '+0' станет '0' и пройдёт проверку.
       """
    if not s:
        return False
    # Убираем ведущий + (только первый символ)
    s = s.lstrip('+')
    # Проверяем, что строка состоит из опционального знака, цифр и опциональной дробной части
    return re.match(r'^[+-]?\d+(?:[.,]\d+)?$', s) is not None

def parse_coordinates_from_list(lst):
    """
    Извлекает из списка (одномерного) три координаты X,Y,Z.
    Координаты могут быть:
      - тремя отдельными элементами списка, идущими подряд,
      - одной строкой, содержащей три числа, разделённые пробелами и/или запятыми,
        при этом внутри числа может быть десятичная запятая (например, '123,45').
    Примеры: '123 345 -34,1', '123, 345, -34,1', '123.5 456,2 78.9'
    Возвращает:
        coords: список из трёх строк (координаты в исходном виде, с десятичной запятой, если была)
        filtered: список оставшихся элементов (без извлечённых координат)
    """
    coords = ['', '', '']

    # --- 1. Поиск трёх чисел в одной строке ---
    for idx, elem in enumerate(lst):
        if isinstance(elem, str):
            # Регулярное выражение ищет числа, включая:
            #   - опциональный знак + или -
            #   - одну или более цифр
            #   - опциональную дробную часть, где разделитель может быть точкой или запятой
            #   - числа могут быть разделены пробелами, запятыми, точками с запятой и т.д.
            # При этом дробная часть не теряется.
            pattern = r'[+-]?\d+(?:[.,]\d+)?'
            matches = re.findall(pattern, elem)
            # Если нашли три и более числа, берём первые три
            if len(matches) >= 3:
                coords = matches[:3]
                # Удаляем этот элемент из списка
                filtered = [elem for i, elem in enumerate(lst) if i != idx]
                return coords, filtered

    # --- 2. Если не нашли в одной строке, ищем три числа подряд в разных элементах ---
    number_positions = []
    for i, elem in enumerate(lst):
        # Для проверки: временно заменяем десятичную запятую на точку
        test_elem = str(elem).replace(',', '.') if isinstance(elem, str) else elem
        if is_number(test_elem):
            number_positions.append(i)
    # Проверяем, есть ли три числа, идущие подряд (разница индексов 1)
    if len(number_positions) >= 3 and number_positions[2] - number_positions[0] == 2:
        coords = [lst[number_positions[0]], lst[number_positions[1]], lst[number_positions[2]]]
        # Удаляем эти три элемента из списка
        filtered = [elem for i, elem in enumerate(lst) if i not in number_positions[:3]]
        return coords, filtered

    # --- 3. Если ничего не нашли, возвращаем пустые координаты и исходный список ---
    return ['', '', ''], lst

def parse_kks_block(block_cells, building_bounds):
    """
    Разбирает блок "Откуда" или "Куда", который может быть представлен:
    - одним списком (одна ячейка) — случаи 4.1, 4.2
    - тремя списками (три ячейки) — случай 4.3
    Возвращает:
        rooms: список уникальных KKS помещений (типа 00UKS10R032)
        equip: список уникальных KKS оборудования (типа 00CFN54)
        coords: список из трёх координат (X,Y,Z) как строки (пустые, если нет)
    """
    # Если block_cells — это список списков (т.е. несколько ячеек)
    if isinstance(block_cells, list) and len(block_cells) > 0 and isinstance(block_cells[0], list):
        # Если внутри всего одна ячейка, то разворачиваем её в плоский список
        if len(block_cells) == 1:
            block_cells = block_cells[0]
        else:
            # Случай с тремя отдельными списками (каждый список — одна ячейка)
            # Собираем все KKS из всех списков, координаты берём из каждого списка последовательно
            all_rooms = set()
            all_equip = set()
            coords = ['', '', '']
            for i, cell in enumerate(block_cells[:3]):  # обрабатываем первые три ячейки (больше не нужно)
                if not isinstance(cell, list):
                    cell = [cell]
                # Ищем координату в этой ячейке: первое попавшееся число
                coord_value = ''
                for elem in cell:
                    if is_number(elem):
                        coord_value = elem
                        break
                if i < 3:
                    coords[i] = coord_value
                # Извлекаем KKS (помещения и оборудование) из всех элементов ячейки
                for elem in cell:
                    if config.regular_KKS_room.search(elem):
                        all_rooms.add(elem)
                    elif config.regular_KKS_equipment.search(elem):
                        all_equip.add(elem)
            return list(all_rooms), list(all_equip), coords

    # Один список (одна ячейка) — случай 4.1 или 4.2
    if isinstance(block_cells, list):
        # Сначала пытаемся извлечь координаты (если три числа подряд)
        coords, filtered = parse_coordinates_from_list(block_cells)
        # Из оставшихся элементов извлекаем KKS
        rooms = set()
        equip = set()
        for elem in filtered:
            if config.regular_KKS_room.search(elem):
                rooms.add(elem)
            elif config.regular_KKS_equipment.search(elem):
                equip.add(elem)
        # Если после удаления координат не осталось KKS, возможно координаты были в конце,
        # а KKS в начале — тогда пробуем найти KKS среди всех элементов, исключая только что найденные координаты
        if not rooms and not equip:
            for elem in block_cells:
                if elem not in coords:
                    if config.regular_KKS_room.search(elem):
                        rooms.add(elem)
                    elif config.regular_KKS_equipment.search(elem):
                        equip.add(elem)
        return list(rooms), list(equip), coords

    # Если ничего не подошло (например, пустой блок)
    return [], [], ['', '', '']


def row_parser(input_row, building_bounds):
    """
    Основной парсер строки кабельного журнала.
    Возвращает список из 20 строковых полей.
    """
    array_row = [''] * 20

    # ===== 1. Базовая проверка длины =====
    if len(input_row) < 7:
        return []

    # ===== 2. Очистка от символов '>' в начале строк =====
    # Некоторые журналы используют '>' для обозначения вложенности.
    cleaned_row = []
    for cell in input_row:
        if isinstance(cell, list):
            cleaned_cell = []
            for text in cell:
                if isinstance(text, str):
                    text = text.lstrip('>').strip()
                    if text:
                        cleaned_cell.append(text)
                else:
                    cleaned_cell.append(text)
            cleaned_row.append(cleaned_cell)
        else:
            cleaned_row.append(cell)
    input_row = cleaned_row

    # ===== 3. Извлечение простых полей (индексы 0-3) =====
    # 0. Журнал
    array_row[0] = cleanCyrFromLat(input_row[0]).strip()
    # 1. Номер кабеля
    if len(input_row) > 1 and input_row[1]:
        array_row[1] = input_row[1][0].replace(',', '.').strip()
    # 2. ККС кабеля
    if len(input_row) > 2 and input_row[2]:
        array_row[2] = cleanCyrFromLat(input_row[2][0]).strip()
    # 3. Группа раскладки (может содержать несколько чисел через запятую)
    if len(input_row) > 3 and input_row[3]:
        array_row[3] = input_row[3][0].strip()

    # ===== 4. Поиск марки и сечения =====
    # Марка и сечение могут находиться в любой из ячеек, начиная с третьей.
    for cell in input_row[2:]:
        if cell:
            for line in cell:
                if not line:
                    continue
                # Марка кабеля
                if not array_row[4]:
                    line_clean = line.replace('/', '').strip().split()[0]
                    if config.regular_cableMark.search(line_clean):
                        array_row[4] = line_clean
                    else:
                        for mark in config.arrayCableMarks:
                            if mark in line:
                                array_row[4] = line
                                break
                # Сечение (например, 2x2x0.35)
                if not array_row[5]:
                    match = config.regular_cableSection.search(line)
                    if match:
                        array_row[5] = match.group().replace('.', ',').replace('х', 'x').replace('×', 'x')

    # ===== 5. Поиск длины кабеля =====
    # Длина находится в одной из последних трёх ячеек (чаще всего предпоследняя).
    # Может содержать звёздочки, которые нужно удалить.
    length_idx = -1
    for idx in [-1, -2, -3]:
        if abs(idx) <= len(input_row) and input_row[idx]:
            length_raw = input_row[idx][0].replace(',', '.')
            number_pattern = r'\d*\.?\d+'
            # Удаляем звёздочки в начале или конце
            match_start = re.match(rf'^\*{{1,3}}({number_pattern})$', length_raw)
            if match_start:
                length_raw = match_start.group(1)
            else:
                match_end = re.match(rf'^({number_pattern})\*{{1,3}}$', length_raw)
                if match_end:
                    length_raw = match_end.group(1)
            if length_raw and length_raw.replace('.', '').replace('-', '').isdigit():
                array_row[7] = length_raw
                length_idx = len(input_row) + idx if idx < 0 else idx
                break

    # ===== 6. Трасса =====
    # Последняя ячейка (или несколько ячеек) — описание трассы.
    if input_row[-1]:
        trace = ', '.join(input_row[-1])
        trace = trace.replace(';', ',').replace('  ', ' ').replace(',,', ',')
        if trace and trace[-1] == ',':
            trace = trace[:-1]
        if trace and len(trace) > 2:
            array_row[8] = trace

    # ========== 7. РАЗБОР БЛОКОВ "ОТКУДА" И "КУДА" ==========
    # Определяем диапазон ячеек, которые относятся к блокам.
    # Начало: всегда индекс 4 (после группы).
    # Конец: до ячейки с длиной (если найдена), иначе до предпоследней ячейки (перед трассой).
    start_idx = 4
    end_idx = length_idx if length_idx != -1 else len(input_row) - 2

    if end_idx <= start_idx:
        # Нет блоков — возвращаем то, что уже набрали
        return array_row

    # Перебираем возможные размеры блока "Откуда" (1 или 3 ячейки).
    # Выбираем вариант, который даёт осмысленные KKS или координаты в обоих блоках.
    # ПРОБЛЕМА: этот простой перебор может ошибаться, когда и при size_from=1,
    # и при size_from=3 оба блока имеют какие-то данные. Сейчас берётся первый подходящий,
    # что может быть неправильно (например, для шести ячеек size_from=1 тоже подходит,
    # но даёт неверное разбиение). Более правильным было бы перебрать все возможные size_from
    # (от 1 до total-1) и выбрать вариант с максимальной «полезностью» (количество найденных
    # координат и KKS). Ниже реализован старый вариант (только 1 и 3).
    # Для полной надёжности нужно заменить на перебор с оценкой.
    best_rooms_from = []
    best_equip_from = []
    best_coords_from = []
    best_rooms_to = []
    best_equip_to = []
    best_coords_to = []
    
    possible_sizes = [1, 3]
    for size_from in possible_sizes:
        if size_from > (end_idx - start_idx):
            continue
        from_cells = input_row[start_idx:start_idx+size_from]
        to_cells = input_row[start_idx+size_from:end_idx]
        if not to_cells:
            continue
        rooms_f, equip_f, coords_f = parse_kks_block(from_cells, building_bounds)
        rooms_t, equip_t, coords_t = parse_kks_block(to_cells, building_bounds)
        # Условие: хотя бы один непустой элемент в каждом блоке
        if (rooms_f or equip_f or any(coords_f)) and (rooms_t or equip_t or any(coords_t)):
            best_rooms_from, best_equip_from, best_coords_from = rooms_f, equip_f, coords_f
            best_rooms_to, best_equip_to, best_coords_to = rooms_t, equip_t, coords_t
            break  # берём первый подходящий — это может быть причиной ошибок

    # Если ни один вариант не подошёл, считаем, что весь диапазон — это "Откуда", а "Куда" пусто.
    if not best_rooms_from and not best_equip_from and not any(best_coords_from):
        from_cells = input_row[start_idx:end_idx]
        best_rooms_from, best_equip_from, best_coords_from = parse_kks_block(from_cells, building_bounds)

    # Сохраняем результаты в переменные для дальнейшего использования
    list_of_KKS_start = best_rooms_from
    list_of_KKS_equipment_start = best_equip_from
    list_of_axis_start = best_coords_from if len(best_coords_from) == 3 else []

    list_of_KKS_end = best_rooms_to
    list_of_KKS_equipment_end = best_equip_to
    list_of_axis_end = best_coords_to if len(best_coords_to) == 3 else []

    # Удаляем дубликаты (могут быть повторения из‑за нескольких ячеек)
    list_of_KKS_start = list(set(list_of_KKS_start))
    list_of_KKS_end = list(set(list_of_KKS_end))
    list_of_KKS_equipment_start = list(set(list_of_KKS_equipment_start))
    list_of_KKS_equipment_end = list(set(list_of_KKS_equipment_end))

    # ========== 8. ЗАПОЛНЕНИЕ ПОЛЕЙ ==========
    # Поле 19: резервирование (если в ячейках 3 или 4 есть слово "резерв")
    for cell in input_row[3:5]:
        if cell:
            for line in cell:
                if re.search(r'резерв', line, re.IGNORECASE):
                    array_row[19] = 'Резерв'
                    break

    # Парсим помещения и оборудование через общую функцию (из kks_utils).
    # Эта функция пытается выделить помещение (regular_KKS_room) и оборудование.
    rs, es = parse_kks_room_and_equip(list_of_KKS_start, building_bounds)
    rr, ee = parse_kks_room_and_equip(list_of_KKS_end, building_bounds)

    # Если оборудование не найдено, используем отдельно собранные списки
    if not es and list_of_KKS_equipment_start:
        es = list_of_KKS_equipment_start[0]
    if not ee and list_of_KKS_equipment_end:
        ee = list_of_KKS_equipment_end[0]

    # Корректировка для случая, когда кабель идёт в одном помещении (только одна сторона имеет координаты)
    if not rr and rs:
        if len(list_of_axis_end) == 3:
            for kks in list_of_KKS_start:
                if kks != rs and config.regular_KKS_room.search(kks):
                    rr = kks
                    break
            if not rr:
                rr = rs
    if not rs and rr:
        if len(list_of_axis_start) == 3:
            for kks in list_of_KKS_end:
                if kks != rr and config.regular_KKS_room.search(kks):
                    rs = kks
                    break
            if not rs:
                rs = rr

    # Заполняем итоговые поля
    array_row[9] = rs if rs else ''
    array_row[10] = es if es else ''
    array_row[14] = rr if rr else ''
    array_row[15] = ee if ee else ''

    # ===== 9. НОРМАЛИЗАЦИЯ КООРДИНАТ =====
    # Для начала (откуда)
    if len(list_of_axis_start) == 3:
        norm_start = normalize_coordinates(rs, list_of_axis_start, building_bounds)
        if norm_start:
            # Выходные координаты в Excel должны быть с запятой как разделителем
            array_row[11], array_row[12], array_row[13] = (
                norm_start[0].replace('.', ','),
                norm_start[1].replace('.', ','),
                norm_start[2].replace('.', ',')
            )
        else:
            # Если нормализация не удалась, оставляем как есть (убираем плюс)
            array_row[11] = list_of_axis_start[0].replace('+', '').strip().replace('.', ',')
            array_row[12] = list_of_axis_start[1].replace('+', '').strip().replace('.', ',')
            array_row[13] = list_of_axis_start[2].replace('+', '').strip().replace('.', ',')

    # Для конца (куда)
    if len(list_of_axis_end) == 3:
        norm_end = normalize_coordinates(rr, list_of_axis_end, building_bounds)
        if norm_end:
            array_row[16], array_row[17], array_row[18] = (
                norm_end[0].replace('.', ','),
                norm_end[1].replace('.', ','),
                norm_end[2].replace('.', ',')
            )
        else:
            array_row[16] = list_of_axis_end[0].replace('+', '').strip().replace('.', ',')
            array_row[17] = list_of_axis_end[1].replace('+', '').strip().replace('.', ',')
            array_row[18] = list_of_axis_end[2].replace('+', '').strip().replace('.', ',')

    # ===== 10. КОПИРОВАНИЕ KKS ПРИ НАЛИЧИИ ТОЛЬКО КООРДИНАТ =====
    # Если есть координаты куда, но нет помещения/оборудования куда — копируем из начала
    if (array_row[16] or array_row[17] or array_row[18]) and not (array_row[14] or array_row[15]):
        if array_row[9] or array_row[10]:
            array_row[14] = array_row[9]
            array_row[15] = array_row[10]
    # Аналогично, если есть координаты откуда, но нет помещения/оборудования откуда — копируем из конца
    if (array_row[11] or array_row[12] or array_row[13]) and not (array_row[9] or array_row[10]):
        if array_row[14] or array_row[15]:
            array_row[9] = array_row[14]
            array_row[10] = array_row[15]

    return array_row