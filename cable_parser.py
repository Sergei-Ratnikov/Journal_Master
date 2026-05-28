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



В парсер передается список следующего содержания:
0.	Строка с KKS журнала, например 'AKU.0110.07UBG.0.AE.MB0001-EMB0001'
1.	Список, состоящий из одной строки с номером кабеля, например ['7.0001']
2.	Список, содержащий KKS кабеля, марку кабеля, сечение, например ['00ANE01-2401', 'Кабель контрольный / Control cable', 'KППГЭнг(A)-HF-T', 'KPPGEng(A)-HF-T', '5x2.5']
3.	Список, содержащий группу раскладки и класс безопасности, например ['3', '4']
4.	Блок информации “Откуда” о начале кабеля.

    4.1	может содержать один список, содержащий в себе KKS здания, KKS оборудования, числовые координаты x, y, z. Координаты могут быть 
        в начале списка или в конце. KKS здания и KKS оборудования могут неоднократно повторяться в рамках одного списка.
        Например ['07UBG13R013', '00BCE07', 'Комплектное распредустройство 10 кВ. Шкаф', '07UBG13R013', '00BCE07', '1802.7', '1550.3', '3.3']
        Например ['1781.2', '1758.6', '13.6', '00CMM27', 'АПТС ОНЕГА', 'CEH ONEGA', '00UKS22R017']
        Кроме того, все три координаты или только одна координата z может быть не числом, а '-'
        Например '1802.7', '1550.3', '-'   или , '-', '-', '-'

    4.2	может содержать один список, содержащий в себе KKS здания, KKS оборудования, вместо координат содержать только один знак '-'
        Например ['00UKS10R036', '00CYE36СR001-К02', 'Входной/выходной модуль', 'Input/output module', '-']. 
        Знак '-' может быть в начале списка или в конце. KKS здания и KKS оборудования могут неоднократно повторяться в рамках одного списка.

    4.3	может содержать три списка, содержащих в себе одинаковые себе KKS здания и KKS оборудования 
        и при этом каждый из трех списков содержит одну координату x, y или z. Координаты могут быть в начале списка или в конце. 
        KKS здания и KKS оборудования могут неоднократно повторяться в рамках одного списка.
        Например
        ['00UKS22R017', '00CYE36', 'Контроллер пожарной сигнализации', 'Fire alarm controller', '1779.1'], 
        ['00UKS22R017', '00CYE36', 'Контроллер пожарной сигнализации', 'Fire alarm controller', '1756.3'], 
        ['00UKS22R017', '00CYE36', 'Контроллер пожарной сигнализации', 'Fire alarm controller', '13.9']

5.	Блок информации “Куда” о конце кабеля. Структура аналогична блоку “Откуда”, блоки могут не совпадать. Например, Откуда может содержать 1 список, а Куда – 3 списка.
6.	Список, содержащий длину кабеля. Строка с положительным числовым значением или '-' или пусто или содержит помимо чисел звездочки, которые нужно удалить.
7.	Трасса

"""

import re
from kks_utils import cleanCyrFromLat, normalize_coordinates, parse_kks_room_and_equip
import config

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
    Извлекает из списка три координаты X,Y,Z.
    
    Алгоритм (по приоритету):
        1. Ищет строку с тремя числами/прочерками в НАЧАЛЕ списка (первые 3 элемента)
        2. Ищет строку с тремя числами/прочерками в КОНЦЕ списка (последние 3 элемента)
        3. Ищет три числа/прочерка ПОДРЯД в НАЧАЛЕ списка (первые 6 элементов)
        4. Ищет три числа/прочерка ПОДРЯД в КОНЦЕ списка (последние 6 элементов)
        5. Ищет три числа/прочерка в ЛЮБОЙ строке списка
    
    Координаты могут быть:
      - тремя отдельными элементами списка, идущими подряд,
      - одной строкой, содержащей три числа, разделённые пробелами и/или запятыми,
      - элементы, равные '-', интерпретируются как отсутствие координаты (пустая строка).
    
    Args:
        lst: список строк (одна ячейка таблицы)
    
    Returns:
        list: [X, Y, Z] где отсутствующие координаты заменены на пустые строки
    """
    coords = ['', '', '']
    
    def extract_from_string(s):
        """Извлекает три числа/прочерка из строки"""
        # Разбиваем строку на части по пробелам
        parts = s.split()
        # Если нет пробелов, пробуем разбить по запятым
        if len(parts) == 1 and ',' in s:
            parts = s.split(',')
        
        # Очищаем части от лишних пробелов
        parts = [p.strip() for p in parts if p.strip()]
        
        # Ищем три подряд идущих элемента, которые являются либо числом, либо '-'
        for i in range(len(parts) - 2):
            is_valid = True
            for j in range(3):
                p = parts[i + j]
                if p == '-':
                    continue
                test_p = p.replace(',', '.')
                if not re.match(r'^[+-]?\d+(?:\.\d+)?$', test_p):
                    is_valid = False
                    break
            
            if is_valid:
                result = ['', '', '']
                for j in range(3):
                    p = parts[i + j]
                    if p == '-':
                        result[j] = ''
                    else:
                        result[j] = p.lstrip('+')
                return result
        return None
    
    def extract_from_sublist(sublist):
        """Извлекает три числа/прочерка из подсписка (отдельные элементы)"""
        for i in range(len(sublist) - 2):
            candidates = sublist[i:i+3]
            
            # Проверяем, что каждый элемент либо число, либо '-'
            valid = True
            for val in candidates:
                if val == '-':
                    continue
                test_val = str(val).replace(',', '.')
                if not is_number(test_val):
                    valid = False
                    break
            
            if valid:
                result = ['', '', '']
                for j, val in enumerate(candidates):
                    if val == '-':
                        result[j] = ''
                    else:
                        result[j] = str(val).lstrip('+')
                return result
        return None
    
    # =============================================================
    # 1. ПОИСК СТРОКИ С КООРДИНАТАМИ В НАЧАЛЕ СПИСКА
    # =============================================================
    # Проверяем первые 3 элемента (если они строки)
    for i in range(min(3, len(lst))):
        elem = lst[i]
        if isinstance(elem, str) and (' ' in elem or ',' in elem):
            result = extract_from_string(elem)
            if result:
                return result
    
    # =============================================================
    # 2. ПОИСК СТРОКИ С КООРДИНАТАМИ В КОНЦЕ СПИСКА
    # =============================================================
    # Проверяем последние 3 элемента (если они строки)
    for i in range(-1, -min(4, len(lst)+1), -1):
        elem = lst[i]
        if isinstance(elem, str) and (' ' in elem or ',' in elem):
            result = extract_from_string(elem)
            if result:
                return result
    
    # =============================================================
    # 3. ПОИСК ТРЁХ ЧИСЕЛ ПОДРЯД В НАЧАЛЕ СПИСКА
    # =============================================================
    start_slice = lst[:6]  # берём первые 6 элементов
    result = extract_from_sublist(start_slice)
    if result:
        return result
    
    # =============================================================
    # 4. ПОИСК ТРЁХ ЧИСЕЛ ПОДРЯД В КОНЦЕ СПИСКА
    # =============================================================
    end_slice = lst[-6:] if len(lst) >= 6 else lst[:]
    result = extract_from_sublist(end_slice)
    if result:
        return result
    
    # =============================================================
    # 5. ПОИСК В ЛЮБОЙ СТРОКЕ СПИСКА (запасной вариант)
    # =============================================================
    for elem in lst:
        if isinstance(elem, str):
            result = extract_from_string(elem)
            if result:
                return result
    
    # =============================================================
    # 6. ПОИСК ОТДЕЛЬНОГО ЧИСЛА (например, только Z)
    # =============================================================
    numbers = []
    for val in lst:
        if val == '-':
            numbers.append('')
        else:
            test_val = str(val).replace(',', '.')
            if is_number(test_val):
                numbers.append(str(val).lstrip('+'))
            else:
                numbers.append(None)
    
    # Берём первые три числа/прочерка подряд (игнорируя текст)
    result = []
    for val in numbers:
        if val is not None:
            result.append(val)
            if len(result) == 3:
                return result
    
    # =============================================================
    # 7. НИЧЕГО НЕ НАШЛИ
    # =============================================================
    return ['', '', '']

def parse_kks_block(block_cells, building_bounds):
    """
    Разбирает блок "Откуда" или "Куда", который может быть представлен в разных форматах.
    
    Блок может быть:
        1. ОДНИМ СПИСКОМ - случаи 4.1 и 4.2.
           Например: ['00UKS22R017', '00CYE36', 'Контроллер...', '1779.1']
           В таком списке могут быть:
               - KKS помещения (например, '00UKS22R017')
               - KKS оборудования (например, '00CYE36')
               - текстовые описания (на русском и английском)
               - координаты X, Y, Z (три числа, идущие ПОДРЯД в списке)
               - или знак '-' вместо координат
        
        2. ТРЕМЯ СПИСКАМИ - случай 4.3.
           Например: [
               ['00UKS22R017', '00CYE36', '...', '1779.1'],
               ['00UKS22R017', '00CYE36', '...', '1756.3'],
               ['00UKS22R017', '00CYE36', '...', '13.9']
           ]
           В этом случае каждая ячейка содержит одну координату (X, Y или Z),
           а KKS могут повторяться в каждой ячейке.
    
    Args:
        block_cells: список ячеек (каждая ячейка - это список строк).
                    Может содержать одну ячейку или несколько.
        building_bounds: словарь с границами зданий (не используется напрямую,
                        но передаётся для единообразия с другими функциями).
    
    Returns:
        tuple: (rooms, equip, coords)
            rooms: список уникальных KKS помещений (например, ['00UKS22R017'])
            equip: список уникальных KKS оборудования (например, ['00CYE36'])
            coords: список из трёх координат [X, Y, Z] (пустые строки, если нет)
    """
    
    # =========================================================================
    # БЛОК А: ПРОВЕРКА - ЯВЛЯЕТСЯ ЛИ БЛОК СПИСКОМ СПИСКОВ (НЕСКОЛЬКО ЯЧЕЕК)
    # =========================================================================
    # Входные данные всегда являются списком. Но каждый элемент этого списка
    # может быть либо строкой, либо списком строк.
    # block_cells[0] - это первая ячейка. Если она является списком, значит у нас
    # несколько ячеек (каждая ячейка - свой список).
    
    if isinstance(block_cells, list) and len(block_cells) > 0 and isinstance(block_cells[0], list):
        # isinstance() — это встроенная функция Python, которая проверяет, принадлежит ли объект указанному типу (классу) или нескольким типам
        
        # ---------------------------------------------------------
        # СЛУЧАЙ 1: ОДНА ЯЧЕЙКА, НО ОНА ОБЁРНУТА В СПИСОК
        # ---------------------------------------------------------
        # Иногда при формировании данных одна ячейка может быть представлена как
        # список, содержащий один элемент-список. Например:
        #   block_cells = [ ['00UKS22R017', '00CYE36', '...', '1779.1'] ]
        # В этом случае нужно "развернуть" блок, взяв внутренний список.
        
        if len(block_cells) == 1:
            block_cells = block_cells[0]
            # Теперь block_cells - это плоский список, и мы перейдём к обработке (БЛОК Б)
        

        # ---------------------------------------------------------
        # СЛУЧАЙ 2: ТРИ ОТДЕЛЬНЫХ СПИСКА (три ячейки)
        # ---------------------------------------------------------
        else:
            # Это случай 4.3: три ячейки, каждая содержит одну координату
            # и повторяющиеся KKS.
            
            # =========================================================
            # ШАГ 1: ПРОВЕРКА ЦЕЛОСТНОСТИ БЛОКА
            # =========================================================
            # Убеждаемся, что все три ячейки относятся к одному блоку.
            # Для этого сравниваем KKS помещений и оборудования во всех трёх ячейках.
            # Они должны быть одинаковыми (или отсутствовать) во всех трёх строках.
            
            # Собираем уникальные KKS из каждой ячейки
            rooms_per_cell = []
            equip_per_cell = []
            
            for cell in block_cells[:3]:
                cell_rooms = set()
                cell_equip = set()
                if isinstance(cell, list):
                    for elem in cell:
                        if config.regular_KKS_room.search(elem):
                            cell_rooms.add(elem)
                        elif config.regular_KKS_building.search(elem):
                            cell_rooms.add(elem)  # здание считаем помещением
                        elif config.regular_KKS_equipment.search(elem):
                            cell_equip.add(elem)
                rooms_per_cell.append(cell_rooms)
                equip_per_cell.append(cell_equip)
            
            # Проверяем, что KKS помещений одинаковы во всех ячейках
            # (или пустые во всех)
            rooms_consistent = True
            first_rooms = rooms_per_cell[0] if rooms_per_cell else set()
            for rooms_set in rooms_per_cell[1:]:
                if rooms_set != first_rooms:
                    rooms_consistent = False
                    break
            
            # Проверяем, что KKS оборудования одинаковы во всех ячейках
            equip_consistent = True
            first_equip = equip_per_cell[0] if equip_per_cell else set()
            for equip_set in equip_per_cell[1:]:
                if equip_set != first_equip:
                    equip_consistent = False
                    break
            
            # Если KKS не совпадают во всех ячейках, это означает,
            # что в блок попали данные из разных мест.
            # В этом случае возвращаем пустые результаты (блок будет обработан по-другому).
            if not rooms_consistent or not equip_consistent:
                # Проблема: смешанные данные из "Откуда" и "Куда"
                # Возвращаем пустые значения, чтобы вышестоящий код попробовал другое разбиение
                return [], [], ['', '', '']
            
            # =========================================================
            # ШАГ 2: СБОР KKS И КООРДИНАТ
            # =========================================================
            all_rooms = set()
            all_equip = set()
            coords = ['', '', '']
            
            for i, cell in enumerate(block_cells[:3]):
                if not isinstance(cell, list):
                    cell = [cell]
                
                # --- ПОИСК КООРДИНАТЫ В ЭТОЙ ЯЧЕЙКЕ ---
                # Координата может быть ТОЛЬКО первым или последним элементом списка
                coord_value = ''
                if cell and cell[0] == '-':
                    coord_value = ''
                elif cell and is_number(cell[0]):
                    coord_value = cell[0]
                
                if coord_value == '' and len(cell) > 1:
                    if cell[-1] == '-':
                        coord_value = ''
                    elif is_number(cell[-1]):
                        coord_value = cell[-1]
                
                if i < 3:
                    coords[i] = coord_value
                
                # --- ИЗВЛЕЧЕНИЕ KKS ИЗ ЭТОЙ ЯЧЕЙКИ ---
                for elem in cell:
                    if config.regular_KKS_room.search(elem):
                        all_rooms.add(elem)
                    elif config.regular_KKS_building.search(elem):
                        all_rooms.add(elem)
                    elif config.regular_KKS_equipment.search(elem):
                        all_equip.add(elem)
            
            return list(all_rooms), list(all_equip), coords


    
    # =========================================================================
    # БЛОК Б: ОДИН СПИСОК (одна ячейка) - случаи 4.1 и 4.2
    # =========================================================================
    # Сюда попадаем, если block_cells - это плоский список строк.
    # Например: ['00UKS22R017', '00CYE36', 'Контроллер...', '1779.1', '1756.3', '13.9']
    # Или: ['00UKS10R036', '00CYE36СR001-К02', '...', '-']  (координаты заменены на '-')
    
    if isinstance(block_cells, list):
        # ---------------------------------------------------------
        # ШАГ 1: ИЗВЛЕЧЕНИЕ КООРДИНАТ ИЗ СПИСКА
        # ---------------------------------------------------------
        # Вызываем вспомогательную функцию, которая:
        #   - ищет три числа, идущие ПОДРЯД в списке
        #   - возвращает координаты и список без этих трёх чисел
        coords = parse_coordinates_from_list(block_cells)
        #   coords = ['1779.1', '1756.3', '13.9']

        

        # ---------------------------------------------------------
        # ШАГ 2: ИЗВЛЕЧЕНИЕ KKS ИЗ ОСТАВШИХСЯ ЭЛЕМЕНТОВ
        # ---------------------------------------------------------
        rooms = set()
        equip = set()
        
        # Перебираем отфильтрованный список (без координат)
        for elem in block_cells:
            # 1. KKS помещения
            if config.regular_KKS_room.search(elem):
                rooms.add(elem)
            # 2. KKS здания
            elif config.regular_KKS_building.search(elem):
                rooms.add(elem)      # здание тоже считаем "помещением"
            # 3. KKS оборудования
            elif config.regular_KKS_equipment.search(elem):
                equip.add(elem)
            # Остальное (текстовые описания) игнорируем

        # ---------------------------------------------------------
        # ШАГ 3: ВОЗВРАТ РЕЗУЛЬТАТА
        # ---------------------------------------------------------
        # Преобразуем множества в списки и возвращаем
        return list(rooms), list(equip), coords
    
    # =========================================================================
    # БЛОК В: НИЧЕГО НЕ ПОДОШЛО (ПУСТОЙ ИЛИ НЕКОРРЕКТНЫЙ БЛОК)
    # =========================================================================
    # Если block_cells не является списком списков и не является списком,
    # или если он пустой, возвращаем пустые результаты.
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

    # ===== 3. Извлечение первых полей (индексы 0-3) =====
    # 0. Журнал
    array_row[0] = cleanCyrFromLat(input_row[0]).strip()
    # 1. Номер кабеля
    if input_row[1]:
        array_row[1] = input_row[1][0].replace(',', '.').strip()
    # 2. ККС кабеля
    if input_row[2]:
        array_row[2] = cleanCyrFromLat(input_row[2][0]).strip()
    # 3. Группа раскладки
    if input_row[3]:
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
        if input_row[idx]:
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
        trace = trace.replace(';', ',').replace('  ', ' ').replace(',,', ',').strip().rstrip(',')  
        if len(trace) > 2:
            array_row[8] = trace


    # ========== 7. РАЗБОР БЛОКОВ "ОТКУДА" И "КУДА" ==========
    '''
    АЛГОРИТМ:
    
    Входные данные: input_row — список, где:
      input_row[0] = журнал
      input_row[1] = номер кабеля
      input_row[2] = ККС кабеля
      input_row[3] = группа
      input_row[4...] = блоки "Откуда" и "Куда" (переменное количество ячеек)
      input_row[length_idx] = длина (если найдена)
      input_row[-1] = трасса
    
    Задача: определить, какие ячейки относятся к блоку "Откуда", а какие — к "Куда".
    
    Возможные варианты:
      - Всего 2 ячейки: первая = "Откуда", вторая = "Куда"
      - Всего 6 ячеек: первые три = "Откуда", вторые три = "Куда"
      - Всего 4 ячейки:
          * Вариант А: первая ячейка содержит все данные (KKS и координаты) = "Откуда",
                       остальные три = "Куда" (каждая с одной координатой)
          * Вариант Б: последняя ячейка содержит все данные = "Куда",
                       первые три = "Откуда" (каждая с одной координатой)
    
    Решение:
      1. Если строк 2 → первая = "Откуда", вторая = "Куда"
      2. Если строк 6 → первые три = "Откуда", вторые три = "Куда"
      3. Если строк 4:
         3.1 Проверяем первую строку: ищем в ней ККС помещения, ККС оборудования и ТРИ координаты.
              Если всё найдено → первая строка = "Откуда", остальные три = "Куда"
         3.2 Если не найдено, проверяем последнюю строку: ищем в ней ККС помещения, ККС оборудования и ТРИ координаты.
              Если всё найдено → последняя строка = "Куда", первые три = "Откуда"
    '''
    # ======================================================================

    # Определяем границы диапазона ячеек, которые содержат блоки "Откуда" и "Куда"
    start_idx = 4
    end_idx = length_idx if length_idx != -1 else len(input_row) - 2
    
    if end_idx <= start_idx:
        return array_row
    
    # Получаем список ячеек с блоками
    block_cells = input_row[start_idx:end_idx]
    num_cells = len(block_cells)
    
    # Инициализация переменных для результатов
    list_of_KKS_start = []
    list_of_KKS_equipment_start = []
    list_of_axis_start = []
    
    list_of_KKS_end = []
    list_of_KKS_equipment_end = []
    list_of_axis_end = []
    
    # =========================================================
    # СЛУЧАЙ 1: 2 ячейки (1+1)
    # =========================================================
    if num_cells == 2:
        # Первая ячейка — "Откуда"
        rooms_f, equip_f, coords_f = parse_kks_block([block_cells[0]], building_bounds)
        # Вторая ячейка — "Куда"
        rooms_t, equip_t, coords_t = parse_kks_block([block_cells[1]], building_bounds)
        
        list_of_KKS_start = rooms_f
        list_of_KKS_equipment_start = equip_f
        list_of_axis_start = coords_f if len(coords_f) == 3 else []
        
        list_of_KKS_end = rooms_t
        list_of_KKS_equipment_end = equip_t
        list_of_axis_end = coords_t if len(coords_t) == 3 else []
    
    # =========================================================
    # СЛУЧАЙ 2: 6 ячеек (3+3)
    # =========================================================
    elif num_cells == 6:
        # Первые три ячейки — "Откуда"
        rooms_f, equip_f, coords_f = parse_kks_block(block_cells[:3], building_bounds)
        # Вторые три ячейки — "Куда"
        rooms_t, equip_t, coords_t = parse_kks_block(block_cells[3:], building_bounds)
        
        list_of_KKS_start = rooms_f
        list_of_KKS_equipment_start = equip_f
        list_of_axis_start = coords_f if len(coords_f) == 3 else []
        
        list_of_KKS_end = rooms_t
        list_of_KKS_equipment_end = equip_t
        list_of_axis_end = coords_t if len(coords_t) == 3 else []
    
    # =========================================================
    # СЛУЧАЙ 3: 4 ячейки (1+3 или 3+1)
    # =========================================================
    elif num_cells == 4:
        

        def cell_has_full_data(cell):
            """
            Проверяет, содержит ли ячейка (один список) все три координаты (X,Y,Z).
            KKS помещения и оборудования могут отсутствовать — это допустимо.
            
            Возвращает True, если найдены три координаты (не пустые строки).
            """
            # Парсим ячейку
            _, _, coords = parse_kks_block([cell], building_bounds)
            
            # Проверяем: есть ли три координаты (все три не пустые строки)
            # Или хотя бы три элемента (могут быть пустыми?)
            # Координата считается найденной, если это число (не пустая строка)
            has_three_coords = len(coords) == 3 and all(c != '' for c in coords)
            
            return has_three_coords
        
        # -------------------------------------------------
        # ВАРИАНТ А: первая ячейка содержит все данные = "Откуда"
        # -------------------------------------------------
        if cell_has_full_data(block_cells[0]):
            # Первая ячейка — "Откуда" (содержит всё)
            rooms_f, equip_f, coords_f = parse_kks_block([block_cells[0]], building_bounds)
            # Остальные три ячейки — "Куда" (каждая с одной координатой)
            rooms_t, equip_t, coords_t = parse_kks_block(block_cells[1:4], building_bounds)
            
            list_of_KKS_start = rooms_f
            list_of_KKS_equipment_start = equip_f
            list_of_axis_start = coords_f if len(coords_f) == 3 else []
            
            list_of_KKS_end = rooms_t
            list_of_KKS_equipment_end = equip_t
            list_of_axis_end = coords_t if len(coords_t) == 3 else []
        
        # -------------------------------------------------
        # ВАРИАНТ Б: последняя ячейка содержит все данные = "Куда"
        # -------------------------------------------------
        elif cell_has_full_data(block_cells[-1]):
            # Первые три ячейки — "Откуда" (каждая с одной координатой)
            rooms_f, equip_f, coords_f = parse_kks_block(block_cells[:3], building_bounds)
            # Последняя ячейка — "Куда" (содержит всё)
            rooms_t, equip_t, coords_t = parse_kks_block([block_cells[-1]], building_bounds)
            
            list_of_KKS_start = rooms_f
            list_of_KKS_equipment_start = equip_f
            list_of_axis_start = coords_f if len(coords_f) == 3 else []
            
            list_of_KKS_end = rooms_t
            list_of_KKS_equipment_end = equip_t
            list_of_axis_end = coords_t if len(coords_t) == 3 else []
        
        # -------------------------------------------------
        # Если не удалось определить — используем старый метод (перебор)
        # -------------------------------------------------
        else:
            # Запасной вариант: пытаемся определить по старой логике
            best_rooms_from = []
            best_equip_from = []
            best_coords_from = []
            best_rooms_to = []
            best_equip_to = []
            best_coords_to = []
            
            for size_from in [1, 3]:
                if size_from > num_cells:
                    continue
                from_cells = block_cells[:size_from]
                to_cells = block_cells[size_from:]
                if not to_cells:
                    continue
                
                rooms_f, equip_f, coords_f = parse_kks_block(from_cells, building_bounds)
                rooms_t, equip_t, coords_t = parse_kks_block(to_cells, building_bounds)
                
                if (rooms_f or equip_f or any(coords_f)) and (rooms_t or equip_t or any(coords_t)):
                    best_rooms_from, best_equip_from, best_coords_from = rooms_f, equip_f, coords_f
                    best_rooms_to, best_equip_to, best_coords_to = rooms_t, equip_t, coords_t
                    break
            
            list_of_KKS_start = best_rooms_from
            list_of_KKS_equipment_start = best_equip_from
            list_of_axis_start = best_coords_from if len(best_coords_from) == 3 else []
            
            list_of_KKS_end = best_rooms_to
            list_of_KKS_equipment_end = best_equip_to
            list_of_axis_end = best_coords_to if len(best_coords_to) == 3 else []
    
    # =========================================================
    # СЛУЧАЙ 4: другое количество ячеек (запасной вариант)
    # =========================================================
    else:
        # Используем старый метод перебора для всех остальных случаев
        best_rooms_from = []
        best_equip_from = []
        best_coords_from = []
        best_rooms_to = []
        best_equip_to = []
        best_coords_to = []
        
        # Перебираем возможные размеры
        for size_from in [1, 3]:
            if size_from > num_cells:
                continue
            from_cells = block_cells[:size_from]
            to_cells = block_cells[size_from:]
            if not to_cells:
                continue
            
            rooms_f, equip_f, coords_f = parse_kks_block(from_cells, building_bounds)
            rooms_t, equip_t, coords_t = parse_kks_block(to_cells, building_bounds)
            
            if (rooms_f or equip_f or any(coords_f)) and (rooms_t or equip_t or any(coords_t)):
                best_rooms_from, best_equip_from, best_coords_from = rooms_f, equip_f, coords_f
                best_rooms_to, best_equip_to, best_coords_to = rooms_t, equip_t, coords_t
                break
        
        # Если не нашли подходящий вариант, берём всё как "Откуда"
        if not best_rooms_from and not best_equip_from and not any(best_coords_from):
            rooms_f, equip_f, coords_f = parse_kks_block(block_cells, building_bounds)
            list_of_KKS_start = rooms_f
            list_of_KKS_equipment_start = equip_f
            list_of_axis_start = coords_f if len(coords_f) == 3 else []
        else:
            list_of_KKS_start = best_rooms_from
            list_of_KKS_equipment_start = best_equip_from
            list_of_axis_start = best_coords_from if len(best_coords_from) == 3 else []
            
            list_of_KKS_end = best_rooms_to
            list_of_KKS_equipment_end = best_equip_to
            list_of_axis_end = best_coords_to if len(best_coords_to) == 3 else []
    
    # =========================================================
    # Удаляем дубликаты
    # =========================================================
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