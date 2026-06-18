# cable_merger.py
"""
Объединение кабелей, разбитых на отрезки, по совпадающему ККС.

Логика работы:
    1. Загружается Excel-файл базы данных, созданный программой cable_master.
    2. Все кабели группируются по полю "ККС" (кабели с одинаковым ККС попадают в одну группу).
    3. Обрабатываются ТОЛЬКО группы с 2 и более кабелями.
       Одиночные кабели игнорируются (не выводятся).
    4. Для каждой группы с 2 кабелями:
        - Проверяется возможность объединения (совпадение группы, сечения, оборудования или здания)
        - Если всё совпадает → кабели объединяются в один (суммируется длина, объединяется трасса)
        - Если не совпадают → попадают на лист "Задвоение KKS"
    5. Группы с 3 и более кабелями попадают на лист "На рассмотрение" (требуют ручного разбора)
    6. Между группами кабелей добавляется одна пустая строка для визуального разделения
    7. Применяются ширина столбцов и автофильтр (как в excel_utils.py)
"""

import re
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from pathlib import Path
from collections import defaultdict
import config


def get_column_index(header_name):
    """
    Возвращает индекс колонки по её названию из config.HEADERS.
    
    Args:
        header_name: название колонки (например, 'Журнал', 'ККС')
    
    Returns:
        int: индекс колонки (0-based)
    """
    for idx, name, width in config.HEADERS:
        if name == header_name:
            return idx
    raise ValueError(f"Колонка '{header_name}' не найдена в config.HEADERS")

def set_column_widths(sheet, headers_config):
    """
    Устанавливает ширину столбцов на листе.
    
    Args:
        sheet: лист Excel
        headers_config: список заголовков из config.HEADERS
    """
    for idx, name, width in headers_config:
        col_letter = get_column_letter(idx + 1)
        sheet.column_dimensions[col_letter].width = width

def apply_autofilter(sheet, num_cols):
    """
    Применяет автофильтр ко всем данным на листе.
    
    Args:
        sheet: лист Excel
        num_cols: количество колонок
    """
    sheet.auto_filter.ref = sheet.dimensions

def freeze_panes(sheet):
    """Замораживает первую строку (заголовки)"""
    sheet.freeze_panes = 'A2'

def setup_excel_sheet(sheet, title, headers, headers_config):
    """
    Настраивает лист Excel: заголовки, ширина колонок, заморозка панели.
    
    Args:
        sheet: лист Excel
        title: название листа
        headers: список заголовков
        headers_config: конфигурация заголовков из config.HEADERS
    """
    sheet.title = title
    
    # Записываем заголовки
    for col_idx, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=col_idx, value=header)
    
    # Устанавливаем ширину колонок
    set_column_widths(sheet, headers_config)
    
    # Замораживаем первую строку
    freeze_panes(sheet)


# ========== ОПРЕДЕЛЕНИЕ ИНДЕКСОВ КОЛОНОК ==========
COL_JOURNAL = get_column_index('Журнал')
COL_CABLE_NUM = get_column_index('Номер кабеля')
COL_KKS = get_column_index('ККС')
COL_GROUP = get_column_index('Группа')
COL_MARK = get_column_index('Марка')
COL_SECTION = get_column_index('Сечение')
COL_LENGTH = get_column_index('Длина')
COL_TRACE = get_column_index('Трасса')
COL_FROM_ROOM = get_column_index('Откуда помещение')
COL_FROM_EQUIP = get_column_index('Откуда оборудование')
COL_TO_ROOM = get_column_index('Куда помещение')
COL_TO_EQUIP = get_column_index('Куда оборудование')
COL_MERGE_STATUS = get_column_index('Статус объединения')

def parse_float(value):
    """Преобразует строку с запятой или точкой в число float"""
    if value is None:
        return 0.0
    s = str(value).strip()
    if s == '':
        return 0.0
    s = s.replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

def get_building(kks_room):
    """Извлекает код здания из KKS помещения (две цифры и три буквы)"""
    if not kks_room:
        return ''
    match = re.match(r'(\d{2}[A-Z]{3})', str(kks_room))
    return match.group(1) if match else ''

def can_merge(row1, row2):
    """
    Проверяет, можно ли объединить два кабеля.
    
    Условия:
        1. Одинаковая группа и сечение
        2. Одинаковое оборудование (откуда, куда) ИЛИ одинаковое здание (откуда, куда)
           Порядок не важен (используются множества)
    """
    # Группа и сечение
    if row1[COL_GROUP] != row2[COL_GROUP]:
        return False
    if row1[COL_SECTION] != row2[COL_SECTION]:
        return False
    
    # Оборудование как множество (порядок не важен)
    equip1 = {row1[COL_FROM_EQUIP], row1[COL_TO_EQUIP]}
    equip2 = {row2[COL_FROM_EQUIP], row2[COL_TO_EQUIP]}
    if equip1 == equip2:
        return True
    
    # Здание как множество (порядок не важен)
    building1 = {get_building(row1[COL_FROM_ROOM]), get_building(row1[COL_TO_ROOM])}
    building2 = {get_building(row2[COL_FROM_ROOM]), get_building(row2[COL_TO_ROOM])}
    if building1 == building2:
        return True
    
    return False

def merge_two_cables(cable1, cable2):
    """
    Объединяет два кабеля в один.
    
    Returns:
        list: новая строка (объединённый кабель)
    """
    # Суммируем длины
    total_length = parse_float(cable1[COL_LENGTH]) + parse_float(cable2[COL_LENGTH])
    
    # Объединяем трассы
    traces = []
    if cable1[COL_TRACE]:
        traces.append(str(cable1[COL_TRACE]).strip())
    if cable2[COL_TRACE]:
        traces.append(str(cable2[COL_TRACE]).strip())
    merged_trace = ', '.join(traces) if traces else ''
    
    # Создаём новый кабель на основе первого
    new_cable = cable1.copy()
    new_cable[COL_JOURNAL] = '-'
    new_cable[COL_CABLE_NUM] = ''
    new_cable[COL_LENGTH] = total_length
    new_cable[COL_TRACE] = merged_trace
    new_cable[COL_MERGE_STATUS] = 'Объединён'
    
    # Очищаем координаты и KKS (откуда и куда)
    for idx in range(COL_FROM_ROOM, COL_TO_EQUIP + 1):
        new_cable[idx] = ''
    
    return new_cable

def process_cable_base(input_path, output_path):
    """
    Основная функция обработки базы кабелей.
    
    Args:
        input_path: путь к исходному Excel-файлу (Cable base ver.*.xlsx)
        output_path: путь для сохранения результата
    """
    print(f"\n{'='*60}")
    print("КАБЕЛЬНЫЙ ЖУРНАЛ - ОБЪЕДИНЕНИЕ ОТРЕЗКОВ")
    print(f"{'='*60}")
    print(f"Загрузка: {input_path}")
    
    # ========== 1. ЗАГРУЗКА ДАННЫХ ИЗ EXCEL ==========
    wb_in = load_workbook(input_path, data_only=True)
    sheet_in = wb_in.active
    
    # Читаем заголовки (первая строка)
    headers = [cell.value for cell in sheet_in[1]]
    num_cols = len(headers)
    
    # Читаем данные (со второй строки)
    data = []
    for row in sheet_in.iter_rows(min_row=2, values_only=True):
        if any(row):
            data.append(list(row))
    
    print(f"Всего строк: {len(data)}")
    
    # ========== 2. ГРУППИРОВКА ПО ККС ==========
    groups = defaultdict(list)
    for row in data:
        kks = row[COL_KKS]
        if kks:
            groups[kks].append(row)
    
    # Фильтруем: оставляем только группы с 2 и более кабелями
    new_groups = {}
    for kks, rows in groups.items():
        if len(rows) >= 2:
            new_groups[kks] = rows
    groups = new_groups
    
    print(f"Всего групп (2+ кабелей): {len(groups)}")
    
    # ========== 3. СОЗДАНИЕ ВЫХОДНОГО ФАЙЛА ==========
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    
    # Создаём три листа
    sheet_review = wb_out.create_sheet("На рассмотрение")
    sheet_merged = wb_out.create_sheet("Объединено")
    sheet_duplicate = wb_out.create_sheet("Задвоение KKS")
    
    # Настраиваем каждый лист (заголовки, ширина колонок, заморозка)
    setup_excel_sheet(sheet_review, "На рассмотрение", headers, config.HEADERS)
    setup_excel_sheet(sheet_merged, "Объединено", headers, config.HEADERS)
    setup_excel_sheet(sheet_duplicate, "Задвоение KKS", headers, config.HEADERS)
    
    # Пустая строка для разделения групп
    empty_row = [''] * num_cols
    
    # Счётчики для статистики
    stats = {
        'merged': 0,      # объединённые кабели
        'review': 0,      # на рассмотрение (3+ кабелей)
        'duplicate': 0    # задвоение KKS
    }
    
    # ========== 4. ОБРАБОТКА ГРУПП ==========
    first_merged = True
    first_duplicate = True
    first_review = True
    
    for kks, group_rows in groups.items():
        # ---- СЛУЧАЙ 1: Ровно 2 кабеля ----
        if len(group_rows) == 2:       
            # Проверяем, можно ли объединить
            if can_merge(group_rows[0], group_rows[1]):
                # Записываем исходные кабели
                sheet_merged.append(group_rows[0])
                sheet_merged.append(group_rows[1])
                # Записываем объединённый кабель
                sheet_merged.append(merge_two_cables(group_rows[0], group_rows[1]))
                # Записываем пустую строку
                sheet_merged.append(empty_row.copy())
                
                stats['merged'] += 3  # 2 исходных + 1 объединённый
                print(f"  ✅ Объединены (2→1): {kks}")
            else:
                # Несовместимы — в задвоение
                sheet_duplicate.append(group_rows[0])
                sheet_duplicate.append(group_rows[1])
                sheet_duplicate.append(empty_row.copy())
                stats['duplicate'] += 2
                print(f"  ❌ Задвоение KKS (несовместимы): {kks}")
        
        # ---- СЛУЧАЙ 2: 3 и более кабелей ----
        else:
            # Записываем все кабели группы на лист "На рассмотрение"
            for row in group_rows:
                sheet_review.append(row)
                stats['review'] += 1
            sheet_review.append(empty_row.copy())
            print(f"  ⚠️ Группа кабелей (требуется ручная проверка): {kks}")
    
    # ========== 5. ПРИМЕНЕНИЕ АВТОФИЛЬТРА ==========
    # Автофильтр применяется ко всем данным на листе
    if sheet_merged.max_row > 1:
        apply_autofilter(sheet_merged, num_cols)
    if sheet_duplicate.max_row > 1:
        apply_autofilter(sheet_duplicate, num_cols)
    if sheet_review.max_row > 1:
        apply_autofilter(sheet_review, num_cols)
    
    # ========== 6. СОХРАНЕНИЕ РЕЗУЛЬТАТА ==========
    wb_out.save(output_path)
    
    # ========== 7. СТАТИСТИКА ==========
    print(f"\n{'='*60}")
    print("СТАТИСТИКА ОБРАБОТКИ")
    print(f"{'='*60}")
    print(f"   Всего групп: {len(groups)}")
    print(f"   Пропущено (одиночные кабели): {sum(1 for g in groups.values() if len(g) == 1)}")
    print(f"   Объединено (2 кабеля → 1): {stats['merged']} строк")
    print(f"   Задвоение KKS (несовместимые пары): {stats['duplicate']} строк")
    print(f"   На рассмотрение (3+ кабелей): {stats['review']} строк")
    print(f"\n✅ Сохранено: {output_path}")
    print(f"{'='*60}")



# ========== ТОЧКА ВХОДА ==========
if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog, messagebox
    
    root = tk.Tk()
    root.withdraw()
    
    input_file = filedialog.askopenfilename(
        title="Выберите файл базы данных (Cable base ver.*.xlsx)",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if not input_file:
        print("Файл не выбран. Программа завершена.")
        exit()
    
    output_file = filedialog.asksaveasfilename(
        title="Сохранить результат как...",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if not output_file:
        print("Место сохранения не выбрано. Программа завершена.")
        exit()
    
    try:
        process_cable_base(input_file, output_file)
        messagebox.showinfo("Готово!", f"Обработка завершена!\n\nРезультат сохранён:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n\n{str(e)}")