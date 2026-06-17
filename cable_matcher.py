# cable_matcher.py
"""
Поиск ответных частей кабелей между журналами.

Логика работы:
    1. Загружается Excel-база, созданная excel_utils.py (файл Cable base ver.*.xlsx).
    2. Пользователь вводит ККС журнала, требующего объединения.
    3. Просматриваются все кабели из этого журнала.
    4. Для каждого кабеля проверяется:
       - Статус объединения (колонка 28) пуст
       - Требования к объединению (колонка 32) не пусты
    5. Если условия выполнены:
       5.1. Если есть Ответная часть (колонка 33) и Наличие в базе (колонка 34) не пусто
            → ищем в указанных журналах кабель с тем же ККС
       5.2. Иначе → ищем во всей базе кабель с тем же ККС (кроме самого себя)
    6. Результат записывается в новый Excel-файл.
"""

import sys
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font


# ========== КОНФИГУРАЦИЯ ==========
HEADERS = [
    [0, '№ п/п', 8],
    [1, 'ККС кабеля', 25],
    [2, 'Исходный журнал (+ источник, дата)', 50],
    [3, 'Ответный журнал (+ источник, дата)', 50],
    [4, 'Примечание', 40],
    [5, 'Проблемы', 40],
]


def get_column_index(sheet, header_name):
    """Находит индекс колонки по её названию."""
    for col_idx, cell in enumerate(sheet[1], start=1):
        if cell.value and header_name in str(cell.value):
            return col_idx
    return None


def get_journal_info(sheet, journal_kks, row_idx=None):
    """
    Получает информацию о журнале: источник и дату.
    Если передан row_idx, берёт данные из этой строки.
    Иначе ищет первую строку с данным журналом.
    """
    source_col = get_column_index(sheet, 'Источник информации')
    date_col = get_column_index(sheet, 'Дата ревизии')
    
    if source_col is None or date_col is None:
        return '-', '-'
    
    if row_idx is not None:
        # Берём данные из конкретной строки
        source = sheet.cell(row=row_idx, column=source_col).value or '-'
        date = sheet.cell(row=row_idx, column=date_col).value or '-'
        return str(source).strip(), str(date).strip()
    
    # Ищем первую строку с данным журналом
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row and row[0] and str(row[0]).strip() == journal_kks:
            source = row[source_col - 1] if source_col <= len(row) else '-'
            date = row[date_col - 1] if date_col <= len(row) else '-'
            return str(source).strip() if source else '-', str(date).strip() if date else '-'
    
    return '-', '-'


def find_matching_cable(sheet, cable_kks, source_journal, response_journals=None):
    """
    Ищет кабель с таким же ККС в базе.
    
    Args:
        sheet: лист Excel с данными
        cable_kks: ККС искомого кабеля
        source_journal: ККС исходного журнала (чтобы исключить его)
        response_journals: список ККС журналов для поиска (если есть)
    
    Returns:
        tuple: (найден_ли, журнал, источник, дата, строка_с_кабелем)
    """
    if not cable_kks:
        return False, None, None, None, None
    
    kks_col = get_column_index(sheet, 'ККС')
    journal_col = 1  # колонка 'Журнал' всегда первая
    
    if kks_col is None:
        return False, None, None, None, None
    
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if not row:
            continue
        
        # Проверяем ККС
        row_kks = str(row[kks_col - 1]).strip() if kks_col <= len(row) and row[kks_col - 1] else ''
        if row_kks != cable_kks:
            continue
        
        # Проверяем, что это не тот же кабель (исключаем исходный журнал)
        row_journal = str(row[0]).strip() if row[0] else ''
        if row_journal == source_journal:
            continue
        
        # Если указаны ответные журналы — проверяем, что кабель из одного из них
        if response_journals:
            if row_journal not in response_journals:
                continue
        
        # Нашли подходящий кабель
        source, date = get_journal_info(sheet, row_journal, row_idx)
        return True, row_journal, source, date, row_idx
    
    return False, None, None, None, None


def process_journal(excel_path, journal_kks, output_path):
    """
    Основная функция обработки журнала.
    
    Args:
        excel_path: путь к базе данных
        journal_kks: ККС журнала для обработки
        output_path: путь для сохранения результата
    """
    print(f"\n{'='*60}")
    print("ПОИСК ОТВЕТНЫХ ЧАСТЕЙ КАБЕЛЕЙ")
    print(f"{'='*60}")
    print(f"Загрузка: {excel_path}")
    print(f"Обработка журнала: {journal_kks}")
    
    # Загружаем базу
    wb = load_workbook(excel_path, data_only=True)
    sheet = wb.active
    
    # Находим индексы нужных колонок
    kks_col = get_column_index(sheet, 'ККС')
    status_col = get_column_index(sheet, 'Статус объединения')
    req_col = get_column_index(sheet, 'Требования к объединению')
    response_col = get_column_index(sheet, 'Ответная часть (из КЖ)')
    availability_col = get_column_index(sheet, 'Наличие в базе')
    note_col = get_column_index(sheet, 'Примечание')
    
    if kks_col is None:
        raise ValueError("Колонка 'ККС' не найдена")
    
    # Создаём выходной файл
    wb_out = Workbook()
    sheet_out = wb_out.active
    sheet_out.title = "Ответные части"
    
    # Заголовки
    for head in HEADERS:
        col = head[0] + 1
        sheet_out.cell(row=1, column=col, value=head[1])
        sheet_out.column_dimensions[get_column_letter(col)].width = head[2]
    sheet_out.freeze_panes = 'A2'
    
    # Собираем все кабели из указанного журнала
    cables = []
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if not row:
            continue
        row_journal = str(row[0]).strip() if row[0] else ''
        if row_journal == journal_kks:
            cables.append((row_idx, row))
    
    print(f"Найдено кабелей в журнале: {len(cables)}")
    
    # Обрабатываем каждый кабель
    results = []
    problem_count = 0
    result_counter = 0  # ← счётчик для порядкового номера в таблице

    for row_idx, row in cables:
        # Проверяем условия: статус пуст И требования не пусты
        status = str(row[status_col - 1]).strip() if status_col and status_col <= len(row) and row[status_col - 1] else ''
        requirements = str(row[req_col - 1]).strip() if req_col and req_col <= len(row) and row[req_col - 1] else ''
        
        if status or not requirements:
                    continue
        
        cable_kks = str(row[kks_col - 1]).strip() if kks_col <= len(row) and row[kks_col - 1] else ''
        if not cable_kks:
            continue
        
        # Получаем информацию об исходном журнале
        src_source, src_date = get_journal_info(sheet, journal_kks, row_idx)
        src_info = f"{journal_kks} ({src_source}, {src_date})"
        
        # Получаем примечание
        note = str(row[note_col - 1]).strip() if note_col and note_col <= len(row) and row[note_col - 1] else ''
        
        # Ищем ответную часть
        response_journals = None
        response_info = None
        found = False
        problems = []
        
        # Случай 3.1.1: есть Ответная часть и Наличие в базе
        response_val = str(row[response_col - 1]).strip() if response_col and response_col <= len(row) and row[response_col - 1] else ''
        availability_val = str(row[availability_col - 1]).strip() if availability_col and availability_col <= len(row) and row[availability_col - 1] else ''
        
        if response_val and availability_val:
            # Разбиваем ответную часть на отдельные ККС журналов
            response_journals = [j.strip() for j in response_val.split() if j.strip()]
            # Ищем кабель в указанных журналах
            found, resp_journal, resp_source, resp_date, _ = find_matching_cable(
                sheet, cable_kks, journal_kks, response_journals
            )
            if found:
                response_info = f"{resp_journal} ({resp_source}, {resp_date})"
            else:
                problems.append("Не найден кабель в указанных журналах ответной части")
        
        # Случай 3.1.2: если не нашли, ищем во всей базе
        if not found:
            found, resp_journal, resp_source, resp_date, _ = find_matching_cable(
                sheet, cable_kks, journal_kks, None
            )
            if found:
                response_info = f"{resp_journal} ({resp_source}, {resp_date})"
            else:
                problems.append("Ответная часть не найдена в базе")
                problem_count += 1
        
        # Формируем запись
        result_counter += 1  # ← увеличиваем счётчик
        results.append({
            'num': result_counter,  # ← порядковый номер в таблице
            'kks': cable_kks,
            'source': src_info,
            'response': response_info if response_info else '-',
            'note': note if note else '-',
            'problems': '; '.join(problems) if problems else '-'
        })
    
    # Записываем результаты в Excel
    row_idx = 2
    for result in results:
        sheet_out.cell(row=row_idx, column=1, value=result['num'])
        sheet_out.cell(row=row_idx, column=2, value=result['kks'])
        sheet_out.cell(row=row_idx, column=3, value=result['source'])
        sheet_out.cell(row=row_idx, column=4, value=result['response'])
        sheet_out.cell(row=row_idx, column=5, value=result['note'])
        
        # Проблемы подсвечиваем красным
        problems_cell = sheet_out.cell(row=row_idx, column=6, value=result['problems'])
        if result['problems'] != '-':
            problems_cell.font = Font(color="FF0000")
        
        row_idx += 1
    
    # Применяем автофильтр
    sheet_out.auto_filter.ref = sheet_out.dimensions
    
    # Сохраняем
    wb_out.save(output_path)
    
    # Статистика
    print(f"\n{'='*60}")
    print("СТАТИСТИКА")
    print(f"{'='*60}")
    print(f"   Всего кабелей в журнале: {len(cables)}")
    print(f"   Обработано: {len(results)}")
    print(f"   С проблемами: {problem_count}")
    print(f"   Без проблем: {len(results) - problem_count}")
    print(f"\n✅ Сохранено: {output_path}")
    print(f"{'='*60}")
    
    return results


# ========== ТОЧКА ВХОДА ==========
if __name__ == "__main__":
    print("\n" + "="*60)
    print("ПОИСК ОТВЕТНЫХ ЧАСТЕЙ КАБЕЛЕЙ")
    print("="*60)
    
    # Ввод пути к базе
    excel_path = input("Введите путь к файлу базы данных (Cable base ver.*.xlsx): ").strip()
    excel_path = excel_path.strip('"').strip("'")
    
    if not excel_path:
        print("❌ Путь не указан. Программа завершена.")
        sys.exit(1)
    
    if not Path(excel_path).exists():
        print(f"❌ Файл не найден: {excel_path}")
        sys.exit(1)
    
    # Ввод ККС журнала
    journal_kks = input("Введите ККС журнала для поиска ответных частей: ").strip()
    journal_kks = journal_kks.strip('"').strip("'")
    
    if not journal_kks:
        print("❌ ККС журнала не введён. Программа завершена.")
        sys.exit(1)
    
    # Ввод пути для сохранения результата
    output_path = input("Введите путь для сохранения результата (имя_файла.xlsx): ").strip()
    output_path = output_path.strip('"').strip("'")
    
    if not output_path:
        print("❌ Путь для сохранения не указан. Программа завершена.")
        sys.exit(1)
    
    # Добавляем расширение .xlsx, если его нет
    if not output_path.endswith('.xlsx'):
        output_path += '.xlsx'
    
    try:
        process_journal(excel_path, journal_kks, output_path)
        print("\n🏁 Готово!")
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        sys.exit(1)