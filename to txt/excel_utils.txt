# excel_utils.py
'''
Создание Excel-базы кабельных журналов
Версионирование, перенос старых записей, обработка новых журналов
'''

import json
import openpyxl
import re
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from pathlib import Path
from datetime import date
from docx import Document
from tqdm import tqdm

from doc_utils import take_all_docx_from_dir, extract_all_text_from_docx
from cable_parser import row_parser
import config


def current_version_finder(dir_in, b_name='Cable base ver.'):
    """
    Находит последнюю версию кабельной базы в указанной папке.
    
    Args:
        dir_in: путь к папке с базами
        b_name: шаблон имени базы (по умолчанию 'Cable base ver.')
    
    Returns:
        int: номер следующей версии (текущая версия + 1)
    """
    current_version = 1
    local_list_of_bases_vers = []
    xlsx_files = list(Path(dir_in).glob(b_name + '*.xlsx'))
    
    for file in xlsx_files:
        dot_index = str(file.stem).rfind('.')
        local_list_of_bases_vers.append(int(str(file.stem)[dot_index + 1:]))
    
    if not local_list_of_bases_vers:
        print(f'Кабельные базы не найдены в указанной папке')
    else:
        current_version = max(local_list_of_bases_vers) + 1
        print(f'Найдена последняя версия базы {b_name}{current_version - 1}, текущая версия - {current_version}')
    
    return current_version


def remove_plus_from_numbers(obj):
    """
    Рекурсивно обходит списки и удаляет символ '+' из строк,
    которые начинаются с '+' и затем содержат число (целое или дробное).
    """
    if isinstance(obj, list):
        return [remove_plus_from_numbers(item) for item in obj]
    elif isinstance(obj, str):
        if obj.startswith('+') and re.match(r'^\+\d+(?:[.,]\d+)?$', obj):
            return obj[1:]
        else:
            return obj
    else:
        return obj


def build_cable_database(journals_dir, output_dir, progress_callback=None):
    """
    Создаёт кабельную базу из журналов в указанной папке.
    
    Args:
        journals_dir: папка с журналами (.doc, .docx)
        output_dir: папка, в которую будет сохранена новая база
        progress_callback: функция для обновления прогресса (принимает current, total, message)
    """
    print("\nЗапуск build_cable_database...")

    # Загрузка границ зданий из JSON
    with open('KKS_building_bounds.json', 'r', encoding='utf-8') as f:
        building_bounds = json.load(f)

    # Поиск старой версии базы
    b_name = config.BASE_NAME
    current_version = current_version_finder(output_dir, b_name)
    
    # Получение списка всех журналов (.doc и .docx)
    journals = take_all_docx_from_dir(journals_dir)
    journals_names = [j.stem for j in journals] if journals else []
    total_journals = len(journals)
    
    # Отчёт о прогрессе: найдено журналов
    if progress_callback:
        progress_callback(0, total_journals, f"Найдено {total_journals} журналов")

    # Создание новой базы
    wb = Workbook()
    sheetWrite = wb.active
    sheetWrite.title = f"База данных вер. {current_version}"

    # Заголовки таблицы
    heads = config.HEADERS
    for head in heads:
        col = head[0] + 1
        sheetWrite.cell(row=1, column=col, value=f"{head[0]+1}. {head[1]}")
        sheetWrite.column_dimensions[get_column_letter(col)].width = head[2]
    sheetWrite.freeze_panes = 'A2'

    today = date.today()
    date_string = today.strftime("%Y.%m.%d")
    
    rowWrite = 2
    
    # ----- Копирование из старой версии базы -----
    if current_version > 1:
        old_path = Path(output_dir) / f"{b_name}{current_version - 1}.xlsx"
        if old_path.exists():
            rb = load_workbook(old_path, data_only=True)
            sheetRead = rb.active

            rows_to_copy = []
            for idx, row in enumerate(sheetRead.iter_rows(min_row=2, values_only=True), start=2):
                journal_name = str(row[0]) if row[0] else ''
                if journal_name and journal_name not in journals_names:
                    rows_to_copy.append(row)

            total_old = len(rows_to_copy)
            if progress_callback:
                progress_callback(0, total_journals + total_old, f"Перенос старых записей (всего {total_old})")
            
            for i, row_data in enumerate(rows_to_copy):
                for col_idx, value in enumerate(row_data, start=1):
                    sheetWrite.cell(row=rowWrite, column=col_idx, value=value)
                rowWrite += 1
                if progress_callback and i % 10 == 0:
                    progress_callback(i, total_journals + total_old, f"Перенос старых записей: {i}/{total_old}")

    # ----- Обработка новых журналов -----
    if journals:
        if progress_callback:
            progress_callback(0, total_journals, "Начало обработки журналов")
        
        for i, journal in enumerate(journals):
            try:
                if progress_callback:
                    progress_callback(i, total_journals, f"Обработка: {journal.stem}")
                
                j_raw_content = extract_all_text_from_docx(Document(journal))
                if not j_raw_content:
                    if progress_callback:
                        progress_callback(i, total_journals, f"Нет данных в {journal.stem}")
                    continue
                
                rows_processed = 0
                for raw_row in j_raw_content:
                    raw_row = remove_plus_from_numbers(raw_row)
                    
                    if len(raw_row) > 6:
                        raw_row.insert(0, journal.stem)
                        processed_row = row_parser(raw_row, building_bounds)

                    if processed_row:
                        processed_row.append(date_string)
                        processed_row.append(current_version)
                        
                        for idx, val in enumerate(processed_row, start=1):
                            sheetWrite.cell(row=rowWrite, column=idx, value=val)
                        
                        sheetWrite.cell(row=rowWrite, column=30, value=str(raw_row))
                        
                        # Расчёт минимальной длины кабеля
                        try:
                            # Функция для преобразования координаты (замена запятой на точку)
                            def to_float(val):
                                if not val:
                                    return None
                                return float(str(val).replace(',', '.'))
                            
                            length = processed_row[7]
                            if not length:
                                raise ValueError("Нет длины")
                            
                            length_val = float(length)
                            
                            from_x = to_float(processed_row[11])
                            from_y = to_float(processed_row[12])
                            from_z = to_float(processed_row[13])
                            to_x = to_float(processed_row[16])
                            to_y = to_float(processed_row[17])
                            to_z = to_float(processed_row[18])
                            
                            # Проверяем, что все координаты есть
                            if None in (from_x, from_y, from_z, to_x, to_y, to_z):
                                raise ValueError("Не все координаты")
                            
                            min_len = round(
                                abs(from_x - to_x) + 
                                abs(from_y - to_y) + 
                                abs(from_z - to_z)
                            )
                            
                            sheetWrite.cell(row=rowWrite, column=31, value=min_len)
                            
                            if length_val < min_len:
                                sheetWrite.cell(row=rowWrite, column=32, value='ДА')
                                
                        except Exception:
                            # Если что-то пошло не так — просто пропускаем
                            pass
                        
                        rowWrite += 1
                        rows_processed += 1

                if progress_callback:
                    progress_callback(i + 1, total_journals, f"Готово: {journal.stem} ({rows_processed} строк)")

            except Exception as e:
                print(f"Ошибка обработки {journal.stem}: {e}")
                if progress_callback:
                    progress_callback(i + 1, total_journals, f"Ошибка в {journal.stem}: {str(e)[:50]}", is_error=True)
    
    # Применяем автофильтр и сохраняем
    sheetWrite.auto_filter.ref = sheetWrite.dimensions
    output_path = Path(output_dir) / f"{b_name}{current_version}.xlsx"
    wb.save(output_path)
    
    if progress_callback:
        progress_callback(total_journals, total_journals, "Сохранение базы данных")
    
    print(f"\n✅ База данных сохранена: {output_path}")
    print("🏁 Готово!")