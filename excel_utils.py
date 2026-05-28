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

# ----- функция для поиска актуальной версии базы -----

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

# ----- функция для удаления "+" из координат -----

def remove_plus_from_numbers(obj):
    """
    Рекурсивно обходит списки и удаляет символ '+' из строк,
    которые начинаются с '+' и затем содержат число (целое или дробное).
    """
    if isinstance(obj, list):
        return [remove_plus_from_numbers(item) for item in obj]
    elif isinstance(obj, str):
        # Если строка начинается с '+' и далее число (возможно с десятичной точкой или запятой)
        if obj.startswith('+') and re.match(r'^\+\d+(?:[.,]\d+)?$', obj):
            return obj[1:]   # удаляем '+'
        else:
            return obj
    else:
        return obj


# ----- основная функция создания базы -----

def build_cable_database(journals_dir, output_dir):
    """
    Создаёт кабельную базу из журналов в указанной папке.
    
    Args:
        journals_dir: папка с журналами (.doc, .docx)
        output_dir: папка, в которую будет сохранена новая база
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
            # переносятся только журналы, которых нет среди новых журналов
            for idx, row in enumerate(sheetRead.iter_rows(min_row=2, values_only=True), start=2):
                journal_name = str(row[0]) if row[0] else ''
                if journal_name and journal_name not in journals_names:
                    rows_to_copy.append(row)

            print(f"\n📦 Перенос данных из предыдущей версии ({old_path.name})...")
            with tqdm(total=len(rows_to_copy), desc="Копирование", unit="строка") as pbar:
                for row_data in rows_to_copy:
                    for col_idx, value in enumerate(row_data, start=1):
                        sheetWrite.cell(row=rowWrite, column=col_idx, value=value)
                    rowWrite += 1
                    pbar.update(1)

    # ----- Обработка новых журналов -----
    if journals:
        print("\nОбработка журналов...")
        for journal in tqdm(journals, desc="Обработка", unit="журнал"):
            try:
                # Извлечение содержимого таблиц из документа
                j_raw_content = extract_all_text_from_docx(Document(journal))
                if not j_raw_content:
                    continue
                
                for raw_row in j_raw_content:
                    # Предобработка: удаляем '+' из чисел
                    raw_row = remove_plus_from_numbers(raw_row)
                    
                    if len(raw_row) > 6:
                        raw_row.insert(0, journal.stem)
                        processed_row = row_parser(raw_row, building_bounds)

                    if processed_row:
                        # Добавляем дату и версию
                        processed_row.append(date_string)
                        processed_row.append(current_version)
                        
                        # Записываем основные данные
                        for idx, val in enumerate(processed_row, start=1):
                            sheetWrite.cell(row=rowWrite, column=idx, value=val)
                        
                        # Сырая строка (Raw) — в столбец AD (индекс 30, если считать с 0)
                        sheetWrite.cell(row=rowWrite, column=30, value=str(raw_row))
                        
                        # Расчёт минимальной длины кабеля
                        try:
                            length = float(processed_row[7])
                            xyz_from = [float(processed_row[11]), float(processed_row[12]), float(processed_row[13])]
                            xyz_to = [float(processed_row[16]), float(processed_row[17]), float(processed_row[18])]
                            min_len = round(abs(xyz_from[0] - xyz_to[0]) + abs(xyz_from[1] - xyz_to[1]) + abs(xyz_from[2] - xyz_to[2]))
                            sheetWrite.cell(row=rowWrite, column=31, value=min_len)
                            if length < min_len:
                                sheetWrite.cell(row=rowWrite, column=32, value='ДА')
                        except Exception as e:
                            # Если что-то пошло не так — просто пропускаем
                            pass
                        
                        rowWrite += 1

            except Exception as e:
                print(f"Ошибка обработки {journal.stem}: {e}")
    
    # Применяем автофильтр и сохраняем
    sheetWrite.auto_filter.ref = sheetWrite.dimensions
    output_path = Path(output_dir) / f"{b_name}{current_version}.xlsx"
    wb.save(output_path)
    print(f"\n✅ База данных сохранена: {output_path}")
    print("🏁 Готово!")