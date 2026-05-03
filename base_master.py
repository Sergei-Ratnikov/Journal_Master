import re
import openpyxl
from openpyxl import load_workbook
from docx import Document
from pathlib import Path
from datetime import date
import win32com.client
import os
import shutil
from utils import convert_doc_to_docx
from utils import cleanCyrFromLat
from utils import is_subset_with_lists
from utils import all_deep_empty
from utils import get_mismatch_indices
from utils import move_file
import config
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# ----- блок функций для поиска актуальной версии журнала

def current_version_finder(dir_in, b_name = 'Cable base ver.'):
    """
    ищу наличие базы в папке и определяю текущую версию базы

    Args:
        dir_in: адрес папки, где ищу базу
        b_name: шаблон имени базы

    Returns:
        current_version - ВАЖНО! Текущая версия - это та, которую я делаю! Если не было базы, текущая - 1, если была в папке 1, то текущая 2 и т.д.
    """

    current_version = 1 # версия 1 по умолчанию

    local_list_of_bases_vers = [] # список int из версий, которые лежат в папке
    xlsx_files = list(Path(dir_in).glob(b_name + '*.xlsx')) # нахожу все файлы, начинающиеся как название базы и заканивающиеся .xlsx
    for file in xlsx_files:
        dot_index = str(file.stem).rfind('.')
        local_list_of_bases_vers.append(int(str(file.stem)[dot_index + 1:]))  # отрезаю от названия все после точки, превращаю в int и добавляю в список
    
    if not local_list_of_bases_vers:
        print(f'Кабельные базы не найдены в указанной папке')
    else:
        current_version = max(local_list_of_bases_vers) + 1
        print(f'найдена последняя версия кабельной базы {b_name}{current_version - 1}, текущая версия для внесения изменений - {current_version}')
    
    return current_version

# ----- блок функций для работы с docx

def convert_numbering_to_text(doc_path):
    """
    вспомогательная функция

    Конвертирует автонумерацию в текстовом документе в обычный текст и пересохраняет .docx
    """
    abs_path = os.path.abspath(doc_path)     # абсолютный путь

    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")  # запускает Microsoft Word в фоновом режиме и создаёт объект, через который Python может управлять Word'ом
        word.Visible = False
        word.DisplayAlerts = False  # отключаем все предупреждения!
        doc = word.Documents.Open(abs_path) # открываю документ
        doc.ConvertNumbersToText() # Конвертирую нумерацию в текст

        # новый путь
        # new_path = abs_path.replace('.docx', '_numbered.docx')
        # doc.SaveAs(new_path)
        # return new_path

        # пересохраняю    
        doc.SaveAs(abs_path)
        return Path(abs_path)
    except Exception as e:
        print(f"Ошибка: {e}")
        return None
    finally:
        if doc:
            doc.Close()
        if word:
            word.Quit()

# ----- блок функций для извлечения данных из таблиц

def remove_duplicate_pairs(lst):
    """
    вспомогательная функция
    исправляет ошибку журнала с задвоением строк.
    Убирает дубли из списка, где каждый элемент повторяется дважды подряд.
    если list1 = [ [1], [1],   [2,2], [2,2],   [3,4], [3,4]  ]
    то возвращает list2 = [ [1],   [2,2],   [3,4] ]
    иначе возвращает list1
    """
    if len(lst) % 2 != 0:
        return lst
    
    result = []
    for i in range(0, len(lst), 2):
        if lst[i] != lst[i + 1]:
            return lst
        result.append(lst[i])
    return result

def remove_duplicates(lst):
    """
    Убирает дубликаты, сохраняя последовательность.
    Особые случаи:
    - три подряд ['0'], ['0'], ['0'] — оставляем все три
    - три подряд ['-'], ['-'], ['-'] — оставляем все три
    
    Примеры использования
    list1 = [[1], [1], [2,2], [2,2], [3,4]]
    print(remove_duplicates_with_exceptions(list1))
    [[1], [2, 2], [3, 4]]

    С особыми случаями
    list2 = [['0'], ['0'], ['0'], [1], [1], [2], [3], [3]]
    print(remove_duplicates_with_exceptions(list2))
    [['0'], ['0'], ['0'], [1], [2], [3]]

    # Смешанный случай
    list4 = [['0'], ['0'], ['0'], [1], [1], [2], [2], [3], [3]]
    print(remove_duplicates_with_exceptions(list4))
    # [['0'], ['0'], ['0'], [1], [2], [3]]

    """
    if not lst:
        return []
    
    result = []
    skip = 0
    
    for i in range(len(lst)):
        # Пропускаем элементы, которые уже обработаны
        if skip > 0:
            skip -= 1
            continue
        
        # Проверяем особый случай: три подряд ['0']
        if i + 2 < len(lst) and lst[i] == ['0'] and lst[i + 1] == ['0'] and lst[i + 2] == ['0']:
            result.extend([['0'], ['0'], ['0']])
            skip = 2  # пропускаем следующие 2 элемента (всего 3, текущий уже обработан)
            continue
        
        # Проверяем особый случай: три подряд ['-']
        if i + 2 < len(lst) and lst[i] == ['-'] and lst[i + 1] == ['-'] and lst[i + 2] == ['-']:
            result.extend([['-'], ['-'], ['-']])
            skip = 2
            continue
        
        # Обычная обработка: добавляем текущий элемент
        result.append(lst[i])
        
        # Пропускаем все последующие одинаковые элементы
        j = i + 1
        while j < len(lst) and lst[j] == lst[i]:
            j += 1
        skip = j - i - 1  # количество пропускаемых элементов
    
    return result

def find_first_and_last_sublist_index(big_list, small_list):
    """
    Находит индексы первого и последнего вхождения small_list в big_list

    # Пример
    big =   [1, 2, 3, 1, 2, 3, 4, 1, 2, 3]
    small = [1, 2, 3]

    first, last = find_first_and_last_sublist_index(big, small)
    print(f"Первое вхождение: индекс {first}")  # 0
    print(f"Последнее вхождение: индекс {last}")  # 7

    """
    len_small = len(small_list)
    len_big = len(big_list)
    
    first_index = -1
    last_index = -1
    
    for i in range(len_big - len_small + 1):
        if big_list[i:i + len_small] == small_list:
            if first_index == -1:
                first_index = i
            last_index = i
    
    return first_index, last_index

def extract_all_text_from_cell(cell):
    """
    Извлекает ВЕСЬ текст из ячейки, включая вложенные таблицы, возвращает список строк ЯЧЕЙКИ

    Returns:
        result = [] возвращаю содержимое ячейки, список текстовых строк
    """
    result = []
    # Текст из параграфов в ячейке
    if cell:
        for para in cell.paragraphs:
            if para.text.strip():
                text = para.text.strip()
                text = ' '.join(text.split())
                text = text.replace('\xad', '-').replace('\n', ' ')
                result.append(text)
        
        # Текст из вложенных таблиц построчно вносится в result
        for inner_table in cell.tables: # перебор всех вложенных в ячейку таблиц
            # result.append('Вложенная таблица')
            for it_row in inner_table.rows:
                for nested_cell in it_row.cells:
                    inner_table_result = extract_all_text_from_cell(nested_cell)  # РЕКУРСИЯ для обработки вложенных таблиц
                    result.extend(inner_table_result)
    return result

def extract_all_text_from_table(table):
    """
    0. Извлекает ВЕСЬ текст из таблицы, возвращает массив строк таблицы - ячеек - строк содержимого
    1. если в строке столбец слева или справа пустой, то их удаляем
    2. TODO проверка на резервируемость если в таблице в строке первый список имеет более 1 элемента
    3. если в таблице в строке первая ячейка не содержит регулярное выражение regular_num - строку удаляем. а если содержит, 
        то оставляем только ту часть, которая соответствует регулярному выражению
    4. заплатка ошибки повторения строк из-за вложенной таблицы в столбце 2 (группа раскладки) -  если две строки совпадают 
        во всем кроме группы, то нижняя удаляется
    5. обработка записи координат, указанных в одной ячейке через символ табуляции: '1783.3 \t1780.2 \t1.0' - 
        такую запись нужно разбить и одну ячейку заменить тремя
        ищу строку типа   ['00UKS10R088', '00BYF49', 'Шкаф УСО / RTU cabinet', '1783.3 \t1780.2 \t1.0']
        заменяю на        ['00UKS10R088', '00BYF49', 'Шкаф УСО / RTU cabinet', '1783.3', '1780.2', '1.0']

    6. заплатка ошибки, которая расщепляет одну строку на несколько. Я нахожу все строки с одним номером в 0 столбце и 
        склеиваю их в 1 общую строку по принциу:
        если значения совпадают, то игнор, если нет, то склейка
        и затем удаляю расщепленные строки, заменяя одной объединенной

    Returns:
        table_contents =
            [   [строка таблицы 1: [ячейка 1]   [ячейка 2]    [ячейка 3] ] , 
                [строка таблицы 2: [ячейка 1]   [ячейка 2]    [ячейка 3] ]
            ]
            возвращаю содержимое таблицы, список строк, каждая строка - список ячеек, каждая ячейка - 
            список текстовых строк
    """
    # 0
    if not table.rows:
        print("Таблица пуста")
        return
    
    table_contents = []
    try:
        row_index = 0
        for row in table.rows:              # перебор строк в таблице
            current_row = []
            for cell in row.cells:
                current_row.append(extract_all_text_from_cell(cell))

            if len(current_row) > 6 and not all_deep_empty(current_row):
                current_row = remove_duplicate_pairs(current_row)
                table_contents.append(current_row)

                # if current_row[0][0] == '8.0001':
                #     for cell in row.cells:
                #         print(f'row {row_index} - {extract_all_text_from_cell(cell)}')

            row_index += 1

    except Exception as e:
        print(f'Ошибка! extract_all_text_from_table (0) -  {e}')

    if not table_contents:
        return []

    new_table_contents = []
    try:
        for row in table_contents:
            current_row = row
    # 1 проверка пустоты левого столбца
            if all_deep_empty(row[0]) and config.regular_num.search(' '.join(row[1])):
                current_row = row[1:]
    #   если правый столбец пустой, а следующий содержит буквы или минус (столбец с трассой) или тоже пустой (столбец с пустой трассой)
            if all_deep_empty(row[-1]) and (    any(config.regular_letter_minus.search(s) for s in row[-2]) 
                                            or  all_deep_empty(row[-2])):
                current_row = current_row[:-1]
            new_table_contents.append(current_row)
            # print(f'{current_row}')
        table_contents = new_table_contents
        new_table_contents = []
    except Exception as e:
        print(f'Ошибка! extract_all_text_from_table (1) -  {e}')

    # Удаление шапки и примечаний
    # 2
    # 3
    # Удаление строк таблицы, не содержащих в первой ячейке номер
    # TODO проверить на избыточную строгость - удаляются строки без нумерации в некоторых журналах
    
    '''
    предыдущая версия, если все работает, то удалить
    try:
        list_for_delete = [] # список строк для удаления
        # составляю список строк таблицы для удаления

        for index, row in enumerate(table_contents):
            match = config.regular_num.search(row[0][0])
            if match:
                table_contents[index][0] = [match.group()]
                # в этом месте я потерял информацию о резервировании!
            else:
                list_for_delete.append(index)
        new_table_contents = [item for idx, item in enumerate(table_contents) if idx not in list_for_delete] # новый список из неудаляемых строк
        table_contents = new_table_contents
    '''

    try:
        list_for_delete = [] # список строк для удаления
        for index, row in enumerate(table_contents):
            # Проверяем, что строка не пуста и первая ячейка содержит хотя бы один элемент
            if row and row[0] and row[0][0]:
                match = config.regular_num.search(row[0][0])
                if match:
                    table_contents[index][0] = [match.group()]
                else:
                    list_for_delete.append(index)
            else:
                # Если первая ячейка пуста, строку тоже удаляем
                list_for_delete.append(index)
        
        new_table_contents = [item for idx, item in enumerate(table_contents) if idx not in list_for_delete]
        table_contents = new_table_contents

    except Exception as e:
        print(f'Ошибка! extract_all_text_from_table (3) - {e}')

    
    except Exception as e:
        print(f'Ошибка! extract_all_text_from_table (3) -  {e}')

    # 4
    # заплатка ошибки повторения строк из-за вложенной таблицы в столбце 2 (группа раскладки) -  
    # если две строки совпадают во всем кроме группы, то нижняя удаляется
    '''
    предыдущая версия
    try:
        list_for_delete = [] # список строк для удаления
        for i in range (0, len(table_contents) - 1):
            if get_mismatch_indices(table_contents[i], table_contents[i + 1]) == [2]: # если две строки различаются только значением 3 столбца, то
                list_for_delete.append(i + 1)
        new_table_contents = [item for idx, item in enumerate(table_contents) if idx not in list_for_delete] # новый список из неудаляемых строк
        table_contents = new_table_contents
    except Exception as e:
        print(f'Ошибка! extract_all_text_from_table (4) -  {e}')
    '''

    try:
        list_for_delete = []
        for i in range(0, len(table_contents) - 1):
            # Проверяем, что строки существуют и имеют одинаковую длину
            if (i + 1 < len(table_contents) and 
                len(table_contents[i]) == len(table_contents[i + 1])):
                
                mismatch_indices = get_mismatch_indices(table_contents[i], table_contents[i + 1])
                if mismatch_indices == [2]:  # если различаются только в 3 столбце
                    list_for_delete.append(i + 1)
            elif i + 1 < len(table_contents):
                # Строки разной длины — не пытаемся сравнивать
                continue
                
        new_table_contents = [item for idx, item in enumerate(table_contents) if idx not in list_for_delete]
        table_contents = new_table_contents
    except Exception as e:
        print(f'Ошибка! extract_all_text_from_table (4) - {e}')


    # 6
    # заплатка ошибки, которая расщепляет одну строку на несколько. Я нахожу все строки с одним номером в 0 столбце и 
    # склеиваю их в 1 общую строку по принциу:
    # если значения совпадают, то игнор, если нет, то склейка
    # и затем удаляю расщепленные строки, заменяя одной объединенной
    new_table_contents = []

    cable_numbers = [] # список оригинальных номеров кабелей
    try:
        for i_row in range (0, len(table_contents)):
            if table_contents[i_row][0] not in cable_numbers:
                cable_numbers.append(table_contents[i_row][0])
    # здесь я собрал список номеров кабелей, которые есть в table_contents

        for cable_number in cable_numbers:
    # Я беру каждый номер кабеля
            list_of_rows_with_current_number = []  
    # создаю список индексов строк из table_contents, содержащих расщепленную запись о кабеле
            for j, rowrow in enumerate(table_contents):
                if rowrow[0] == cable_number:
                    list_of_rows_with_current_number.append(j)
    # и добавляю в него все индексы строк из table_contents с информацией о кабеле


            # if cable_number[0] == '8.0001':
            #     for j in list_of_rows_with_current_number:
            #         print(f'j len = {len(table_contents[j])}')


            current_row = []    # создаю новую строку, которую буду наполнять информацией о текущем кабеле cable_number
            i = list_of_rows_with_current_number[0]
    # i - первый индекс расщепленной строки для данного кабеля в table_contents
            current_row = table_contents[i]
    # первую из строк просто беру всю
            if len(list_of_rows_with_current_number) > 1:
    # если строка расщеплена, т.е. есть больше одной строки с этим номером кабеля
                for i in range (1, len(list_of_rows_with_current_number)):
                    r = list_of_rows_with_current_number[i] # индекс строки, содержащая расщепленную запись о кабеле в table_contents
                    for i_cell in range (0, len(current_row) - 1): # индекс ячейки в строке
                        if not is_subset_with_lists(table_contents[r][i_cell], current_row[i_cell]):
                            current_row[i_cell].extend(table_contents[r][i_cell])

            new_table_contents.append(current_row)
    # записываю новую строку
        table_contents = new_table_contents
    except Exception as e:
        print(f'Ошибка! extract_all_text_from_table 6 -  {e}')

    # 7
    # удаляю повторяющиеся ячейки из строки
    new_table_contents = []
    try:
        for row in table_contents:
            rowrow = remove_duplicates(row)
            new_table_contents.append(rowrow)
        table_contents = new_table_contents          
    except Exception as e:
        print(f'Ошибка! extract_all_text_from_table 7 -  {e}')


    # 5
    # обработка записи координат, указанных в одной ячейке через символ табуляции или пробел: 
    # '1783.3 \t1780.2 \t1.0'
    # такую запись нужно разбить и одну ячейку заменить тремя
    try:
        regular_axis_local = re.compile(r'[+-]?\d{1,8}\.?,?\d{0,3}')
        for i_row in range (0, len(table_contents)):
            for j_col in range (3, len(table_contents[i_row]) - 1): # table_contents[i_row][j_col] - список строк в ячейке
                # ищу строку типа   ['00UKS10R088', '00BYF49', '1783.3 \t1780.2 \t1.0']
                # заменяю на        ['00UKS10R088', '00BYF49', '1783.3', '1780.2', '1.0']
                if table_contents[i_row][j_col]:
                    # parts = re.split(r'\s*\t\s*', table_contents[i_row][j_col][-1].strip()) # пробую разделить последнюю строку на слова
                    last_string = table_contents[i_row][j_col][-1]
                    parts = last_string.strip().split()
                    # Проверяем, что получилось ровно три части
                    # if len(parts) == 3:
                    #     # Проверяем каждую часть на соответствие регулярному выражению
                    #     if all(regular_axis_local.fullmatch(p) for p in parts):
                    #         del table_contents[i_row][j_col][-1]        
                    #     table_contents[i_row][j_col].extend(parts)

                    if len(parts) == 3 and all(regular_axis_local.fullmatch(p) for p in parts):
                        table_contents[i_row][j_col][:-1].extend(parts)
        # print(f'extract_all_text_from_table 5 ------ table_contents -------- {len(table_contents)}')

    except Exception as e:
        print(f'Ошибка! extract_all_text_from_table 5 -  {e}')



    return table_contents

def extract_all_text_from_docx(docx):
    """
    Извлекает ВЕСЬ текст из всех таблиц документа docx, возвращает массив строк таблицы - ячеек - строк содержимого

    Args:
        docx - Переменная Document(docx_path) имеет тип docx.document.Document.
        Это объект, представляющий открытый документ Word в библиотеке python-docx

    Returns:
        table_contents =  [ [строка таблицы 1: [ячейка 1]   [ячейка 2]    [ячейка 3] ] , 
                            [строка таблицы 2: [ячейка 1]   [ячейка 2]    [ячейка 3] ]
                        ]
    """
    if not docx.tables:
        print("Нет таблиц")
        return
    
    table_contents = []    
    # print(f'extract_all_text_from_docx - docx.tables, таблиц в журнале - {len(docx.tables)}')

    try:
        for table in docx.tables:
            if not table.rows:
                print("Таблица пуста")
                continue
            else:
                table_contents.extend(extract_all_text_from_table(table))
                # print(f'extract_all_text_from_docx - прочитал таблицу')
    except Exception as e:
        print(f'Ошибка! extract_all_text_from_docx - {e}')


    
    return table_contents # возвращаю содержимое всех таблиц файла, список строк, каждая строка - список ячеек, каждая ячейка - список текстовых строк

def take_all_docx_from_dir(journals_directory):
    """
    поиск всех документов doc и docx в папке
    преобразует doc в docx функцией convert_doc_to_docx
    возвращает список файлов

    Args:
        journals_directory - строка с адресом папки

    Returns:
        files - общий список всех docx файлов - существующих + отконвертированных из doc в docx
                        ]
    """
def take_all_docx_from_dir(journals_directory):
    """
    поиск всех документов doc и docx в папке
    преобразует doc в docx функцией convert_doc_to_docx
    переносит исходные .doc файлы в подпапку 'doc files'
    возвращает список файлов .docx (существующих + сконвертированных)

    Args:
        journals_directory - строка с адресом папки

    Returns:
        files - общий список всех docx файлов (объекты Path)
    """
    source_dir = Path(journals_directory).resolve()
    
    # Папка для исходных .doc файлов
    doc_backup_dir = source_dir / 'doc files'
    
    # Собираем все doc и docx файлы
    docx_files = list(source_dir.glob('*.docx'))
    doc_files = list(source_dir.glob('*.doc'))
    
    # Множество имён существующих .docx (без расширения)
    existing_docx_stems = {f.stem for f in docx_files}
    
    # Результат: начинаем с уже существующих .docx
    files = docx_files.copy()
    
    # Обрабатываем каждый .doc файл
    for doc_file in doc_files:
        # Если нет одноимённого .docx, конвертируем
        if doc_file.stem not in existing_docx_stems:
            new_docx_str = convert_doc_to_docx(str(doc_file))
            files.append(Path(new_docx_str))
        
        # Переносим исходный .doc в папку 'doc files'
        try:
            doc_backup_dir.mkdir(exist_ok=True)  # создаём папку, если её нет
            dest_path = doc_backup_dir / doc_file.name
            shutil.move(str(doc_file), str(dest_path))
            print(f"  📁 .doc файл перемещён: {doc_file.name} -> doc files/")
        except Exception as e:
            print(f"  ⚠️ Не удалось переместить {doc_file.name}: {e}")
    
    return files

# ----- блок функций для парсинга

# def extract_coordinates_from_line(line):
#     """
#     Извлекает все координаты из строки (могут быть разделены пробелами/табуляциями)
#     """
#     parts = line.split()
#     coords = []
#     for part in parts:
#         # Очищаем от лишних символов
#         part = part.strip()
#         # Проверяем, соответствует ли часть шаблону координаты
#         if config.regular_axis_full.match(part):
#             coords.append(part)
#     return coords

# def extract_kks_from_line(line):

#     """
#     Извлекает все KKS из строки (здание, оборудование, помещение)

#     result = parse_int_list(int_list)
    
#     print("Координаты начала:", result['list_of_axis_start'])
#     print("Координаты конца:", result['list_of_axis_end'])
#     print("KKS начала:", result['list_of_KKS_start'])
#     print("KKS конца:", result['list_of_KKS_end'])

#     """
    
#     # Проверяем на KKS помещения (самый специфичный)
#     if config.regular_KKS_room.search(line):
#         # Может быть несколько в одной строке
#         matches = config.regular_KKS_room.findall(line)
#         kks_list.extend(matches)
    
#     # Проверяем на KKS здания
#     if config.regular_KKS_building.search(line):
#         matches = config.regular_KKS_building.findall(line)
#         kks_list.extend(matches)
    
#     # Проверяем на KKS оборудования
#     if config.regular_KKS_equipment.search(line):
#         matches = config.regular_KKS_equipment.findall(line)
#         # Фильтруем, чтобы не добавлять уже найденные
#         for match in matches:
#             if match not in kks_list:
#                 kks_list.append(match)
    
#     return kks_list

# def parse_xyz_kks(int_list):
#     """
#     Основная функция парсинга списка int_list

#     result = parse_int_list(int_list)
    
#     print("Координаты начала:", result['list_of_axis_start'])
#     print("Координаты конца:", result['list_of_axis_end'])
#     print("KKS начала:", result['list_of_KKS_start'])
#     print("KKS конца:", result['list_of_KKS_end'])

#     """
#     list_of_axis_start = []
#     list_of_axis_end = []
#     list_of_KKS_start = []
#     list_of_KKS_end = []
    
#     i = 0
#     n = len(int_list)
    
#     while i < n:
#         line = int_list[i]
        
#         # Шаг 1: Пытаемся найти координаты начала (3 подряд координаты)
#         coords_start = []
#         temp_i = i
        
#         # Собираем координаты из текущей строки и следующих
#         while len(coords_start) < 3 and temp_i < n:
#             line_coords = extract_coordinates_from_line(int_list[temp_i])
#             if line_coords:
#                 coords_start.extend(line_coords)
#                 temp_i += 1
#             else:
#                 # Если в текущей строке нет координат, возможно они в следующей
#                 if not coords_start:
#                     # Проверяем, может это KKS или другой текст
#                     temp_i += 1
#                 else:
#                     break
        
#         if len(coords_start) >= 3:
#             # Берём первые 3 координаты
#             list_of_axis_start.append(coords_start[:3])
#             i = temp_i
#         else:
#             i += 1
#             continue
        
#         # Шаг 2: Ищем KKS начала (до координат конца)
#         kks_start_found = []
#         temp_i = i
#         kks_collected = set()
        
#         # Ищем KKS в следующих строках, пока не найдём координаты конца
#         while temp_i < n:
#             # Проверяем, не начались ли координаты конца
#             test_coords = extract_coordinates_from_line(int_list[temp_i])
#             if test_coords:
#                 # Если нашли координаты, выходим из поиска KKS
#                 break
            
#             # Ищем KKS в текущей строке
#             kks_in_line = extract_kks_from_line(int_list[temp_i])
#             for kks in kks_in_line:
#                 if kks not in kks_collected:
#                     kks_collected.add(kks)
#                     kks_start_found.append(kks)
            
#             temp_i += 1
        
#         list_of_KKS_start.append(kks_start_found if kks_start_found else [])
#         i = temp_i
        
#         # Шаг 3: Ищем координаты конца (3 подряд координаты)
#         coords_end = []
        
#         while len(coords_end) < 3 and i < n:
#             line_coords = extract_coordinates_from_line(int_list[i])
#             if line_coords:
#                 coords_end.extend(line_coords)
#                 i += 1
#             else:
#                 if not coords_end:
#                     i += 1
#                 else:
#                     break
        
#         if len(coords_end) >= 3:
#             list_of_axis_end.append(coords_end[:3])
#         else:
#             # Если не нашли координаты конца, пропускаем
#             continue
        
#         # Шаг 4: Ищем KKS конца
#         kks_end_found = []
#         kks_collected = set()
        
#         while i < n:
#             # Проверяем, не началась ли следующая пара координат
#             test_coords = extract_coordinates_from_line(int_list[i])
#             if test_coords:
#                 break
            
#             kks_in_line = extract_kks_from_line(int_list[i])
#             for kks in kks_in_line:
#                 if kks not in kks_collected:
#                     kks_collected.add(kks)
#                     kks_end_found.append(kks)
            
#             i += 1
        
#         list_of_KKS_end.append(kks_end_found if kks_end_found else [])
    
#     return {
#         'list_of_axis_start': list_of_axis_start,
#         'list_of_axis_end': list_of_axis_end,
#         'list_of_KKS_start': list_of_KKS_start,
#         'list_of_KKS_end': list_of_KKS_end
#     }

def extract_kks_from_list(list_of_strings):
    '''
    вспомогательная функция
    поиск ККС в списке строк
    return список ККС
    '''
    list_of_kks = []
    if list_of_strings:
        for st in list_of_strings:
            if config.regular_KKS_any.search(st):
                list_of_kks.append(config.regular_KKS_any.search(st).group().strip())
    return list(set(list_of_kks))

def parse_kks_room_and_equip(list_of_KKS, all_KKS):
    '''
    all_KKS - список всех ККС объекта
    из списка list_of_KKS нахожу
    1. помещение
    2. если нет, то здание
    проверка здания или помещения на соответствие списку
    3. оборудование, которое точно не помещение и не здание
    '''
    KKS_room = ''
    KKS_equipment = ''

    # беру только ККС, не координаты
    all_all_KKS = []
    for line in all_KKS:
        all_all_KKS.append(line[0])
    all_KKS = all_all_KKS

    if list_of_KKS:
        list_of_KKS = list(set(list_of_KKS))

        for kks in list_of_KKS:
            if config.regular_KKS_room.search(kks):
                KKS_building = config.regular_KKS_building.search(kks).group().strip()
                if all_KKS and KKS_building in all_KKS:
                    KKS_room = config.regular_KKS_room.search(kks).group().strip()
                    break
            elif config.regular_KKS_building.search(kks) and not KKS_room:
                KKS_building = config.regular_KKS_building.search(kks).group().strip()
                if all_KKS and KKS_building in all_KKS:
                    KKS_room = KKS_building

        for kks in list_of_KKS:
            if (    config.regular_KKS_equipment.search(kks) 
                and config.regular_KKS_equipment.search(kks).group().strip() != KKS_room
            ):
                KKS_equipment = config.regular_KKS_equipment.search(kks).group().strip()
                break
    return KKS_room, KKS_equipment

def get_all_kks_with_XY(dir_all_KKS):
    '''
    Функция открывает реестр ККС по ссылке
    Составляет и возвращает список all_KKS [[KKS, (если есть) квадрат X, (если есть) квадрат Y], [ , , ], ...]

    Важно! на генплане Аккую ось Y горизонтальная, X вертикальная
    '''

    all_KKS = []
    try:
        rb1 = load_workbook(dir_all_KKS, data_only=True)
        sheetRead1 = rb1.active

        for rowRead in range(2, sheetRead1.max_row + 1):
            current_KKS = str(sheetRead1.cell(row=rowRead, column=1).value)
            currend_XY = str(sheetRead1.cell(row=rowRead, column=4).value)
            if current_KKS:
                current_KKS = cleanCyrFromLat(current_KKS)
                if config.regular_KKS_building.search(current_KKS):
                    # Разделяем строку "X ; Y", предполагаем, что разделитель - точка с запятой
                    parts = currend_XY.split(';')
                    if len(parts) >= 2:
                        x_coord = parts[0].strip().replace(';', '')
                        y_coord = parts[1].strip().replace(';', '')
                        all_KKS.append([current_KKS, x_coord, y_coord])
                    else:
                        # Если нет Y, записываем только X
                        all_KKS.append([current_KKS, currend_XY.strip().replace(';', ''), ''])
    except Exception as e:
        print(f'get_all_kks_with_XY: {e}')
    return all_KKS

def row_parser(input, all_KKS):
    """
    парсер строки

    Args:
        получает список из функции extract_all_text_from_dir
        [имя журнала, [ячейка 1]   [ячейка 2]  ...  [ячейка N] ]
    all_KKS - список всех ККС объекта, применяется для заполнения 9 и 14 ячейки

    Returns:
    array_row = 
                [   
                0. Журнал
                1. номер кабеля
                2. ККС	
                3. Группа
                4. Марка
                5. Сечение
                6. Диаметр
                7. Длина
                8. Трасса
                9. Откуда помещение
                10. Откуда оборудование
                11. Откуда x
                12. Откуда y
                13. Откуда z
                14. Куда помещение
                15. Куда оборудование
                16. Куда x
                17. Куда  y
                18. Куда z 
                19. Резервирование
                ]
    """

    array_row = [ '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '' ]

    list_of_error_journals = []

    if len(input) < 7:
        return []

# 0. Журнал
    array_row[0] = cleanCyrFromLat(input[0]).strip()

# 1. номер кабеля
    # ПЕРЕДЕЛАТЬ! поиск по списку через регулярку
    if not input[1]:
        list_of_error_journals.append('Ошибка номера в журнале ' + input[0])
    else:
        array_row[1] = input[1][0].replace(',', '.').strip()

# 2. ККС
    if len(input[2]) == 0:
        list_of_error_journals.append('Ошибка KKS в журнале ' + input[0])
    else:
        array_row[2] = cleanCyrFromLat(input[2][0]).strip()

# 3. Группа
    if len(input[3]) == 0:
        list_of_error_journals.append('Ошибка группы в журнале ' + input[0])
    else:
        array_row[3] = input[3][0].strip()

# 4. Марка
    for line in input[2][1:]:
            if line != None:
                line_orig = line
                line = line.replace('/', '').strip().split()[0]
                if config.regular_cableMark.search(str(line)):
                    if line == 'КППГЭнг(А)-' or line == 'КППГЭнг(A)-':                        # Заплатка устранения ошибки 'КППГЭнг(А)-  HF-T'
                        array_row[4] = 'КППГЭнг(А)-HF-T'
                    else:
                        array_row[4] = line
                else:
                    for mark in config.arrayCableMarks:
                        if mark in line_orig:
                            if line == 'КППГЭнг(А)-' or line == 'КППГЭнг(A)-':                       # Заплатка устранения ошибки 'КППГЭнг(А)-  HF-T'
                                array_row[4] = 'КППГЭнг(А)-HF-T'
                            else:
                                array_row[4] = line_orig
                            break

# 5. Сечение
    for line in input[2][1:]:
        match = config.regular_cableSection.search(line)
        if match:
            array_row[5] = match.group().replace('.', ',').replace('х', 'x').replace('×', 'x')
            break

# 7. Длина
    if len(input[-2]) == 0:
        list_of_error_journals.append('Ошибка длины в журнале ' + input[0])
    else:
        length = input[-2][0].replace(',', '.')

        #   необходимо обработать длины с * в некоторых журналах/ Если строка начинается с 1-2 звездочек и затем число (*123 или **456) или кончается звездочками -> возвращает только число

        number_pattern = r'\d*\.?\d+'
        pattern_start = rf'^\*{{1,3}}({number_pattern})$'       # Паттерн для звездочек в начале: 1-2 звездочки, затем число
        pattern_end = rf'^({number_pattern})\*{{1,3}}$'     # Паттерн для звездочек в конце: число, затем 1-2 звездочки

        match = re.match(pattern_start, length)
        if match:
            length =  match.group(1)
        match = re.match(pattern_end, length)
        if match:
            length =  match.group(1)

        if length:
            array_row[7] = length

# 8. Трасса
    if input[-1]:
        trace = ', '.join(input[-1])
        trace = trace.replace(';', ',').replace('  ', ' ').replace(',,', ',')
        if trace and trace[-1] == ',':
            trace = trace[:-1]
        if trace:
            array_row[8] = trace

# 9. Откуда помещение
# 10. Откуда оборудование
# 11. Откуда x
# 12. Откуда y
# 13. Откуда z
# 14. Куда помещение
# 15. Куда оборудование
# 16. Куда x
# 17. Куда y
# 18. Куда z

    # Я беру весь список и ищу в нем три координаты подряд - это первая и вторая тройка
    # потом я ищу 2 ККС не подряд и еще 2 ККС не подряд не одинаковых

    list_of_axis_start =    []
    list_of_axis_end =      []
    list_of_KKS_start =     []
    list_of_KKS_end =       []


# ------------ ПАРСИНГ КООРДИНАТ и ККС - 1 ВАРИАНТ ----------------------
    # если в строке имя журнала + 7 списков и в списке начало / конец может не быть ККС или координат
# ['11UQC10R005', '10BFE29GH000', 'Комплектное распредустройство 0,4 кВ /', '4002107.8', '548533.4', '0.0'],       4
# ['11UBN', '10BYA19', 'Стойка приборная автономная/', 'Instrumentation stand autonomous'],                        5
    if len(input) == 8:
        try:
            for i in range (0, len(input[4]) - 1):
                if (    (config.regular_axis_full.search(input[4][i])     or str(input[4][i])     == '-')
                    and (config.regular_axis_full.search(input[4][i + 1]) or str(input[4][i + 1]) == '-')
                    and (config.regular_axis_full.search(input[4][i + 2]) or str(input[4][i + 2]) == '-' or str(input[4][i + 2]) == '0')
                ):
                    list_of_axis_start = [  input[4][i],
                                            input[4][i + 1],
                                            input[4][i + 2]
                                        ]
                    break

            for i in range (0, len(input[5]) - 1):
                if (    (config.regular_axis_full.search(input[5][i])     or str(input[5][i])     == '-')
                    and (config.regular_axis_full.search(input[5][i + 1]) or str(input[5][i + 1]) == '-')
                    and (config.regular_axis_full.search(input[5][i + 2]) or str(input[5][i + 2]) == '-' or str(input[5][i + 2]) == '0')
                ):
                    list_of_axis_end = [ input[5][i],
                                         input[5][i + 1],
                                         input[5][i + 2]
                                    ]
                    break

            list_of_KKS_start = extract_kks_from_list(input[4])
            list_of_KKS_end = extract_kks_from_list(input[5])

        except Exception as e:
            print(f"row_parser Ошибка парсинга 1 ВАРИАНТ {array_row[1]} - {e}")


# ------------ ПАРСИНГ КООРДИНАТ и ККС - 2 ВАРИАНТ ----------------------
#   если в строке имя журнала + 9 списков и в списке 
#   начало 3, в конце 1 без координат /или/ начало 1 без координат , в конце 3

    # ['03UGF10R002', '00CMS12', 'Шкаф / Cabinet', 'Шкаф', '/', 'Cabinet', '4001294.5'],        4  начало
    # ['03UGF10R002', '00CMS12', 'Шкаф / Cabinet', 'Шкаф', '/', 'Cabinet', '549083.0'],         5  начало
    # ['03UGF10R002', '00CMS12', 'Шкаф / Cabinet', 'Шкаф', '/', 'Cabinet', '+1.500'],           6  начало
    # ['07UBG', '00CKY01', 'Комплект специального оборудования/', 'Special equipment set'],     7  конец
# или
    # ['07UBG', '00CKY01', 'Комплект специального оборудования/', 'Special equipment set'],     4  начало
    # ['03UGF10R002', '00CMS12', 'Шкаф / Cabinet', 'Шкаф', '/', 'Cabinet', '4001294.5'],        5  конец
    # ['03UGF10R002', '00CMS12', 'Шкаф / Cabinet', 'Шкаф', '/', 'Cabinet', '549083.0'],         6  конец
    # ['03UGF10R002', '00CMS12', 'Шкаф / Cabinet', 'Шкаф', '/', 'Cabinet', '+1.500'],           7  конец
    elif len(input) == 10:
        try:

            if(     (config.regular_axis_full.search(input[4][-1]) or input[4][-1] == '-')
                and (config.regular_axis_full.search(input[5][-1]) or input[5][-1] == '-')
                and (config.regular_axis_full.search(input[6][-1]) or input[6][-1] == '-')
            ):
                list_of_axis_start = [  input[4][-1],
                                        input[5][-1],
                                        input[6][-1]]

                list_of_KKS_start = extract_kks_from_list(input[4][:-1] + input[5][:-1] + input[6][:-1])
                list_of_KKS_end = extract_kks_from_list(input[7])

            elif(       (config.regular_axis_full.search(input[5][-1]) or input[5][-1] == '-')
                    and (config.regular_axis_full.search(input[6][-1]) or input[6][-1] == '-')
                    and (config.regular_axis_full.search(input[7][-1]) or input[7][-1] == '-')
            ):
                list_of_axis_end = [input[5][-1],
                                    input[6][-1],
                                    input[7][-1]]
                
                list_of_KKS_start = extract_kks_from_list(input[4])
                list_of_KKS_end = extract_kks_from_list(input[5][:-1] + input[6][:-1] + input[7][:-1])

        except Exception as e:
            print(f"row_parser Ошибка парсинга 2 ВАРИАНТ {array_row[1]} - {e}")


# ------------ ПАРСИНГ КООРДИНАТ и ККС - 3 ВАРИАНТ ----------------------
# ['00UAC10R005', '40ARA00GH001', 'Шкаф ПА блока 3', '4002107.8'],  4 начало
# ['40ARA00GH001', 'Шкаф ПА блока 3', '548533.4'],                  5
# ['40ARA00GH001', 'Шкаф ПА блока 3',  '0.0'],                      6
# ['03UGF10R002', '4001294.5'],                                     7 конец
# ['00CMS12', 'Шкаф / Cabinet', '549083.0'],                        8
# ['00CMS12', 'Шкаф / Cabinet',  '+1.500'],                         9
    elif len(input) == 12:
        try:
            if(     (config.regular_axis_full.search(input[4][-1]) or input[4][-1] == '-')
                and (config.regular_axis_full.search(input[5][-1]) or input[5][-1] == '-')
                and (config.regular_axis_full.search(input[6][-1]) or input[6][-1] == '-')
            ):
                list_of_axis_start = [  input[4][-1],
                                        input[5][-1],
                                        input[6][-1]]

                list_of_KKS_start = extract_kks_from_list(input[4][:-1] + input[5][:-1] + input[6][:-1])
                
            if(     (config.regular_axis_full.search(input[7][-1]) or input[7][-1] == '-')
                and (config.regular_axis_full.search(input[8][-1]) or input[8][-1] == '-')
                and (config.regular_axis_full.search(input[9][-1]) or input[9][-1] == '-')
            ):
                list_of_axis_end = [input[7][-1],
                                    input[8][-1],
                                    input[9][-1]]

                list_of_KKS_end = extract_kks_from_list(input[7][:-1] + input[8][:-1] + input[9][:-1])

                # -------------   есть вариант, когда координата в начале строки -------------------------

            if(     (config.regular_axis_full.search(input[4][0]) or input[4][0] == '-')
                and (config.regular_axis_full.search(input[5][0]) or input[5][0] == '-')
                and (config.regular_axis_full.search(input[6][0]) or input[6][0] == '-')
            ):
                list_of_axis_start = [  input[4][0],
                                        input[5][0],
                                        input[6][0]]

                list_of_KKS_start = extract_kks_from_list(input[4][1:] + input[5][1:] + input[6][1:])

            if(     (config.regular_axis_full.search(input[7][0]) or input[7][0] == '-')
                and (config.regular_axis_full.search(input[8][0]) or input[8][0] == '-')
                and (config.regular_axis_full.search(input[9][0]) or input[9][0] == '-')
            ):
                list_of_axis_end = [input[7][0],
                                    input[8][0],
                                    input[9][0]]

                list_of_KKS_end = extract_kks_from_list(input[7][1:] + input[8][1:] + input[9][1:])

        except Exception as e:
            print(f"row_parser Ошибка парсинга 3 ВАРИАНТ {array_row[1]} - {e}")


# ------------ ПАРСИНГ КООРДИНАТ и ККС - 4 ВАРИАНТ ----------------------
            # координаты расположены подряд в отдельных ячейках
    elif len(input) > 12:

# собираю в общий промежуточный список всю информацию о начале и конце трассы из ячеек между группой и длиной
        int_list = []  # intermediate list - 
        for cell in input[4:-2]:
            if cell:
                int_list.extend(cell)

        start_last_index = -1
        as_f = -1
        as_l = -1
        ae_f = -1
        ae_l = -1

        try:
# просматриваю список сначала и беру первые три подряд кординаты - это начало
            for i in range (0, len(int_list)):
                if (
                        (config.regular_axis_full.search(int_list[i])     or int_list[i]     == '-')
                    and (config.regular_axis_full.search(int_list[i + 1]) or int_list[i + 1] == '-')
                    and (config.regular_axis_full.search(int_list[i + 2]) or int_list[i + 2] == '-')
                ):
                    list_of_axis_start = [  int_list[i],
                                            int_list[i + 1],
                                            int_list[i + 2]
                                            ]
                    start_last_index = i + 2
                    as_f, as_l = find_first_and_last_sublist_index(int_list, list_of_axis_start)  # axis start_first, axis_start last
                    break

# просматриваю список с конца и беру первые три подряд координаты - это конец   
            for i in range(len(int_list) - 1, -1, -1):
                if (
                        (config.regular_axis_full.search(int_list[i])     or int_list[i]     == '-')
                    and (config.regular_axis_full.search(int_list[i - 1]) or int_list[i - 1] == '-')
                    and (config.regular_axis_full.search(int_list[i - 2]) or int_list[i - 2] == '-')
                    and start_last_index > 0
                    and start_last_index < i - 2
                ):
                    list_of_axis_end = [int_list[i - 2],
                                        int_list[i - 1],
                                        int_list[i]
                                        ]
                    ae_f, ae_l = find_first_and_last_sublist_index(int_list, list_of_axis_end)  # axis end_first, axis end_last
                    break
        except Exception as e:
            print(f"row_parser Ошибка парсинга 4 ВАРИАНТ (поиск координат) {array_row[1]} - {e}")

# ----ищу KKS - 1 комбинация
# ---- порядок считывания: XYZ начала ---> ККС начала ---> XYZ конца ---> ККС конца
        if (
                len(list_of_axis_start) > 0
            and len(list_of_axis_end) > 0
            and as_f == 0
            and ae_f > 3
        ):

            try:
# 1.1. ищу ККС начала между первой и последней координатой                
                for i in range (3, ae_f):
                    kks = ''
                    if config.regular_KKS_any.search(int_list[i]):
                        kks = config.regular_KKS_any.search(int_list[i]).group().strip()
                        if kks not in list_of_KKS_start:
                            list_of_KKS_start.append(kks)
                            if len(list_of_KKS_start) == 3:
                                break
# 1.2. ищу ККС конца между индексом последней координаты конца и концом массива
                for i in range (ae_l, len(int_list)):
                    kks = ''
                    if config.regular_KKS_any.search(int_list[i]):
                        kks = config.regular_KKS_any.search(int_list[i]).group().strip()
                        if kks and kks not in list_of_KKS_end:
                            list_of_KKS_end.append(kks)
                            if len(list_of_KKS_end) == 3:
                                break
            except Exception as e:
                print(f"row_parser Ошибка парсинга 4 ВАРИАНТ (поиск ККС 1 комбинация) {array_row[1]} - {e}")

# ----- ищу KKS - 2 комбинация
# ----- порядок считывания: ККС начала ---> XYZ начала ---> ККС конца ---> XYZ конца
        if (
            len(list_of_axis_start) > 0
            and len(list_of_axis_end) > 0
            and as_f > 2
            and as_l < ae_f
        ):
            try:
# 2.1. ищу ККС начала между 0 и индексом первой координаты
                for i in range (0, as_f): 
                    kks = ''          
                    if config.regular_KKS_any.search(int_list[i]):
                        kks = config.regular_KKS_any.search(int_list[i]).group().strip()
                        if kks not in list_of_KKS_start:
                            list_of_KKS_start.append(kks)
                            if len(list_of_KKS_start) == 3:
                                break
# 2.2. ищу ККС конца между координатами начала и координатами конца
                for i in range (as_l, ae_f):
                    kks = ''
                    if config.regular_KKS_any.search(int_list[i]):
                        kks = config.regular_KKS_any.search(int_list[i]).group().strip()
                        if kks not in list_of_KKS_end:
                            list_of_KKS_end.append(kks)
                            if len(list_of_KKS_end) == 3:
                                break
            except Exception as e:
                print(f"row_parser Ошибка парсинга 4 ВАРИАНТ (поиск ККС 2 комбинация) {array_row[1]} - {e}")

# ВЫВОД ДАННЫХ
    if len(list_of_axis_start) == 3:
        array_row[11] = list_of_axis_start[0].replace('+', '').strip()
        array_row[12] = list_of_axis_start[1].replace('+', '').strip()
        array_row[13] = list_of_axis_start[2].replace('+', '').strip()
    
    if len(list_of_axis_end) == 3:
        array_row[16] = list_of_axis_end[0].replace('+', '').strip()
        array_row[17] = list_of_axis_end[1].replace('+', '').strip()
        array_row[18] = list_of_axis_end[2].replace('+', '').strip()

    rs, es = parse_kks_room_and_equip(list_of_KKS_start, all_KKS)
    array_row[9] = str(rs)
    array_row[10] = str(es)

    rr, ee = parse_kks_room_and_equip(list_of_KKS_end, all_KKS)
    array_row[14] = str(rr)
    array_row[15] = str(ee)

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    # 19. Резервирование
# TODO ПЕРЕДЕЛАТЬ! поиск по input[0] через регулярку 
    cell = input[3]    
    pattern = re.compile(r'резерв', re.IGNORECASE)
    for line in cell:
        match = pattern.search(line)
        if match:
            array_row[19] = 'Резерв'

    # print(f'row_parser отработал')
    return array_row

# ----- Работа с журналами

def base_master_start(dir_journals, dir_in, dir_all_KKS):
    """
    0. Получаю список всех ККС зданий объекта из таблицы по адресу dir_all_KKS
    1. Нахожу текущую версию кабельного журнала ЕСЛИ он лежит в этой же папке функцией current_version_finder
    2. Получаю список журналов в формате docX из указанной папки
    3. Создаю новую базу со следующей версией
  
    если база обновляется, то
    4. открываем предыдущую версию базы, последнюю на момент запуска программы
    5. переписываем все из старой версии базы в новую за исключением журналов из списка journals
        TODO проверить на полноту данных и выделить строку желтым

    6.1 пожурнально прогоняем список journals через convert_numbering_to_text (замена автонумерации),
        затем extract_all_text_from_docx, получаю список строк
        TODO  Если функция падает - переношу имя журнала в список list_of_troubles
    6.2 каждую строку (список) полученного списка прогоняю через row_parser, получаю строку (список) с данными
        Если функция падает - переношу имя журнала в список list_of_troubles
    6.3 каждую полученную строку (список) проверяю на полноту содержимого
        TODO Если чего-то не хватает - ? переношу имя журнала в список list_of_void
    6.4 Если журнал не в списке list_of_troubles, то записываю в новую базу
            журнал
            строку (список)
            дату новой базы
            версию новой базы
            TODO проверить на полноту данных и выделить строку желтым
    7. сохраняю базу по указанному адресу
    8. перемещаю журналы из list of troubles в отдельную директорию

    Args:
        dir_journals: папка с журналами 
        dir_in: папка с базой 
    """

    journals = []  # список  журналов
    list_of_troubles = set() # множество проблемных журналов
    list_of_void = set()

    checklist = [
                3, #'Группа'
                7, #'Длина'
                9, #'Откуда помещение'
                10, #'Откуда оборудование'
                11, #'Откуда x'
                12, #'Откуда y'
                13, #'Откуда z'
                14, #'Куда помещение'
                15, #'Куда оборудование'
                16, #'Куда x'
                17, #'Куда y'
                18 #'Куда z'
                    ]
# 0
    all_KKS = get_all_kks_with_XY(dir_all_KKS)
    
    # try:
    #     rb1 = load_workbook(dir_all_KKS, data_only=True) # открывается предыдущая версию базы
    #     sheetRead1 = rb1.active

    #     for rowRead in range (2, sheetRead1.max_row + 1):
    #         current_KKS = str(sheetRead1.cell(row = rowRead, column = 1).value)
    #         if current_KKS:
    #             current_KKS = cleanCyrFromLat(current_KKS)
    #             if config.regular_KKS_building.search(current_KKS):
    #                 all_KKS.append(current_KKS)
    # except Exception as e:
    #     print(f'не шмогла, {e}')

# 1
    b_name = 'Cable base ver.'
    current_version = current_version_finder(dir_in, b_name)  # текущая версия базы, та, которую мы ДЕЛАЕМ сейчас

# 2
    journals = take_all_docx_from_dir(dir_journals)
    # print(f'список журналов из функции take_all_docx_from_dir {journals}')
    journals_names = []  # список имен журналов
    if journals:
        for j in journals:
            journals_names.append(j.stem) # .stem - имена без расширений

# 3.
    wb = openpyxl.Workbook()
    sheetWrite = wb.active
    sheetWrite.title = "База данных вер. " + str(current_version)

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # пишу заголовки в новый файл - номер, название, ширина столбца
    heads = [   [0, 'Журнал', 40], 
                [1, 'Номер кабеля', 9], 
                [2, 'ККС', 21], 
                [3, 'Группа', 7], 
                [4, 'Марка', 22], 
                [5, 'Сечение', 10], 
                [6, 'Диаметр', 8], 
                [7, 'Длина', 6], 
                [8, 'Трасса', 20], 
                [9, 'Откуда помещение', 14], 
                [10, 'Откуда оборудование', 20], 
                [11, 'Откуда x', 11], 
                [12, 'Откуда y', 11], 
                [13, 'Откуда z', 11], 
                [14, 'Куда помещение', 14], 
                [15, 'Куда оборудование', 20], 
                [16, 'Куда x', 11], 
                [17, 'Куда y', 11], 
                [18, 'Куда z', 11], 
                [19, 'Резервирование', 8], 
                [20, 'Дата добавления в базу', 10], 
                [21, 'Версия базы', 7],
                [22, 'id кабеля', 8],
                [23, 'Номер ревизии', 6],
                [24, 'Дата ревизии', 10],
                [25, 'Класс безопасности', 7],
                [26, 'Статус прокладки', 10],
                [27, 'Источник информации ВК/СУПИР', 8],
                [28, 'Статус объединения', 10],
                [29, 'Raw', 25],
                [30, 'min длина', 10],
                [31, 'Длина меньше min', 10]
           ]
    
    for head in heads:
        c = head[0] + 1
        sheetWrite.cell(row = 1, column = c, value = f'{head[0]+ 1}. {head[1]}')
        sheetWrite.column_dimensions[get_column_letter(c)].width = head[2]

    sheetWrite.freeze_panes = 'A2' # закрепление первой строки

    # Строка с текущей датой
    today = date.today()
    date_string = today.strftime("%Y.%m.%d")  # '2026.03.09'

    rowWrite = 2        # переменная - строка записи в excel
# 4 Самая старая версия
    # if current_version > 1:     # если изменяется существующая  база
    #     rb = load_workbook(dir_in + '/' + b_name + str(current_version - 1) + '.xlsx', data_only=True) # открывается предыдущая версию базы
    #     sheetRead = rb.active

        # for rowRead in range (2, sheetRead.max_row + 1):
        #     if str(sheetRead.cell(row = rowRead, column = 1).value) not in journals_names: # если журнал не в списке
        #         for col in range(1, sheetRead.max_column + 1):
        #             sheetWrite.cell(row = rowWrite, column = col, value = sheetRead.cell(row = rowRead, column = col).value)
        #         rowWrite += 1



# 4  первая версия дипсика, добавил прогресс-бар, но по-прежнему очень медленно работает
#     if current_version > 1:
#         rb = load_workbook(dir_in + '/' + b_name + str(current_version - 1) + '.xlsx', data_only=True)
#         sheetRead = rb.active
        
#         from tqdm import tqdm
        
#         total_rows = sheetRead.max_row - 1  # количество строк данных (без заголовка)
#         copied_count = 0
# # 5
#         print(f"\n📦 Перенос данных из предыдущей версии (Cable base ver.{current_version - 1}.xlsx)...")
#         with tqdm(total=total_rows, desc="Копирование строк", unit="строка") as pbar:
#             for rowRead in range(2, sheetRead.max_row + 1):
#                 journal_name = str(sheetRead.cell(row=rowRead, column=1).value)
#                 if journal_name not in journals_names:
#                     for col in range(1, sheetRead.max_column + 1):
#                         sheetWrite.cell(row=rowWrite, column=col, value=sheetRead.cell(row=rowRead, column=col).value)
#                     rowWrite += 1
#                     copied_count += 1
#                 pbar.update(1)
        
#         print(f"✅ Скопировано {copied_count} строк, пропущено (обработано сейчас): {total_rows - copied_count}")

# 4
    copied_count = 0
    if current_version > 1:
        rb = load_workbook(dir_in + '/' + b_name + str(current_version - 1) + '.xlsx', data_only=True)
        sheetRead = rb.active
        
        # Создаём список имён журналов из колонки A (без заголовка)
        journal_names_from_old = []
        rows_to_copy = []
        
        for idx, row in enumerate(sheetRead.iter_rows(min_row=2, values_only=True), start=2):
            journal_name = str(row[0]) if row[0] else ''
            journal_names_from_old.append(journal_name)
            if journal_name not in journals_names:
                rows_to_copy.append((idx, row))
        
        from tqdm import tqdm
# 5        
        print(f"\n📦 Перенос данных из предыдущей версии...")
        with tqdm(total=len(rows_to_copy), desc="Копирование строк", unit="строка") as pbar:
            for idx, row_data in rows_to_copy:
                for col_idx, value in enumerate(row_data, start=1):
                    sheetWrite.cell(row=rowWrite, column=col_idx, value=value)
                rowWrite += 1
                copied_count += 1
                pbar.update(1)




# 6.1
    if journals:
        for ii, journal in enumerate(journals):
            log_of_warnings = set()
            # print(f'({ii + 1} / {len(journals)})  {journal.stem}', end=" ---> ")  
            print(f'({ii + 1} / {len(journals)})  {journal.stem}')

            j_raw_content = None
            # print(f'base_master 6.1 journal - {journal}')
            try:
                # journal = convert_numbering_to_text(journal)  # TEST
                j_raw_content = extract_all_text_from_docx(Document(journal)) # список необработанных строк (списков) из текущего кабельного журнала
                # print(f'base_master 6.1 - {len(j_raw_content)} строк')
            except Exception as e:
                print(f"Ошибка извлечения данных из журнала {journal.stem}: {e} (ошибка функции base_master 6.1 )")
#                 print(f'Пробую исправить ошибку нумерации', end=" ---> ")
# # 6.1.1
#                 try:
#                     journal = convert_numbering_to_text(journal)
#                     j_raw_content = extract_all_text_from_docx(Document(journal)) # список необработанных строк (списков) из текущего кабельного журнала
#                     print(f'base_master 6.1.1 - {len(j_raw_content)} строк')
#                 except Exception as e:
#                     print(f"Все равно ошибка извлечения данных из журнала {journal.stem}: {e} (ошибка функции base_master 6.1.1 )")
#                     list_of_troubles.add(journal.stem)
# 6.2               
            if j_raw_content:
                for raw_row in j_raw_content:
                    if len(raw_row) > 6:
                        raw_row.insert(0, journal.stem)
                        current_row = []
                        try:
                            current_row = row_parser(raw_row, all_KKS) # Обработанная строка журнала
                        except Exception as e:
                            print(f"Ошибка обработки данных из журнала {journal.stem} в строке {raw_row[1]}: {e} (ошибка функции base_master 6.2 )")
                            log_of_warnings.add(str(f"Ошибка обработки данных из журнала {journal.stem} в строке {raw_row[1]}: {e} (ошибка функции base_master 6.2 )"))
                            list_of_troubles.add(journal.stem)

# 6.3               
                        if current_row:
                            for i in checklist:
                                if not current_row[i]:
                                    log_of_warnings.add(str(f"Недостаточность данных в журнале {journal.stem} в строке {current_row[1]}"))
                                    list_of_void.add(journal.stem)
# 6.4
                            if current_row[0] not in list_of_troubles:
                                current_row.append(date_string)
                                current_row.append(int(current_version))

                                for indx, current_cell in enumerate(current_row):
                                    col = indx + 1
                                    # запись в эксель группы как число и длины как число
                                    if indx == 3: # группа
                                        try:
                                            sheetWrite.cell(row = rowWrite, column = col, value = int(current_cell))
                                        except Exception as e:
                                            # print(f"Ошибка группы в журнале {journal.stem} в кабеле {current_row[1]}: {e}. Записано в журнал как есть")
                                            sheetWrite.cell(row = rowWrite, column = col, value = str(current_cell))

                                    elif indx in [7, 11, 12, 13, 16, 17, 18, 21]: # длина и координаты
                                        try:
                                            sheetWrite.cell(row = rowWrite, column = col, value = float(current_cell))
                                        except Exception as e:
                                            # print(f"Ошибка длины в журнале {journal.stem} в кабеле {current_row[1]}: {e}. Записано в журнал как есть")
                                            sheetWrite.cell(row = rowWrite, column = col, value = str(current_cell))
                                    else:
                                        sheetWrite.cell(row = rowWrite, column = col, value = str(current_cell))

                                    # пишу raw
                                    sheetWrite.cell(row = rowWrite, column = 30, value = str(raw_row))
# подсчет минимальной длины кабеля
                                try:
                                    length = float(current_row[7])

                                    xyz_from = [float(current_row[11]), 
                                                float(current_row[12]),
                                                float(current_row[13])]
                                    
                                    xyz_to =   [float(current_row[16]), 
                                                float(current_row[17]),
                                                float(current_row[18])]
                                    
                                    min_len = round(abs(xyz_from[0] - xyz_to[0]) + abs(xyz_from[1] - xyz_to[1]) + abs(xyz_from[2] - xyz_to[2]))
                                    sheetWrite.cell(row = rowWrite, column = 31, value = min_len)
                                    
                                    if length < min_len:
                                        sheetWrite.cell(row = rowWrite, column = 32, value = 'ДА')
                                except Exception as e:
                                    donothing = True                             

                                rowWrite += 1
            
            # if log_of_warnings:
            #     for w in sorted(log_of_warnings):
            #         print(w)

# TODO на втором листе вести лог изменений
# 10
    sheetWrite.auto_filter.ref = sheetWrite.dimensions
    # Этот код включает автофильтр в Excel-таблице на весь диапазон данных, 
    # который в данный момент есть на листе

    wb.save(dir_in + '/' + b_name + str(current_version)+ '.xlsx')  # Сохраняем файл на диск
    print('Base done!')

'''
# 11
    for journal in list_of_troubles:
        move_file (dir_journals + '/' + journal + '.docx', dir_in + '/troubles')
    for journal in list_of_void:
        move_file (dir_journals + '/' + journal + '.docx', dir_in + '/void_fields')

# Укажите полный путь к файлу
    file_path = dir_in + '\\' + 'log.txt'  # для Windows
# Открываем файл для записи и сохраняем текст

    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(log)

    print(f"Файл сохранён: {file_path}")
'''