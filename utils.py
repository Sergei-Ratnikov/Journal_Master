import os
import shutil
import win32com.client
from docx import Document
import re

def cleanCyrFromLat(lineIn):
    '''
    замена русских букв на латинские
    '''
    try:
        return str(lineIn).replace('А', 'A').replace('В', 'B').replace('С', 'C').replace('Е', 'E').replace('Н', 'H').replace('К', 'K').replace('М', 'M').replace('О', 'O').replace('Р', 'P').replace('Т', 'T').replace('Х', 'X').replace(' ', '').replace(',', '.')
    except Exception as e:
        raise

def contains_latin_ignore_case(text):
    '''
    Проверяет наличие латинских букв (регистронезависимо)
    '''
    try:
        return bool(re.search(r'[a-z]', str(text), re.IGNORECASE))
    except Exception as e:
        raise

def reverse_trace(text):
    '''
    Реверс трассы
    '''
    try:
        traces = [item.strip() for item in str(text).split(',')]
        traces.reverse()
        return ', '.join(traces)
    except Exception as e:
        raise

def move_file(source_path, destination_dir):
    """
    Переносит файл из source_path в папку destination_dir
    
    Args:
        source_path: полный путь к файлу (например, "C:/files/doc.docx")
        destination_dir: путь к папке назначения (например, "C:/archive")
    """
    # Проверяем, существует ли исходный файл
    if not os.path.exists(source_path):
        print(f"Файл не найден: {source_path}")
        return False

    try:    
        # Создаём папку назначения, если её нет
        os.makedirs(destination_dir, exist_ok=True)
        
        # Получаем имя файла из полного пути
        filename = os.path.basename(source_path)
        
        # Формируем полный путь назначения
        destination_path = os.path.join(destination_dir, filename)
    
        # Переносим файл
        shutil.move(source_path, destination_path)
        print(f"Файл перемещён: {source_path} → {destination_path}")
        return True
    except Exception as e:
        print(f"Ошибка при переносе: {e}")
        return False

def convert_doc_to_docx(doc_path):
    '''
        Конвертирует .doc в .docx
    '''
    word = win32com.client.Dispatch("Word.Application")     #создаёт COM-объект, который запускает Microsoft Word в фоновом режиме.    word — теперь это объект, через который можно управлять Word'ом.
    word.Visible = False                                    # Word будет работать в фоновом режиме, без отображения окна
    try:
        doc = word.Documents.Open(os.path.abspath(doc_path))    # преобразует относительный путь в абсолютный (например, "file.doc" → "C:\Users\...\file.doc").
        docx_path = doc_path.replace('.doc', '.docx')
        doc.SaveAs(docx_path, FileFormat=16)
                                                                # сохраняет документ в новом формате.
                                                                # FileFormat=16 — это числовой код, соответствующий формату .docx.
                                                                # 16 = wdFormatDocumentDefault — стандартный формат Word (.docx). Другие форматы: 0 = .doc (старый), 17 = .pdf, и т.д.
        doc.Close()
        word.Quit()
        return docx_path
    except Exception as e:
        raise

def make_hashable(item):
    """
    вспомогательная функция

    Рекурсивно преобразует списки в кортежи для возможности хеширования"""
    if isinstance(item, list):    # является ли item списком
        return tuple(make_hashable(x) for x in item)
    return item

def is_subset_with_lists(list1, list2):
    """
    вспомогательная функция is_subset_with_lists(list1, list2)
    Проверяет, является ли list1 подмножеством list2, где элементы могут быть списками
    list1 = [[1, 2], 3, [4, [5, 6]]]
    list2 = [[1, 2], 3, 4, [4, [5, 6]], [7, 8]]
    print(is_subset_with_lists(list1, list2))  # True
    """
    # Преобразуем оба списка в хешируемые кортежи
    set1 = {make_hashable(x) for x in list1}
    set2 = {make_hashable(x) for x in list2}
    
    return set1.issubset(set2)

def all_deep_empty(lst):
    """
    вспомогательная функция

    Рекурсивно проверяет, пусты ли все вложенные списки
    выводит true если все вложенные списки пустые
    """
    if not isinstance(lst, list):
        return False
    
    if not lst:  # пустой список
        return True
    
    # Проверяем каждый элемент
    return all(all_deep_empty(item) for item in lst)

def get_mismatch_indices(list1, list2):
    """
    вспомогательная функция

    Возвращает индексы элементов, которые различаются в двух списках
    Списки должны быть одинаковой длины
    """
    if len(list1) != len(list2):
        raise ValueError("Списки должны быть одинаковой длины")
    
    mismatch_indices = []
    for i, (a, b) in enumerate(zip(list1, list2)):
        if a != b:
            mismatch_indices.append(i)
    
    return mismatch_indices