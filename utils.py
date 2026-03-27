import os
import shutil
import win32com.client
from docx import Document
import re

def cleanCyrFromLat(lineIn):
    '''
    замена русских букв на латинские
    '''
    return str(lineIn).replace('А', 'A').replace('В', 'B').replace('С', 'C').replace('Е', 'E').replace('Н', 'H').replace('К', 'K').replace('М', 'M').replace('О', 'O').replace('Р', 'P').replace('Т', 'T').replace('Х', 'X').replace(' ', '').replace(',', '.')


def contains_latin_ignore_case(text):
    '''
    Проверяет наличие латинских букв (регистронезависимо)
    '''
    return bool(re.search(r'[a-z]', text, re.IGNORECASE))


def reverse_trace(text):
    '''
    Реверс трассы
    '''
    traces = [item.strip() for item in text.split(',')]
    traces.reverse()
    return ', '.join(traces)


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
    
    # Создаём папку назначения, если её нет
    os.makedirs(destination_dir, exist_ok=True)
    
    # Получаем имя файла из полного пути
    filename = os.path.basename(source_path)
    
    # Формируем полный путь назначения
    destination_path = os.path.join(destination_dir, filename)
    
    # Переносим файл
    try:
        shutil.move(source_path, destination_path)
        print(f"Файл перемещён: {source_path} → {destination_path}")
        return True
    except Exception as e:
        print(f"Ошибка при переносе: {e}")
        return False


def convert_doc_to_docx(doc_path):
    '''
        Конвертирует .doc в .docx (только Windows!)
    '''
    word = win32com.client.Dispatch("Word.Application")     #создаёт COM-объект, который запускает Microsoft Word в фоновом режиме.    word — теперь это объект, через который можно управлять Word'ом.
    word.Visible = False                                    # Word будет работать в фоновом режиме, без отображения окна
    
    doc = word.Documents.Open(os.path.abspath(doc_path))    # преобразует относительный путь в абсолютный (например, "file.doc" → "C:\Users\...\file.doc").
    docx_path = doc_path.replace('.doc', '.docx')
    doc.SaveAs(docx_path, FileFormat=16)
                                                            # сохраняет документ в новом формате.
                                                            # FileFormat=16 — это числовой код, соответствующий формату .docx.
                                                            # 16 = wdFormatDocumentDefault — стандартный формат Word (.docx). Другие форматы: 0 = .doc (старый), 17 = .pdf, и т.д.
    doc.Close()
    word.Quit()
    return docx_path