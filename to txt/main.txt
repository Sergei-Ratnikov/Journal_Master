# main.py
import time
from excel_utils import build_cable_database

start_time = time.time()

build_cable_database(
    journals_dir='C:/Python/WordWorks/Well',
    output_dir='C:/Python/WordWorks/Well'
)

end_time = time.time()    
duration = end_time - start_time
print(f"Время выполнения: {duration:.6f} секунд")