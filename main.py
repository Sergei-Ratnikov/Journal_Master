import time
from base_master import base_master_start

start_time = time.time()

base_master_start('C:/Users/User/Documents/Python/WordWorks/Исходники', 'C:/Users/User/Documents/Python/WordWorks/Исходники')

end_time = time.time()    
duration = end_time - start_time
print(f"Время выполнения: {duration:.6f} секунд")