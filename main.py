import time
from base_master import base_master

start_time = time.time()

# base_master_test('Word Works')
# base_master('Word Works/to base', 'Word Works/to base')
# base_master('WordWorks/troubles', 'WordWorks/troubles')
base_master('WordWorks/test', 'WordWorks/test', 'C:/Users/User/Documents/Python/JournalMaster/KKS.xlsx')

end_time = time.time()    
duration = end_time - start_time
print(f"Время выполнения: {duration:.6f} секунд")