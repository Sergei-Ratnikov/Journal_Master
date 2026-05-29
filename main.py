# main.py
import time
from excel_utils import build_cable_database

def main():
    """
    Консольный запуск программы.
    Пути к папкам запрашиваются у пользователя.
    """
    print("\n" + "=" * 60)
    print("КАБЕЛЬНЫЙ ЖУРНАЛ - ПАРСЕР И СОЗДАНИЕ БАЗЫ ДАННЫХ")
    print("=" * 60)
    
    # Запрашиваем пути у пользователя
    journals_dir = input("Введите путь к папке с журналами: ").strip()
    output_dir = input("Введите путь к папке для сохранения базы: ").strip()
    
    # Убираем кавычки, если пользователь их ввёл
    journals_dir = journals_dir.strip('"').strip("'")
    output_dir = output_dir.strip('"').strip("'")
    
    print(f"\nПапка с журналами: {journals_dir}")
    print(f"Папка для сохранения: {output_dir}")
    print("-" * 60)
    
    start_time = time.time()
    
    try:
        build_cable_database(journals_dir, output_dir)
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        return
    
    end_time = time.time()
    duration = end_time - start_time
    print(f"\n⏱️ Время выполнения: {duration:.2f} секунд")
    print("\n🏁 Готово!")

if __name__ == "__main__":
    main()