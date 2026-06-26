import ezdxf
import openpyxl
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread
import time

class DXFReplacerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DXF Текст-Заменитель")
        self.root.geometry("700x550")
        self.root.resizable(True, True)
        
        # Переменные для хранения путей к файлам
        self.excel_file = tk.StringVar()
        self.dxf_file = tk.StringVar()
        self.output_file = tk.StringVar()
        
        # Статус и прогресс
        self.status_var = tk.StringVar(value="Готов к работе")
        self.progress_var = tk.DoubleVar(value=0)
        
        # Создаем интерфейс
        self.create_widgets()
        
        # Центрируем окно
        self.center_window()
        
    def center_window(self):
        """Центрирует окно на экране"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Создает все виджеты интерфейса"""
        
        # Главный фрейм с отступами
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        title_label = ttk.Label(
            main_frame, 
            text="🔧 Замена текста в DXF файлах", 
            font=('Arial', 16, 'bold')
        )
        title_label.pack(pady=(0, 20))
        
        # Фрейм для выбора файлов
        files_frame = ttk.LabelFrame(main_frame, text="Выбор файлов", padding="10")
        files_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Excel файл
        ttk.Label(files_frame, text="📊 Excel файл с заменами:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(files_frame, textvariable=self.excel_file, width=50).grid(row=0, column=1, padx=(10, 5), pady=5)
        ttk.Button(files_frame, text="Обзор...", command=self.browse_excel).grid(row=0, column=2, pady=5)
        
        # DXF файл
        ttk.Label(files_frame, text="📐 DXF файл для обработки:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(files_frame, textvariable=self.dxf_file, width=50).grid(row=1, column=1, padx=(10, 5), pady=5)
        ttk.Button(files_frame, text="Обзор...", command=self.browse_dxf).grid(row=1, column=2, pady=5)
        
        # Выходной файл
        ttk.Label(files_frame, text="💾 Сохранить как:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(files_frame, textvariable=self.output_file, width=50).grid(row=2, column=1, padx=(10, 5), pady=5)
        ttk.Button(files_frame, text="Обзор...", command=self.browse_output).grid(row=2, column=2, pady=5)
        
        # Настройки
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Цвет
        ttk.Label(settings_frame, text="🎨 Цвет замененного текста:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.color_var = tk.StringVar(value="2")  # По умолчанию желтый
        
        # Создаем фрейм для радиокнопок
        color_frame = ttk.Frame(settings_frame)
        color_frame.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        colors = [
            ("Желтый", "2"),
            ("Красный", "1"),
            ("Зеленый", "3"),
            ("Голубой", "4"),
            ("Синий", "5"),
            ("Фиолетовый", "6"),
            ("Белый", "7"),
            ("Оранжевый", "30"),
        ]
        
        # Разбиваем на две строки, чтобы не было слишком широко
        for i, (text, value) in enumerate(colors):
            row = i // 4
            col = i % 4
            ttk.Radiobutton(
                color_frame, 
                text=text, 
                value=value, 
                variable=self.color_var
            ).grid(row=row, column=col, padx=(0, 15), pady=2, sticky=tk.W)
        
        # Кнопка запуска
        self.run_button = ttk.Button(
            main_frame,
            text="🚀 ЗАПУСТИТЬ ЗАМЕНУ",
            command=self.run_replacement
        )
        self.run_button.pack(pady=15)
        
        # Прогресс-бар
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode='determinate'
        )
        self.progress_bar.pack(pady=(0, 10))
        
        # Статус
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(status_frame, text="Статус:").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(status_frame, textvariable=self.status_var, font=('Arial', 9)).pack(side=tk.LEFT)
        
        # Текстовое поле для лога
        log_frame = ttk.LabelFrame(main_frame, text="Лог операций", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Создаем текстовое поле и скроллбар
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, font=('Courier', 9))
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        
        # Исправлено: yscrollcommand вместо ystickcommand
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Размещаем элементы
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Начальное сообщение в логе
        self.log("Программа запущена. Выберите файлы и нажмите 'Запустить замену'.")
        
    def log(self, message):
        """Добавляет сообщение в лог"""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def browse_excel(self):
        """Выбор Excel файла"""
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл с заменами",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file.set(filename)
            self.log(f"Выбран Excel файл: {os.path.basename(filename)}")
    
    def browse_dxf(self):
        """Выбор DXF файла"""
        filename = filedialog.askopenfilename(
            title="Выберите DXF файл",
            filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")]
        )
        if filename:
            self.dxf_file.set(filename)
            # Автоматически предлагаем имя для выходного файла
            dirname = os.path.dirname(filename)
            basename = os.path.basename(filename)
            name, ext = os.path.splitext(basename)
            self.output_file.set(os.path.join(dirname, f"{name}_modified{ext}"))
            self.log(f"Выбран DXF файл: {basename}")
    
    def browse_output(self):
        """Выбор выходного файла"""
        filename = filedialog.asksaveasfilename(
            title="Сохранить результат как",
            defaultextension=".dxf",
            filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
            self.log(f"Выходной файл: {os.path.basename(filename)}")
    
    def validate_files(self):
        """Проверяет, что все файлы выбраны и существуют"""
        if not self.excel_file.get():
            messagebox.showerror("Ошибка", "Выберите Excel файл с заменами!")
            return False
        
        if not os.path.exists(self.excel_file.get()):
            messagebox.showerror("Ошибка", f"Excel файл не найден:\n{self.excel_file.get()}")
            return False
        
        if not self.dxf_file.get():
            messagebox.showerror("Ошибка", "Выберите DXF файл для обработки!")
            return False
        
        if not os.path.exists(self.dxf_file.get()):
            messagebox.showerror("Ошибка", f"DXF файл не найден:\n{self.dxf_file.get()}")
            return False
        
        if not self.output_file.get():
            messagebox.showerror("Ошибка", "Укажите путь для сохранения результата!")
            return False
        
        # Проверяем, что выходной файл не совпадает с входным
        if self.output_file.get() == self.dxf_file.get():
            reply = messagebox.askyesno(
                "Предупреждение",
                "Выходной файл совпадает с входным! Это перезапишет исходный файл.\nПродолжить?"
            )
            if not reply:
                return False
        
        return True
    
    def run_replacement(self):
        """Запускает процесс замены в отдельном потоке"""
        if not self.validate_files():
            return
        
        # Блокируем кнопку запуска
        self.run_button.config(state=tk.DISABLED)
        self.status_var.set("Идет обработка...")
        self.progress_var.set(0)
        self.log("\n" + "="*50)
        self.log("Начинается процесс замены...")
        
        # Запускаем в отдельном потоке, чтобы интерфейс не зависал
        thread = Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
    
    def process_files(self):
        """Основная логика обработки файлов (выполняется в отдельном потоке)"""
        try:
            # Шаг 1: Загрузка замен из Excel
            self.log("📊 Загрузка замен из Excel...")
            self.progress_var.set(20)
            replacements = self.load_replacements()
            
            if not replacements:
                self.root.after(0, lambda: messagebox.showwarning(
                    "Предупреждение", 
                    "Таблица замен пуста!\nУбедитесь, что в Excel есть данные в столбцах A и B."
                ))
                self.finish_process()
                return
            
            self.log(f"✅ Загружено {len(replacements)} пар замен")
            self.progress_var.set(40)
            
            # Показываем первые 5 замен
            preview = list(replacements.items())[:5]
            for find_text, replace_text in preview:
                self.log(f"   '{find_text}' → '{replace_text}'")
            if len(replacements) > 5:
                self.log(f"   ... и еще {len(replacements) - 5} замен")
            
            # Шаг 2: Замена текста в DXF
            self.log("📐 Обработка DXF файла...")
            self.progress_var.set(60)
            
            result = self.replace_in_dxf(replacements)
            
            if result['success']:
                self.progress_var.set(100)
                self.log(f"✅ ГОТОВО! Выполнено замен: {result['count']}")
                self.log(f"💾 Файл сохранен: {os.path.basename(self.output_file.get())}")
                
                if result['count'] > 0:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Успешно!",
                        f"✅ Замена выполнена успешно!\n"
                        f"Выполнено замен: {result['count']}\n"
                        f"Результат сохранен в:\n{self.output_file.get()}"
                    ))
                else:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Внимание",
                        "✅ Файл обработан, но замен не найдено.\n"
                        "Проверьте, что искомый текст точно совпадает."
                    ))
            else:
                self.log(f"❌ Ошибка: {result['error']}")
                self.progress_var.set(0)
                self.root.after(0, lambda: messagebox.showerror(
                    "Ошибка", 
                    f"Ошибка при обработке DXF:\n{result['error']}"
                ))
            
            self.finish_process()
            
        except Exception as e:
            self.log(f"❌ Критическая ошибка: {str(e)}")
            self.progress_var.set(0)
            self.root.after(0, lambda: messagebox.showerror(
                "Ошибка", 
                f"Произошла ошибка:\n{str(e)}"
            ))
            self.finish_process()
    
    def load_replacements(self):
        """Загружает замены из Excel файла"""
        try:
            wb = openpyxl.load_workbook(self.excel_file.get(), data_only=True)
            sheet = wb.active
            
            replacements = {}
            for row in sheet.iter_rows(min_row=1, values_only=True):
                if len(row) >= 2 and row[0] is not None and row[1] is not None:
                    find_text = str(row[0]).strip()
                    replace_text = str(row[1]).strip()
                    
                    # Пропускаем пустые и заголовки
                    if (find_text and replace_text and 
                        not find_text.upper() in ('FIND', 'ИСКАТЬ', 'A', 'НАЙТИ', 'ЧТО ИСКАТЬ')):
                        replacements[find_text] = replace_text
            
            return replacements
            
        except Exception as e:
            self.log(f"❌ Ошибка чтения Excel: {str(e)}")
            return {}
    
    def replace_in_dxf(self, replacements):
        """Заменяет текст в DXF файле"""
        try:
            # Загружаем DXF
            doc = ezdxf.readfile(self.dxf_file.get())
            msp = doc.modelspace()
            
            # Получаем цвет
            color = int(self.color_var.get())
            
            total_replacements = 0
            
            # Проходим по всем текстовым объектам
            for entity in msp:
                dxftype = entity.dxftype()
                
                if dxftype in ('TEXT', 'MTEXT'):
                    current_text = entity.dxf.text
                    
                    # Проверяем точное совпадение
                    if current_text in replacements:
                        entity.dxf.text = replacements[current_text]
                        entity.dxf.color = color
                        total_replacements += 1
            
            # Сохраняем результат
            doc.saveas(self.output_file.get())
            
            return {'success': True, 'count': total_replacements}
            
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def finish_process(self):
        """Завершает обработку и разблокирует интерфейс"""
        self.root.after(0, lambda: self.run_button.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.status_var.set("Готов"))
        self.root.after(0, lambda: self.log("="*50 + "\n"))

def main():
    root = tk.Tk()
    app = DXFReplacerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()