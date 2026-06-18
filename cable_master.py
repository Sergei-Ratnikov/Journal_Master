# cable_master.py
"""
Графический интерфейс для программы парсинга кабельных журналов.
Вкладка 1: Парсер баз данных (существующий функционал)
Вкладка 2: Поиск ответных частей (cable_matcher.py)
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import subprocess
import os
from excel_utils import build_cable_database
from cable_matcher import process_journal
from pathlib import Path

class CableParserGUI:
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Кабельный журнал - Парсер баз данных")
        self.root.geometry("650x600")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")
        
        # Переменные для вкладки 1 (Парсер)
        self.journals_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.open_folder = tk.BooleanVar(value=True)  # ← ПЕРЕМЕННАЯ
        
        # Переменные для вкладки 2 (Поиск ответных частей)
        self.matcher_db_path = tk.StringVar()
        self.matcher_journal_kks = tk.StringVar()
        # self.matcher_output_path = tk.StringVar()
        self.matcher_open_folder = tk.BooleanVar(value=True)
        
        self.is_running = False
        self.last_saved_file = None
        
        self.setup_ui()

    def show_about(self):
        """Показывает окно 'О программе' с информацией о парсере"""
        about_text = """Парсер кабельных журналов

        Назначение
        Программа собирает данные из кабельных журналов (Word-документы .docx/.doc) в единую Excel-базу.
        Имя файла: ККС_журнала_источник_дата.docx (например, AKU.0120.00UKS.0.EM.MB0002-EMB0001_ВКРД_15.09.2024.docx). 
        Источник - ВКРД/ВКИИ/СРД/СИИ.

        Что делает программа
        • Сканирует папку с журналами, находит все .docx и .doc файлы.
        • Конвертирует .doc в .docx (файлы перемещаются в папку doc files внутри папки с журналами).
        • Конвертирует автоматическую нумерацию в обычный текст.
        • Извлекает данные из таблиц: номер кабеля, ККС, марка, сечение, длина, трасса, координаты X/Y/Z, помещения и оборудование "Откуда" и "Куда".
        • Из имени файла извлекает источник и дату, записывает их в соответствующие колонки.
        • Переносит примечания из старой версии базы, если кабель с таким же ККС присутствовал в предыдущей версии.

        Создание базы
        • Если база создаётся впервые — формируется новый файл Cable base ver.1.xlsx.
        • Если база уже существует — создаётся новая версия (ver.2, ver.3 и т.д.). В неё переносятся:
        - Все старые записи, кроме тех, что были заменены новыми журналами.
        - Лист "Объединенные кабели" с историей объединений.
        - Примечания к кабелям (если они были в старой версии).

        Дополнительные листы в базе
        • "Объединенные кабели" — список ККС кабелей, которые были объединены.
        • "Лог" — отчёт о работе: какие журналы добавлены, заменены, конвертированы или пропущены.

        Как использовать
        1. Выберите папку с кабельными журналами (исходные файлы).
        2. Выберите папку для сохранения базы данных.
        3. Нажмите "Запустить обработку".
        4. Дождитесь завершения — программа покажет прогресс и сообщит о результате.

        Результат — файл Cable base ver.N.xlsx в выбранной папке.

        Важно: Журналы должны быть в формате .docx или .doc. Файлы .doc будут автоматически конвертированы в .docx. Старые версии базы не перезаписываются — каждая новая версия сохраняется отдельно."""

        messagebox.showinfo("О программе", about_text)

    def setup_ui(self):
        """Создаёт интерфейс с вкладками"""
        
        # Заголовок
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=90)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)
        
        tk.Label(title_frame, text="JOURNAL MASTER", 
                font=("Arial", 18, "bold"), fg="white", bg="#2c3e50").pack(expand=True)
        tk.Label(title_frame, text="Парсер и создание базы данных",
                font=("Arial", 10), fg="#bdc3c7", bg="#2c3e50").pack()
        tk.Label(title_frame, text="АО Сосновоборэлектромонтаж, Проектное бюро, ver. 1.0, 2026",
                font=("Arial", 10), fg="#bdc3c7", bg="#2c3e50").pack()
        
        # Основная область с вкладками
        main_frame = tk.Frame(self.root, padx=20, pady=20, bg="#f0f0f0")
        main_frame.pack(fill="both", expand=True)
        
        # Создаём виджет с вкладками
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # ========== ВКЛАДКА 1: ПАРСЕР БАЗ ==========
        self.tab_parser = tk.Frame(self.notebook, bg="#f0f0f0")
        self.notebook.add(self.tab_parser, text="📊 Парсер баз данных")
        self.setup_parser_tab()
        
        # ========== ВКЛАДКА 2: ПОИСК ОТВЕТНЫХ ЧАСТЕЙ ==========
        self.tab_matcher = tk.Frame(self.notebook, bg="#f0f0f0")
        self.notebook.add(self.tab_matcher, text="🔍 Поиск ответных частей")
        self.setup_matcher_tab()
        
        # Информация внизу
        info_frame = tk.Frame(self.root, bg="#ecf0f1", height=50)
        info_frame.pack(fill="x", side="bottom")
        info_frame.pack_propagate(False)
        tk.Label(info_frame, text="Поддерживаются форматы: .docx, .doc\nБаза сохраняется в .xlsx",
                font=("Arial", 8), fg="#7f8c8d", bg="#ecf0f1").pack(expand=True)
    
    # ========== ВКЛАДКА 1: ПАРСЕР БАЗ ==========
    def setup_parser_tab(self):
        """Настройка вкладки 'Парсер баз данных'"""
        main_frame = tk.Frame(self.tab_parser, padx=20, pady=20, bg="#f0f0f0")
        main_frame.pack(fill="both", expand=True)
        
        # Папка с журналами
        journals_frame = tk.LabelFrame(main_frame, text="📄 Исходные журналы", 
                                       padx=10, pady=10, font=("Arial", 10, "bold"), bg="#f0f0f0")
        journals_frame.pack(fill="x", pady=(0, 15))
        
        tk.Entry(journals_frame, textvariable=self.journals_dir, 
                font=("Arial", 10), state="readonly", bg="white").pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Button(journals_frame, text="Выбрать папку", command=self.select_journals_dir,
                 bg="#3498db", fg="white", font=("Arial", 9), padx=10, cursor="hand2").pack(side="right")
        
        # Папка для сохранения
        output_frame = tk.LabelFrame(main_frame, text="💾 Сохранение базы", 
                                     padx=10, pady=10, font=("Arial", 10, "bold"), bg="#f0f0f0")
        output_frame.pack(fill="x", pady=(0, 15))
        
        tk.Entry(output_frame, textvariable=self.output_dir, 
                font=("Arial", 10), state="readonly", bg="white").pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Button(output_frame, text="Выбрать папку", command=self.select_output_dir,
                 bg="#3498db", fg="white", font=("Arial", 9), padx=10, cursor="hand2").pack(side="right")
        
        # Чекбокс
        checkbox_frame = tk.Frame(main_frame, bg="#f0f0f0")
        checkbox_frame.pack(fill="x", pady=(0, 10))
        
        self.open_folder_checkbox = tk.Checkbutton(
            checkbox_frame, 
            text="📂 Открыть папку с результатом после обработки",
            variable=self.open_folder,  # ← используем переменную
            font=("Arial", 9),
            bg="#f0f0f0",
            fg="#2c3e50",
            anchor="w",
            cursor="hand2"
        )
        self.open_folder_checkbox.pack(side="left")
        
        # Прогресс
        progress_frame = tk.LabelFrame(main_frame, text="📊 Прогресс обработки", 
                                       padx=10, pady=10, font=("Arial", 10, "bold"), bg="#f0f0f0")
        progress_frame.pack(fill="x", pady=(0, 15))
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress_bar.pack(fill="x", pady=(0, 5))
        
        self.progress_percent = tk.Label(progress_frame, text="0%", font=("Arial", 9), bg="#f0f0f0", fg="#2c3e50")
        self.progress_percent.pack()
        
        self.progress_detail = tk.Label(progress_frame, text="Ожидание запуска...", 
                                        font=("Arial", 8), bg="#f0f0f0", fg="#7f8c8d")
        self.progress_detail.pack(pady=(5, 0))
        
        # Кнопки управления
        buttons_frame = tk.Frame(main_frame, bg="#f0f0f0")
        buttons_frame.pack(fill="x", pady=(0, 15))
        
        self.run_btn = tk.Button(buttons_frame, text="🚀 ЗАПУСТИТЬ ОБРАБОТКУ",
                                 command=self.run_parser, bg="#27ae60", fg="white",
                                 font=("Arial", 12, "bold"), height=2, cursor="hand2")
        self.run_btn.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        self.about_btn = tk.Button(buttons_frame, text="❓ О программе",
                                   command=self.show_about, bg="#3498db", fg="white",
                                   font=("Arial", 10, "bold"), height=2, cursor="hand2")
        self.about_btn.pack(side="right")
        
        # Статус
        self.status_label = tk.Label(main_frame, text="Готов к работе. Выберите папки и нажмите 'Запустить'.", 
                                     font=("Arial", 9), fg="#7f8c8d", bg="#f0f0f0", wraplength=550, justify="center")
        self.status_label.pack()
    
    # ========== ВКЛАДКА 2: ПОИСК ОТВЕТНЫХ ЧАСТЕЙ ==========
    def setup_matcher_tab(self):
        """Настройка вкладки 'Поиск ответных частей'"""
        main_frame = tk.Frame(self.tab_matcher, padx=20, pady=20, bg="#f0f0f0")
        main_frame.pack(fill="both", expand=True)
        
        # Пояснение
        info_label = tk.Label(main_frame, 
            text="Поиск кабелей, у которых есть ответная часть в другом журнале",
            font=("Arial", 10), fg="#2c3e50", bg="#f0f0f0")
        info_label.pack(pady=(0, 15))
        
        # Файл базы данных
        db_frame = tk.LabelFrame(main_frame, text="📁 Файл базы данных", 
                                 padx=10, pady=10, font=("Arial", 10, "bold"), bg="#f0f0f0")
        db_frame.pack(fill="x", pady=(0, 15))
        
        tk.Entry(db_frame, textvariable=self.matcher_db_path, 
                font=("Arial", 10), state="readonly", bg="white").pack(side="left", fill="x", expand=True, padx=(0, 10))
        tk.Button(db_frame, text="Выбрать файл", command=self.select_matcher_db,
                 bg="#3498db", fg="white", font=("Arial", 9), padx=10, cursor="hand2").pack(side="right")
        
        # ККС журнала
        journal_frame = tk.LabelFrame(main_frame, text="ККС журнала для поиска", 
                                      padx=10, pady=10, font=("Arial", 10, "bold"), bg="#f0f0f0")
        journal_frame.pack(fill="x", pady=(0, 15))
        
        tk.Label(journal_frame, text="ККС журнала:",
                font=("Arial", 9), bg="#f0f0f0", fg="#2c3e50").pack(anchor="w", pady=(0, 5))
        
        self.journal_entry = tk.Entry(journal_frame, textvariable=self.matcher_journal_kks,
                                      font=("Arial", 10), bg="white")
        self.journal_entry.pack(fill="x", pady=(0, 5))
               
        # Чекбокс
        checkbox_frame = tk.Frame(main_frame, bg="#f0f0f0")
        checkbox_frame.pack(fill="x", pady=(0, 10))
        
        self.matcher_open_folder_checkbox = tk.Checkbutton(
            checkbox_frame, 
            text="📂 Открыть папку с результатом после обработки",
            variable=self.matcher_open_folder,
            font=("Arial", 9),
            bg="#f0f0f0",
            fg="#2c3e50",
            anchor="w",
            cursor="hand2"
        )
        self.matcher_open_folder_checkbox.pack(side="left")
        
        # Кнопка запуска
        self.matcher_run_btn = tk.Button(main_frame, text="🔍 НАЙТИ ОТВЕТНЫЕ ЧАСТИ",
                                         command=self.run_matcher, bg="#8e44ad", fg="white",
                                         font=("Arial", 12, "bold"), height=2, cursor="hand2")
        self.matcher_run_btn.pack(fill="x", pady=(0, 15))
        
        # Статус
        self.matcher_status_label = tk.Label(main_frame, text="Готов к работе. Выберите базу, введите ККС журнала и нажмите 'Найти'.", 
                                             font=("Arial", 9), fg="#7f8c8d", bg="#f0f0f0", wraplength=550, justify="center")
        self.matcher_status_label.pack()
    
    # ========== ОБЩИЕ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========
    
    def update_status(self, message, is_error=False):
        """Обновляет статус на активной вкладке"""
        current_tab = self.notebook.select()
        if current_tab == str(self.tab_parser):
            if is_error:
                self.status_label.config(text=f"❌ {message}", fg="#e74c3c")
            else:
                self.status_label.config(text=message, fg="#7f8c8d")
        else:
            if is_error:
                self.matcher_status_label.config(text=f"❌ {message}", fg="#e74c3c")
            else:
                self.matcher_status_label.config(text=message, fg="#7f8c8d")
        self.root.update_idletasks()
    
    def update_progress(self, current, total, message, is_error=False):
        """Обновляет прогресс на вкладке парсера"""
        if total > 0:
            percent = int((current / total) * 100)
            self.progress_bar["value"] = percent
            self.progress_percent.config(text=f"{percent}% ({current}/{total})")
        
        if is_error:
            self.progress_detail.config(text=f"⚠️ {message}", fg="#e74c3c")
        else:
            self.progress_detail.config(text=message, fg="#2c3e50")
        
        self.root.update_idletasks()
    
    def open_folder_path(self, folder_path):  # ← МЕТОД (другое имя!)
        """Открывает папку в проводнике"""
        if not folder_path:
            return
        folder_path = os.path.normpath(folder_path)
        if os.path.exists(folder_path):
            subprocess.Popen(f'explorer "{folder_path}"')
            self.update_status(f"Открыта папка: {folder_path}")
        else:
            self.update_status(f"Папка не существует: {folder_path}", is_error=True)
    
    # ========== ФУНКЦИИ ДЛЯ ВКЛАДКИ 1 (ПАРСЕР) ==========
    
    def select_journals_dir(self):
        directory = filedialog.askdirectory(title="Выберите папку с кабельными журналами")
        if directory:
            self.journals_dir.set(directory)
            self.update_status(f"Папка с журналами: {directory}")
    
    def select_output_dir(self):
        directory = filedialog.askdirectory(title="Выберите папку для сохранения базы данных")
        if directory:
            self.output_dir.set(directory)
            self.update_status(f"Папка для сохранения: {directory}")
    
    def run_parser(self):
        if self.is_running:
            messagebox.showwarning("Внимание", "Обработка уже выполняется!")
            return
        
        if not self.journals_dir.get():
            messagebox.showerror("Ошибка", "Выберите папку с журналами!")
            return
        
        if not self.output_dir.get():
            messagebox.showerror("Ошибка", "Выберите папку для сохранения базы!")
            return
        
        self.is_running = True
        self.run_btn.config(state="disabled", bg="#95a5a6")
        
        self.progress_bar["value"] = 0
        self.progress_percent.config(text="0%")
        self.progress_detail.config(text="Подготовка к обработке...", fg="#2c3e50")
        self.update_status("⏳ Обработка журналов...")
        
        thread = threading.Thread(target=self._process_parser, daemon=True)
        thread.start()
    
    def _process_parser(self):
        start_time = time.time()
        try:
            build_cable_database(
                self.journals_dir.get(), 
                self.output_dir.get(),
                progress_callback=self.update_progress
            )
            self.root.after(0, self._on_parser_success, start_time)
        except Exception as e:
            self.root.after(0, self._on_parser_error, str(e))
    
    def _on_parser_success(self, start_time):
        duration = time.time() - start_time
        self.is_running = False
        self.run_btn.config(state="normal", bg="#27ae60")
        self.update_status(f"✅ Обработка завершена! Время: {duration:.2f} сек")
        self.progress_detail.config(text="Готово! База данных создана.", fg="#27ae60")
        
        if self.open_folder.get():  # ← проверяем переменную
            self.root.after(500, lambda: self.open_folder_path(self.output_dir.get()))  # ← вызываем метод
        
        messagebox.showinfo("Готово!", 
            f"✅ Обработка завершена!\n\n"
            f"📁 Журналы: {self.journals_dir.get()}\n"
            f"💾 База сохранена в: {self.output_dir.get()}\n"
            f"⏱️ Время: {duration:.2f} сек")
    
    def _on_parser_error(self, error_msg):
        self.is_running = False
        self.run_btn.config(state="normal", bg="#27ae60")
        self.update_status(f"Ошибка: {error_msg}", is_error=True)
        self.progress_detail.config(text=f"Ошибка: {error_msg[:80]}", fg="#e74c3c")
        messagebox.showerror("Ошибка", f"При обработке произошла ошибка:\n\n{error_msg}")
    
    # ========== ФУНКЦИИ ДЛЯ ВКЛАДКИ 2 (ПОИСК ОТВЕТНЫХ ЧАСТЕЙ) ==========
    
    def select_matcher_db(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл базы данных",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.matcher_db_path.set(file_path)
            self.update_status(f"База данных: {file_path}")
    
    def select_matcher_output(self):
        file_path = filedialog.asksaveasfilename(
            title="Сохранить результат как...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.matcher_output_path.set(file_path)
            self.update_status(f"Результат будет сохранён в: {file_path}")
    
    def run_matcher(self):
        if self.is_running:
            messagebox.showwarning("Внимание", "Обработка уже выполняется!")
            return
        
        if not self.matcher_db_path.get():
            messagebox.showerror("Ошибка", "Выберите файл базы данных!")
            return
        
        if not self.matcher_journal_kks.get():
            messagebox.showerror("Ошибка", "Введите ККС журнала!")
            return
        
        self.is_running = True
        self.matcher_run_btn.config(state="disabled", bg="#95a5a6")
        self.update_status("⏳ Поиск ответных частей...")
        
        thread = threading.Thread(target=self._process_matcher, daemon=True)
        thread.start()
    
    def _process_matcher(self):
        start_time = time.time()
        try:
            # Формируем имя выходного файла
            db_path = Path(self.matcher_db_path.get())
            journal_kks = self.matcher_journal_kks.get().strip()
            # Очищаем ККС от недопустимых символов для имени файла
            safe_journal = journal_kks.replace('\\', '_').replace('/', '_').replace(':', '_')
            output_filename = f"Отчет по журналу {safe_journal}.xlsx"
            output_path = db_path.parent / output_filename
            
            process_journal(
                self.matcher_db_path.get(),
                journal_kks,
                str(output_path)
            )
            self.root.after(0, self._on_matcher_success, start_time, str(output_path))
        except Exception as e:
            self.root.after(0, self._on_matcher_error, str(e))
    
    def _on_matcher_success(self, start_time, output_path):
        duration = time.time() - start_time
        self.is_running = False
        self.matcher_run_btn.config(state="normal", bg="#8e44ad")
        self.update_status(f"✅ Обработка завершена! Время: {duration:.2f} сек")
        
        if self.matcher_open_folder.get():
            output_dir = os.path.dirname(output_path)
            self.root.after(500, lambda: self.open_folder_path(output_dir))
        
        messagebox.showinfo("Готово!", 
            f"✅ Поиск завершён!\n\n"
            f"📁 База: {self.matcher_db_path.get()}\n"
            f"📌 Журнал: {self.matcher_journal_kks.get()}\n"
            f"💾 Результат: {output_path}\n"
            f"⏱️ Время: {duration:.2f} сек")
    
    def _on_matcher_error(self, error_msg):
        self.is_running = False
        self.matcher_run_btn.config(state="normal", bg="#8e44ad")
        self.update_status(f"Ошибка: {error_msg}", is_error=True)
        messagebox.showerror("Ошибка", f"При обработке произошла ошибка:\n\n{error_msg}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = CableParserGUI()
    app.run()