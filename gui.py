# gui.py
"""
Графический интерфейс для программы парсинга кабельных журналов.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import subprocess
import os
from excel_utils import build_cable_database


class CableParserGUI:
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Кабельный журнал - Парсер баз данных")
        self.root.geometry("650x550")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")
        
        self.journals_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.open_folder = tk.BooleanVar(value=True)  # Чекбокс: открывать папку по умолчанию
        self.is_running = False
        self.last_saved_file = None  # Для хранения пути к сохранённому файлу
        
        self.setup_ui()
        
    def setup_ui(self):
        """Создаёт интерфейс"""
        
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
        
        # Основная область
        main_frame = tk.Frame(self.root, padx=20, pady=20, bg="#f0f0f0")
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
        output_frame.pack(fill="x", pady=(0, 20))
        
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
            variable=self.open_folder,
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
        
        # Кнопка запуска
        self.run_btn = tk.Button(main_frame, text="🚀 ЗАПУСТИТЬ ОБРАБОТКУ",
                                 command=self.run_parser, bg="#27ae60", fg="white",
                                 font=("Arial", 12, "bold"), height=2, cursor="hand2")
        self.run_btn.pack(fill="x", pady=(0, 15))
        
        # Статус
        self.status_label = tk.Label(main_frame, text="Готов к работе. Выберите папки и нажмите 'Запустить'.", 
                                     font=("Arial", 9), fg="#7f8c8d", bg="#f0f0f0", wraplength=550, justify="center")
        self.status_label.pack()
        
        # Информация внизу
        info_frame = tk.Frame(self.root, bg="#ecf0f1", height=50)
        info_frame.pack(fill="x", side="bottom")
        info_frame.pack_propagate(False)
        tk.Label(info_frame, text="Поддерживаются форматы: .docx, .doc\nБаза сохраняется в .xlsx",
                font=("Arial", 8), fg="#7f8c8d", bg="#ecf0f1").pack(expand=True)
    
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
    
    def update_status(self, message, is_error=False):
        if is_error:
            self.status_label.config(text=f"❌ {message}", fg="#e74c3c")
        else:
            self.status_label.config(text=message, fg="#7f8c8d")
        self.root.update_idletasks()
    
    def update_progress(self, current, total, message, is_error=False):
        """Обновляет прогресс-бар (вызывается из excel_utils)"""
        if total > 0:
            percent = int((current / total) * 100)
            self.progress_bar["value"] = percent
            self.progress_percent.config(text=f"{percent}% ({current}/{total})")
        
        if is_error:
            self.progress_detail.config(text=f"⚠️ {message}", fg="#e74c3c")
        else:
            self.progress_detail.config(text=message, fg="#2c3e50")
        
        self.root.update_idletasks()
    
    def open_output_folder(self):
        """Открывает папку с сохранённой базой данных"""
        folder = self.output_dir.get().strip()
        
        if not folder:
            self.update_status("Не выбрана папка для сохранения!", is_error=True)
            return
        
        # Нормализуем путь (преобразуем в формат Windows)
        folder = os.path.normpath(folder)
        
        if os.path.exists(folder):
            # Открываем папку в проводнике Windows
            subprocess.Popen(f'explorer "{folder}"')
            self.update_status(f"Открыта папка: {folder}")
        else:
            self.update_status(f"Папка не существует: {folder}", is_error=True)
    
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
        
        # Сброс прогресса
        self.progress_bar["value"] = 0
        self.progress_percent.config(text="0%")
        self.progress_detail.config(text="Подготовка к обработке...", fg="#2c3e50")
        self.update_status("⏳ Обработка журналов...")
        
        thread = threading.Thread(target=self._process, daemon=True)
        thread.start()
    
    def _process(self):
        start_time = time.time()
        try:
            build_cable_database(
                self.journals_dir.get(), 
                self.output_dir.get(),
                progress_callback=self.update_progress
            )
            self.root.after(0, self._on_success, start_time)
        except Exception as e:
            self.root.after(0, self._on_error, str(e))
    
    def _on_success(self, start_time):
        duration = time.time() - start_time
        self.is_running = False
        self.run_btn.config(state="normal", bg="#27ae60")
        self.update_status(f"✅ Обработка завершена! Время: {duration:.2f} сек")
        self.progress_detail.config(text="Готово! База данных создана.", fg="#27ae60")
        
        # Открываем папку, если чекбокс отмечен
        if self.open_folder.get():
            # Небольшая задержка для завершения всех операций записи
            self.root.after(500, self.open_output_folder)
        
        messagebox.showinfo("Готово!", 
            f"✅ Обработка завершена!\n\n"
            f"📁 Журналы: {self.journals_dir.get()}\n"
            f"💾 База сохранена в: {self.output_dir.get()}\n"
            f"⏱️ Время: {duration:.2f} сек")
    
    def _on_error(self, error_msg):
        self.is_running = False
        self.run_btn.config(state="normal", bg="#27ae60")
        self.update_status(f"Ошибка: {error_msg}", is_error=True)
        self.progress_detail.config(text=f"Ошибка: {error_msg[:80]}", fg="#e74c3c")
        messagebox.showerror("Ошибка", f"При обработке произошла ошибка:\n\n{error_msg}")
    
    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = CableParserGUI()
    app.run()