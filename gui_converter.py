import os
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from math import ceil

# Глобальные переменные для прогресса
progress_count = 0
progress_lock = threading.Lock()

class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Конвертер CDW/SPW -> PDF/XLS")
        self.resizable(False, False)
        
        self.input_dir_var = tk.StringVar()
        self.num_of_threads_var = tk.StringVar(value="10")
        self.progress_var = tk.DoubleVar(value=0)
        self.total_tasks = 0
        
        # Чекбоксы для выбора экспорта
        self.export_cdw_pdf_var = tk.BooleanVar(value=True)
        self.export_spw_pdf_var = tk.BooleanVar(value=True)
        self.export_spw_xls_var = tk.BooleanVar(value=True)
        
        self.create_widgets()
    
    def create_widgets(self):
        padding = {"padx": 10, "pady": 5}
        
        # Папка с исходными файлами
        frame_input = ttk.Frame(self)
        frame_input.pack(fill="x", **padding)
        ttk.Label(frame_input, text="Папка с исходными файлами:").pack(side="left")
        entry_input = ttk.Entry(frame_input, textvariable=self.input_dir_var, width=40)
        entry_input.pack(side="left", padx=5)
        btn_browse_input = ttk.Button(frame_input, text="Обзор...", command=self.browse_input)
        btn_browse_input.pack(side="left")
        
        # Чекбоксы выбора экспорта
        frame_options = ttk.Frame(self)
        frame_options.pack(fill="x", **padding)
        ttk.Checkbutton(frame_options, text="Экспорт чертежей (CDW -> PDF)", variable=self.export_cdw_pdf_var)\
            .grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(frame_options, text="Экспорт спецификаций в PDF (SPW -> PDF)", variable=self.export_spw_pdf_var)\
            .grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(frame_options, text="Экспорт спецификаций в XLS (SPW -> XLS)", variable=self.export_spw_xls_var)\
            .grid(row=2, column=0, sticky="w", padx=5, pady=2)
        
        # Количество потоков
        frame_threads = ttk.Frame(self)
        frame_threads.pack(fill="x", **padding)
        ttk.Label(frame_threads, text="Количество потоков:").pack(side="left")
        entry_threads = ttk.Entry(frame_threads, textvariable=self.num_of_threads_var, width=10)
        entry_threads.pack(side="left", padx=5)
        
        # Кнопка старта
        self.btn_start = ttk.Button(self, text="Запустить конвертацию", command=self.start_conversion)
        self.btn_start.pack(pady=10)
        
        # Полоса прогресса
        self.progressbar = ttk.Progressbar(self, orient="horizontal", length=400, mode="determinate", variable=self.progress_var)
        self.progressbar.pack(pady=10)
        
        self.update_idletasks()
        self.geometry(f"{self.winfo_reqwidth()}x{self.winfo_reqheight()}")
    
    def browse_input(self):
        directory = filedialog.askdirectory(title="Выберите папку с исходными файлами")
        if directory:
            self.input_dir_var.set(directory)
    
    def start_conversion(self):
        input_dir = self.input_dir_var.get()
        if not os.path.isdir(input_dir):
            messagebox.showerror("Ошибка", "Укажите корректную папку с исходными файлами")
            return
        
        try:
            num_threads = int(self.num_of_threads_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Количество потоков должно быть числом")
            return
        
        # Вычисляем родительскую папку и создаём выходные каталоги
        parent_dir = os.path.abspath(os.path.join(input_dir, os.pardir))
        output_cdw_pdf = os.path.join(parent_dir, "CDW-PDF")
        output_spw_pdf = os.path.join(parent_dir, "SPW-PDF")
        output_spw_xls = os.path.join(parent_dir, "SPW-XLS")
        for folder in [output_cdw_pdf, output_spw_pdf, output_spw_xls]:
            if not os.path.isdir(folder):
                os.makedirs(folder, exist_ok=True)
        
        from convert_spw_to_xls import search_spw, search_cdw, do_a_path_for_pdf, do_a_path_for_xls, \
            convert_spw_to_xls_array, convert_files_to_pdf_array
        
        tasks = []
        total = 0
        spw_files = search_spw(input_dir) if (self.export_spw_pdf_var.get() or self.export_spw_xls_var.get()) else []
        cdw_files = search_cdw(input_dir) if self.export_cdw_pdf_var.get() else []
        
        if self.export_spw_xls_var.get():
            spw_xls_files = do_a_path_for_xls(spw_files, output_spw_xls)
            total += len(spw_files)
            def run_spw_xls():
                convert_spw_to_xls_array(spw_files, spw_xls_files, chunk_size=num_threads)
                self.update_progress(len(spw_files))
            tasks.append(run_spw_xls)
        
        if self.export_spw_pdf_var.get():
            spw_pdf_files = do_a_path_for_pdf(spw_files, output_spw_pdf)
            total += len(spw_files)
            def run_spw_pdf():
                convert_files_to_pdf_array(spw_files, spw_pdf_files, chunk_size=num_threads)
                self.update_progress(len(spw_files))
            tasks.append(run_spw_pdf)
        
        if self.export_cdw_pdf_var.get():
            cdw_pdf_files = do_a_path_for_pdf(cdw_files, output_cdw_pdf)
            total += len(cdw_files)
            def run_cdw_pdf():
                convert_files_to_pdf_array(cdw_files, cdw_pdf_files, chunk_size=num_threads)
                self.update_progress(len(cdw_files))
            tasks.append(run_cdw_pdf)
        
        if total == 0:
            messagebox.showinfo("Информация", "Нет файлов для конвертации по выбранным опциям")
            return
        
        self.total_tasks = total
        self.progressbar.configure(maximum=self.total_tasks)
        global progress_count
        progress_count = 0
        self.progress_var.set(0)
        self.btn_start.configure(state="disabled")
        self.start_time = time.perf_counter()
        
        # Запускаем задачи в отдельных потоках
        self.threads = []
        for task in tasks:
            t = threading.Thread(target=task, daemon=True)
            self.threads.append(t)
            t.start()
        
        # Начинаем периодически проверять состояние потоков
        self.check_threads(parent_dir)
    
    def check_threads(self, folder_to_open):
        if any(t.is_alive() for t in self.threads):
            self.after(200, lambda: self.check_threads(folder_to_open))
        else:
            elapsed = time.perf_counter() - self.start_time
            self.show_result_window(elapsed, folder_to_open)
    
    def update_progress(self, n):
        global progress_count
        with progress_lock:
            progress_count += n
            self.progress_var.set(progress_count)
    
    def show_result_window(self, elapsed, folder_to_open):
        result_win = tk.Toplevel(self)
        result_win.title("Конвертация завершена")
        result_win.resizable(False, False)
        time_str = time.strftime('%H:%M:%S', time.gmtime(elapsed))
        ttk.Label(result_win, text=f"Время конвертации: {time_str}").pack(pady=10, padx=10)
        frame_buttons = ttk.Frame(result_win)
        frame_buttons.pack(pady=10)
        btn_open = ttk.Button(frame_buttons, text="Открыть папку", command=lambda: os.startfile(folder_to_open))
        btn_open.pack(side="left", padx=5)
        btn_ok = ttk.Button(frame_buttons, text="ОК", command=result_win.destroy)
        btn_ok.pack(side="left", padx=5)
        self.btn_start.configure(state="normal")

if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()
