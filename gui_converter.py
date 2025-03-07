from math import ceil
import os
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Глобальные переменные для прогресса
progress_count = 0
progress_lock = threading.Lock()

def convert_files_callback(spw_files, xls_files, num_of_threads, update_callback):
    """
    Запускает многопоточную конвертацию с вызовом update_callback после конвертации каждого файла.
    Использует новую версию обработки файлов, созданную на основе convert_spw_to_xls_array.
    """
    total_files = len(spw_files)

    def process_chunk(chunk_spw, chunk_xls):
        import pythoncom
        from convert_spw_to_xls import convert_spw_to_xls
        pythoncom.CoInitialize()
        local_api = get_local_api()  # каждую группу создаём отдельно в потоке
        try:
            for spw, xls in zip(chunk_spw, chunk_xls):
                convert_spw_to_xls(spw, xls, local_api)
                with progress_lock:
                    global progress_count
                    progress_count += 1
                update_callback()  # вызвать обновление виджетов
        finally:
            local_api[3].Quit()
            pythoncom.CoUninitialize()

    # Функция для создания нового API в потоке (копия get_kompas_api7 без Quit)
    def get_local_api():
        from convert_spw_to_xls import get_kompas_api7
        return get_kompas_api7()

    chunk_size = ceil(total_files/num_of_threads)
    threads = []
    for i in range(0, total_files, chunk_size):
        chunk_spw = spw_files[i:i+chunk_size]
        chunk_xls = xls_files[i:i+chunk_size]
        thread = threading.Thread(target=process_chunk, args=(chunk_spw, chunk_xls))
        threads.append(thread)
        thread.start()
    for thread in threads:
        thread.join()

class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Конвертер SPW -> XLS")
        self.resizable(False, False)
        
        # Путь к папке со спецификациями
        self.input_dir_var = tk.StringVar()
        # Путь к папке для XLS
        self.output_dir_var = tk.StringVar()
        # Размер чанка со значением по умолчанию
        self.num_of_threads_var = tk.StringVar(value="10")
        # Прогресс (0..100)
        self.progress_var = tk.DoubleVar(value=0)
        
        # Общее количество файлов (для прогресса)
        self.total_files = 0
        
        self.create_widgets()
    
    def create_widgets(self):
        padding = {"padx": 10, "pady": 5}
        
        # Ввод папки со спецификациями
        frame_input = ttk.Frame(self)
        frame_input.pack(fill="x", **padding)
        ttk.Label(frame_input, text="Папка со спецификациями:").pack(side="left")
        entry_input = ttk.Entry(frame_input, textvariable=self.input_dir_var, width=40)
        entry_input.pack(side="left", padx=5)
        btn_browse_input = ttk.Button(frame_input, text="Обзор...", command=self.browse_input)
        btn_browse_input.pack(side="left")
        
        # Ввод папки вывода XLS
        frame_output = ttk.Frame(self)
        frame_output.pack(fill="x", **padding)
        ttk.Label(frame_output, text="Папка для XLS:").pack(side="left")
        entry_output = ttk.Entry(frame_output, textvariable=self.output_dir_var, width=40)
        entry_output.pack(side="left", padx=5)
        btn_browse_output = ttk.Button(frame_output, text="Обзор...", command=self.browse_output)
        btn_browse_output.pack(side="left")
        
        # Размер чанка
        frame_num_of_threads = ttk.Frame(self)
        frame_num_of_threads.pack(fill="x", **padding)
        ttk.Label(frame_num_of_threads, text="Количество потоков:").pack(side="left")
        entry_num_of_threads = ttk.Entry(frame_num_of_threads, textvariable=self.num_of_threads_var, width=10)
        entry_num_of_threads.pack(side="left", padx=5)
        
        # Кнопка запуска конвертации
        self.btn_start = ttk.Button(self, text="Запустить конвертацию", command=self.start_conversion)
        self.btn_start.pack(pady=10)
        
        # Полоса загрузки
        self.progressbar = ttk.Progressbar(self, orient="horizontal", length=400, mode="determinate", variable=self.progress_var)
        self.progressbar.pack(pady=10)

        self.update_idletasks()
        self.geometry(f"{self.winfo_reqwidth()}x{self.winfo_reqheight()}")
    
    def browse_input(self):
        directory = filedialog.askdirectory(title="Выберите папку со спецификациями")
        if directory:
            self.input_dir_var.set(directory)
    
    def browse_output(self):
        directory = filedialog.askdirectory(title="Выберите папку для XLS")
        if directory:
            self.output_dir_var.set(directory)
    
    def start_conversion(self):
        input_dir = self.input_dir_var.get()
        output_dir = self.output_dir_var.get()
        if not os.path.isdir(input_dir):
            messagebox.showerror("Ошибка", "Укажите корректную папку со спецификациями")
            return
        if not os.path.isdir(output_dir):
            # Если папки нет, создаём её
            os.makedirs(output_dir)
            self.output_dir_var.set(os.path.abspath(output_dir))
        
        try:
            chunk_size = int(self.num_of_threads_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Количество потоков должно быть числом")
            return
        
        # Поиск всех spw файлов в заданной папке
        from convert_spw_to_xls import search_spw  # импорт функции из модуля конвертера
        spw_files = search_spw(input_dir)
        spw_files = spw_files
        if not spw_files:
            messagebox.showinfo("Информация", "Не найдены файлы .spw в указанной папке")
            return
        
        # Создание списка для XLS путей
        from convert_spw_to_xls import do_a_path_for_xls
        xls_files = do_a_path_for_xls(spw_files, output_dir)
        
        # Обновляем общее количество файлов и полосу прогресса
        self.total_files = len(spw_files)
        self.progressbar.configure(maximum=self.total_files)
        global progress_count
        progress_count = 0
        self.progress_var.set(0)
        
        # Отключаем кнопку запуска
        self.btn_start.configure(state="disabled")
        
        # Замер времени
        self.start_time = time.perf_counter()
        
        # Запуск конвертации в отдельном потоке
        threading.Thread(target=self.run_conversion, args=(spw_files, xls_files, chunk_size), daemon=True).start()
        # Запускаем опрос прогресса
        self.after(100, self.update_progress)
    
    def run_conversion(self, spw_files, xls_files, chunk_size):
        # Функция-обёртка для запуска конвертации с callback-ом для обновления прогресса
        convert_files_callback(spw_files, xls_files, chunk_size, lambda: None)
    
    def update_progress(self):
        global progress_count
        self.progress_var.set(progress_count)
        if progress_count < self.total_files:
            self.after(200, self.update_progress)
        else:
            # Конвертация завершена: отключаем кнопку запуска
            self.btn_start.configure(state="normal")
            elapsed = time.perf_counter() - self.start_time
            
            # Создаем новое окно с результатами
            result_win = tk.Toplevel(self)
            result_win.title("Конвертация завершена")
            result_win.resizable(False, False)
            
            lbl_time = ttk.Label(result_win, text=f"Время конвертации: {time.strftime('%H:%M:%S', time.gmtime(elapsed))}")
            lbl_time.pack(pady=20)
            
            frame_buttons = ttk.Frame(result_win)
            frame_buttons.pack(pady=10)
            
            btn_open_folder = ttk.Button(frame_buttons, text="Открыть папку", 
                                         command=lambda: os.startfile(self.output_dir_var.get()))
            btn_open_folder.pack(side="left", padx=10)
            
            btn_ok = ttk.Button(frame_buttons, text="ОК", 
                                command=result_win.destroy)
            btn_ok.pack(side="left", padx=10)
            
            # Обновляем окно и устанавливаем минимальный размер, подходящий под содержимое
            result_win.update_idletasks()
            result_win.geometry(f"{result_win.winfo_reqwidth()}x{result_win.winfo_reqheight()}")

if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()