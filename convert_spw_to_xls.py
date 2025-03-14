import os
import time
import pythoncom
from win32com.client import Dispatch, DispatchEx, gencache
import threading

# Для синхронизации вызова gencache.EnsureModule в многопоточном режиме
gencache_lock = threading.Lock()

def get_kompas_api7():
    with gencache_lock:
        module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        const_module = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0)
    const = const_module.constants
    app = DispatchEx('Kompas.Application.7')
    time.sleep(2)
    print('Создан новый процесс компаса')
    app.Visible = False
    app.HideMessage = const.ksHideMessageNo
    api = module.IKompasAPIObject(
        app._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch)
    )
    return module, api, const, app

def convert_spw_to_xls(spw_file, xls_file, kompas_api):
    try:
        module7, api7, const7, app7 = kompas_api
        print("Начало конвертации SPW в XLS для файла:", spw_file)
        if not os.path.exists(spw_file):
            print(f"Файл не найден: {spw_file}")
            return ""
        print("Открываем документ SPW...")
        doc7 = app7.Documents.Open(PathName=spw_file, Visible=True, ReadOnly=True)
        if doc7 is not None:
            doc7.SaveAs(xls_file)
            print(f'Файл {xls_file} сохранён!')
        else:
            print('Не удалось сохранить документ в XLS')
            return ""
        print("Конвертация SPW в XLS завершена для файла:", spw_file)
        doc7.Close(const7.kdDoNotSaveChanges)
        return xls_file
    except Exception as e:
        print(f"Ошибка при конвертации {spw_file} в XLS: {e}")
        return ""

def convert_to_pdf(input_file, pdf_file, kompas_api):
    try:
        module7, api7, const7, app7 = kompas_api
        print(f"Начало конвертации {input_file} в PDF...")
        if not os.path.exists(input_file):
            print(f"Файл не найден: {input_file}")
            return ""
        doc7 = app7.Documents.Open(PathName=input_file, Visible=True, ReadOnly=True)
        if doc7 is not None:
            doc7.SaveAs(pdf_file)
            print(f'Файл {pdf_file} сохранён!')
        else:
            print('Не удалось сохранить документ в PDF')
            return ""
        print("Конвертация в PDF завершена для файла:", input_file)
        doc7.Close(const7.kdDoNotSaveChanges)
        return pdf_file
    except Exception as e:
        print(f"Ошибка при конвертации {input_file} в PDF: {e}")
        return ""

def search_spw(directory):
    if not os.path.isdir(directory):
        print(f"Директория не найдена или не является каталогом: {directory}")
        return []
    spw_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.spw'):
                spw_files.append(os.path.abspath(os.path.join(root, file)))
    return spw_files

def search_cdw(directory):
    if not os.path.isdir(directory):
        print(f"Директория не найдена или не является каталогом: {directory}")
        return []
    cdw_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.cdw'):
                cdw_files.append(os.path.abspath(os.path.join(root, file)))
    return cdw_files

def do_a_path_for_xls(files, output_dir):
    if not os.path.isabs(output_dir):
        output_dir = os.path.abspath(output_dir)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    xls_files = []
    for file in files:
        base_name = os.path.splitext(os.path.basename(file))[0]
        xls_path = os.path.join(output_dir, base_name + '.xls')
        xls_files.append(os.path.normpath(xls_path))
    return xls_files

def do_a_path_for_pdf(files, output_dir):
    if not os.path.isabs(output_dir):
        output_dir = os.path.abspath(output_dir)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    pdf_files = []
    for file in files:
        base_name = os.path.splitext(os.path.basename(file))[0]
        pdf_path = os.path.join(output_dir, base_name + '.pdf')
        pdf_files.append(os.path.normpath(pdf_path))
    return pdf_files

def convert_spw_to_xls_array(spw_files, xls_files, chunk_size=10, update_callback=None):
    if len(spw_files) != len(xls_files):
        print("Количество SPW файлов и XLS файлов не совпадает.")
        return

    def process_chunk(chunk_spw, chunk_xls):
        import pythoncom
        pythoncom.CoInitialize()
        local_api = get_kompas_api7()
        try:
            for spw, xls in zip(chunk_spw, chunk_xls):
                result = convert_spw_to_xls(spw, xls, local_api)
                if result:
                    print('SPW -> XLS conversion done for:', xls)
                else:
                    print(f"Ошибка конвертации SPW файла: {spw}")
                if update_callback is not None:
                    update_callback(1)
        finally:
            local_api[3].Quit()
            pythoncom.CoUninitialize()

    threads = []
    total_files = len(spw_files)
    for i in range(0, total_files, chunk_size):
        chunk_spw = spw_files[i:i+chunk_size]
        chunk_xls = xls_files[i:i+chunk_size]
        thread = threading.Thread(target=process_chunk, args=(chunk_spw, chunk_xls))
        threads.append(thread)
        thread.start()

    for thread in threads:
        thread.join()

    print("Все потоки для SPW -> XLS завершены.")

def convert_files_to_pdf_array(input_files, pdf_files, chunk_size=10, update_callback=None):
    if len(input_files) != len(pdf_files):
        print("Количество файлов и PDF файлов не совпадает.")
        return

    def process_chunk(chunk_inputs, chunk_pdfs):
        import pythoncom
        pythoncom.CoInitialize()
        local_api = get_kompas_api7()
        try:
            for input_file, pdf in zip(chunk_inputs, chunk_pdfs):
                result = convert_to_pdf(input_file, pdf, local_api)
                if result:
                    print('PDF conversion done for:', pdf)
                else:
                    print(f"Ошибка конвертации файла: {input_file}")
                if update_callback is not None:
                    update_callback(1)
        finally:
            local_api[3].Quit()
            pythoncom.CoUninitialize()

    threads = []
    total_files = len(input_files)
    for i in range(0, total_files, chunk_size):
        chunk_inputs = input_files[i:i+chunk_size]
        chunk_pdfs = pdf_files[i:i+chunk_size]
        thread = threading.Thread(target=process_chunk, args=(chunk_inputs, chunk_pdfs))
        threads.append(thread)
        thread.start()

    for thread in threads:
        thread.join()

    print("Все потоки для PDF конвертации завершены.")

if __name__ == "__main__":
    # Пример использования (замените пути на корректные)
    input_dir = 'path_to_input'
    spw_files = search_spw(input_dir)
    cdw_files = search_cdw(input_dir)
    
    # SPW -> XLS
    xls_output_dir = 'output_xls'
    xls_files = do_a_path_for_xls(spw_files, xls_output_dir)
    convert_spw_to_xls_array(spw_files, xls_files, chunk_size=2, update_callback=lambda n: print("Progress +", n))
    
    # PDF конвертации
    pdf_spec_output_dir = 'output_pdf_spec'
    pdf_drawing_output_dir = 'output_pdf_drawing'
    spw_pdf_files = do_a_path_for_pdf(spw_files, pdf_spec_output_dir)
    cdw_pdf_files = do_a_path_for_pdf(cdw_files, pdf_drawing_output_dir)
    convert_files_to_pdf_array(spw_files, spw_pdf_files, chunk_size=2, update_callback=lambda n: print("Progress +", n))
    convert_files_to_pdf_array(cdw_files, cdw_pdf_files, chunk_size=2, update_callback=lambda n: print("Progress +", n))
