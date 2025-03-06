import os
import time
import pythoncom
from win32com.client import Dispatch, DispatchEx, gencache, GetActiveObject


def get_kompas_api7():  # Получаем АПИ компас 7 версии

    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    const_module = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0)
    const = const_module.constants
    try:
        app = GetActiveObject('Kompas.Application.7')
        print('Подключились к запущенному компасу')
    except Exception:
        app = DispatchEx('Kompas.Application.7')
        time.sleep(5)
        print('Создан новый процесс компаса')

    app.Visible = True
    app.HideMessage = const.ksHideMessageNo
    api = module.IKompasAPIObject(
        app._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch)
    )
    return module, api, const, app


def convert_spw_to_xls(spw_file, xls_file=None, kompas_api=None):

    try:

        if kompas_api is not None:
            module7, api7, const7, app7 = kompas_api
        else:
             module7, api7, const7, app7 = get_kompas_api7()

        print("Начало конвертации...")
        
        
        if not os.path.exists(spw_file):
            print(f"Файл не найден: {spw_file}")
            return ""
        
        print("Открываем документ...")

        # Открываем документ

        doc7 = app7.Documents.Open(PathName=spw_file,
                                   Visible=True,
                                   ReadOnly=True)
        
        if doc7 is not None:
            if xls_file is None:
                xls_file = spw_file[:-4] + '.xls'
            doc7.SaveAs(xls_file)
            print(f'Файл {xls_file} сохранён!')
        else:
            print('Не удалось сохранить документ')
            return ""
        
        print("Конвертация завершена")

        doc7.Close(const7.kdDoNotSaveChanges)

        return xls_file
    except Exception as e:
        print(f"Ошибка: {e}")
        return ""
    

def search_spw(directory):
    # Проверяем, существует ли директория и является ли она каталогом
    if not os.path.isdir(directory):
        print(f"Директория не найдена или не является каталогом: {directory}")
        return []
    
    spw_files = []
    # Проходим по директории и подпапкам
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.spw'):
                spw_files.append(os.path.abspath(os.path.join(root, file)))
    return spw_files


def do_a_path_for_xls(spws, output_dir):
    # Приводим output_dir к абсолютному пути, если он не абсолютный
    if not os.path.isabs(output_dir):
        output_dir = os.path.abspath(output_dir)
    # Если директории нет - создаем её
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    xls_files = []
    for spw in spws:
        # Извлекаем имя файла без расширения
        base_name = os.path.splitext(os.path.basename(spw))[0]
        # Формируем путь к xls файлу с таким же именем
        xls_path = os.path.join(output_dir, base_name + '.xls')
        xls_files.append(xls_path)
        
    return xls_files


def send_to_converter(spws, xlss):
    kompas_api = get_kompas_api7()
    zipped = zip(spws, xlss)
    all = len(spws)
    i = 1
    for spw, xls in zipped:
        convert_spw_to_xls(spw, xls, kompas_api)
        print(f'{i}/{all}')
        i+=1


if __name__ == "__main__":
    # Задайте путь к файлу без расширения
    # kompas_api = get_kompas_api7()
    # sp_file_path = r"O:\Python projects\Kompas Specs to xlsx\sample.spw"
    # convert_spw_to_xls(sp_file_path, kompas_api)
    dir = '415.1-Сварочный портал'
    spws = search_spw(dir)
    xlss = do_a_path_for_xls(spws, 'xls_out')
    send_to_converter(spws, xlss)