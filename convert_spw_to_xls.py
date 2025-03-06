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


def convert_spw_to_xls(sp_file_path, kompas_api):

    

    try:
        # pythoncom.CoInitialize()
        # prog_id = "KOMPAS.Application.7"
        # kmpsApp = Dispatch(prog_id)
        # if kmpsApp is None:
        #     print("КОМПАС не установлен")
        #     return ""

        module7, api7, const7, app7 = kompas_api

        print("Начало конвертации...")
        
        # Формируем полный путь к spw файлу
        spw_file = sp_file_path + ".spw"
        if not os.path.exists(spw_file):
            print(f"Файл не найден: {spw_file}")
            return ""
        
        print("Открываем документ...")

        # Открываем документ

        doc7 = app7.Documents.Open(PathName=spw_file,
                                   Visible=True,
                                   ReadOnly=True)
        
        if doc7 is not None:
            out_file = sp_file_path + '.xls'
            doc7.SaveAs(out_file)
            print(f'Файл {out_file} сохранён!')
        else:
            print('Не удалось сохранить документ')
            return ""
        
        print("Конвертация завершена")

        doc7.Close(const7.kdDoNotSaveChanges)

        # kmpsdoc = kmpsApp.Documents.Open(spw_file)
        # if kmpsdoc is not None:
        #     # Отключаем показ сообщений (аналог ksHideMessageYes)
        #     kmpsApp.HideMessage = True
        #     out_file = sp_file_path + ".xls"
        #     kmpsdoc.SaveAs(out_file)
        #     print(f"Файл [{out_file}] сохранен!")
        # else:
        #     print("Не удалось открыть документ.")
        

        return out_file
    except Exception as e:
        print(f"Ошибка: {e}")
        return ""


if __name__ == "__main__":
    # Задайте путь к файлу без расширения
    sp_file_path = r"O:\Python projects\Kompas Specs to xlsx\sample"
    convert_spw_to_xls(sp_file_path)