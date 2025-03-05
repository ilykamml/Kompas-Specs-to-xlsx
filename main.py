#!/usr/bin/env python
import sys
from win32com.client import gencache

# Импорт констант и модулей из tlb-файлов SDK КОМПАС-3D
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)

def export_specification(input_file, output_file, command, show_param=False):
    """
    Экспортирует спецификацию из файла Компаса в XLSX с использованием интерфейса IConverter.
    
    Параметры:
      input_file  - полный путь к исходному файлу (например, чертежу с встроенной спецификацией)
      output_file - полный путь, куда будет сохранён XLSX файл
      command     - номер команды конвертации (определяется по документации API Компаса)
      show_param  - флаг, определяющий, нужно ли показывать диалог параметров (False, чтобы не показывать)
    """
    try:
        # Получаем запущенный экземпляр Компаса или создаём новый
        app = gencache.EnsureDispatch("Kompas.Application")
        app.Visible = True

        # Получаем интерфейс конвертера
        converter = app.Converter

        # Вызов метода Convert для экспорта спецификации в XLSX
        result = converter.Convert(input_file, output_file, command, show_param)
        if result == 1:
            print("Экспорт спецификации выполнен успешно:")
            print("  Исходный файл:", input_file)
            print("  XLSX файл:", output_file)
        else:
            print("Ошибка при экспорте спецификации. Код результата:", result)
        return result

    except Exception as e:
        print("Ошибка при работе с API Компаса:", e)
        return 0

if __name__ == "__main__":
    # Пример вызова: python export_spec.py "C:\path\to\drawing.cdw" "C:\output\specification.xlsx" 1001
    # Где 1001 – пример номера команды экспорта, его нужно заменить на актуальный.
    if len(sys.argv) < 4:
        print("Использование: python export_spec.py <input_file> <output_file> <command>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    try:
        command = int(sys.argv[3])
    except ValueError:
        print("Команда должна быть числовым значением.")
        sys.exit(1)

    export_specification(input_file, output_file, command)
