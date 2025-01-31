from openpyxl import load_workbook


def get_xls_sheet_quantity(file_path) -> int:
    """Функция возвращает количество листов в указанном xls файле.

    Args:
        file_path (str): Путь к xls файлу
    """
    try:
        # Открываем xls файл
        data_file = load_workbook(file_path)
    except FileNotFoundError:
        print(
            "Ошибка определения количества листов в Excel файле! "
            f"Файл {file_path} не найден. Программа будет завершена."
        )
        sys.exit(33)

    return len(data_file.sheetnames)
