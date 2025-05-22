import os

import pandas as pd
from openpyxl import load_workbook


def save_to_excel(data, output_file:str, sheet_name:str):
    """Сохранение данных в Excel с автоматическим удалением существующего листа."""
    df = pd.DataFrame(data)
    mode = 'a' if os.path.exists(output_file) else 'w'
    if mode == 'a':
        book = load_workbook(output_file)
        if sheet_name in book.sheetnames:
            book.remove(book[sheet_name])
        book.save(output_file)
    with pd.ExcelWriter(output_file, engine='openpyxl', mode=mode) as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

from openpyxl import load_workbook


def reorder_sheets(output_file):
    """Переносит лист 'tasks' на первую позицию"""
    wb = load_workbook(output_file)
    if 'tasks' not in wb.sheetnames:
        print(f"Лист 'tasks' не найден в файле {output_file}")
        return None
    sheet = wb['tasks']
    wb.remove(sheet)
    wb._sheets.insert(0, sheet)
    wb.save(output_file)


def excel_to_dict(excel_file: str):
    """Получение словаря из Excel файла,
    в котором ключи - это текст главы, а значения - столбец ID.
    Включена обработка ошибки на отсутствие нужного файла."""
    try:
        df = pd.read_excel(excel_file, sheet_name='table_of_contents')
        
        result_dict = dict(zip(df['name'], df['id']))
        
        return result_dict
        
    except FileNotFoundError:
        print(f"Файл {excel_file} не найден!")
        return None