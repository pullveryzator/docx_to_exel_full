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


def find_matching_paragraph(cleaned_text: str, toc: dict, trim_chars: int = 2) -> int:
    """
    Ищет наилучшее совпадение в словаре оглавления с возможностью обрезки символов.
    
    Параметры:
        cleaned_text - обработанный текст параграфа
        toc - словарь оглавления {название: id}
        trim_chars - количество символов для обрезки с конца (по умолчанию 2)
    """
    # Пробуем точное совпадение в первую очередь
    if cleaned_text in toc:
        return toc[cleaned_text]
    
    # Пробуем разные варианты обрезки
    for i in range(1, trim_chars + 1):
        truncated = cleaned_text[:-i] if len(cleaned_text) > i else cleaned_text
        for key in toc:
            if key.startswith(truncated):
                return toc[key]
    
    return None