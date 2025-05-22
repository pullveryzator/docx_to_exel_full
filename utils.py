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