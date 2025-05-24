import os
from functools import wraps

import pandas as pd
from docx import Document


def validate_docx_file(func):
    """Декоратор для проверки валидности DOCX файла"""
    @wraps(func)
    def wrapper(input_file, *args, **kwargs):
        if not (os.path.isfile(input_file) and input_file.lower().endswith('.docx')):
            print(f"Ошибка: Файл {input_file} не является DOCX или не существует")
            return None
        
        try:
            Document(input_file)
        except Exception as e:
            print(f"Ошибка: Файл {input_file} поврежден или не является DOCX: {str(e)}")
            return None

        try:
            return func(input_file, *args, **kwargs)
        except Exception as e:
            print(f"Ошибка при обработке файла {input_file}: {str(e)}")
            return None
    
    return wrapper

def validate_excel_file(func):
    """Декоратор для проверки валидности Excel файла (XLSX)"""
    @wraps(func)
    def wrapper(input_file, *args, **kwargs):
        if not (os.path.isfile(input_file) and input_file.lower().endswith(('.xlsx', '.xls'))):
            print(f"Ошибка: Файл {input_file} не является Excel или не существует")
            return None
        
        try:
            with pd.ExcelFile(input_file) as xls:
                if not xls.sheet_names:
                    print(f"Ошибка: Файл {input_file} не содержит листов")
                    return None

        except Exception as e:
            print(f"Ошибка: Файл {input_file} поврежден или не является Excel: {str(e)}")
            return None

        try:
            return func(input_file, *args, **kwargs)
        except Exception as e:
            print(f"Ошибка при обработке файла {input_file}: {str(e)}")
            return None
    
    return wrapper