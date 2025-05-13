import os
from functools import wraps

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