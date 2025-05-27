import re


def fix_difficult_tasks_symb(input_num: str, input_task_part, index: int = None) -> str:
    """
    Обрабатывает маркер сложности (*) в задачах.
    
    Параметры:
        input_num - номер задачи (может содержать '*')
        input_task_part - текст задачи (str) или список текстов (list)
        index - если input_task_part список, индекс элемента для изменения
    
    Возвращает:
        Кортеж (очищенный номер, модифицированный текст)
    """
    # Очищаем номер от *
    has_star = '*' in input_num
    cleaned_num = input_num.replace('*', '') if has_star else input_num

    if isinstance(input_task_part, list) and index is not None:
        modified = input_task_part.copy()
        if has_star and 0 <= index < len(modified):
            modified[index] = '*' + modified[index]
    else:
        modified = ('*' if has_star else '') + str(input_task_part)
    
    return cleaned_num, modified

def fix_degree_to_star(value: str) -> str:
    """
    Исправляет знак градуса (°) на маркер сложности в задачах.
    
    Параметры:
        value: строка (задача или номер задачи)
    Возвращает:
        модифицированный текст
    """
    if '°' in value:
        modified = value.replace('°', '*')
        return modified
    return value


def filter_trailing_dots(text):
    """Удаляет возможные лишние точки в оглавлении перед номером страницы."""
    pattern = re.compile(r'^(\d+\.\d+)\.(.*?)(\.+)$')
    match = pattern.match(text)
    if match:
        section_num = match.group(1)
        section_text = match.group(2).strip()
        return f"{section_num}.{section_text}"
    
    return text
