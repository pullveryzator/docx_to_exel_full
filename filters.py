def fix_difficult_tasks_symb(input_num: str, input_task_part, index: int = None):
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
    
    # Обрабатываем текст
    if isinstance(input_task_part, list) and index is not None:
        modified = input_task_part.copy()  # Создаём копию, чтобы не менять исходный список
        if has_star and 0 <= index < len(modified):
            modified[index] = '*' + modified[index]
    else:
        modified = ('*' if has_star else '') + str(input_task_part)
    
    return cleaned_num, modified