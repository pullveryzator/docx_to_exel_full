import re
from docx import Document
import pandas as pd
from time import sleep

def parse_answers(docx_path, output_file="answers.xlsx"):
    doc = Document(docx_path)
    answers = []
    found_answers = False
    
    # Регулярные выражения для разных форматов ответов
    answer_pattern = re.compile(
        r'(\d+)\.\s*'          # Номер основной задачи
        r'([а-я]\)|\d+\))?\s*'  # Подномер (а), б), 1), 2))
        r'([^;.]*[;.]?)'        # Текст ответа (до точки с запятой или точки)
    )
    
    for para in doc.paragraphs:
        text = para.text.strip()
        
        # Находим начало раздела с ответами
        if not found_answers:
            if "Ответы и советы" in text:
                found_answers = True
            continue
            
        # Пропускаем пустые строки
        if not text:
            continue
        if "Оглавление" in text:
            break
        # Обрабатываем ответы
        print(text)
        sleep(1)
        matches = answer_pattern.finditer(text)
        for match in matches:
            main_num = match.group(1)
            subtask = match.group(2)[0] if match.group(2) else ""
            answer_text = match.group(3).strip()
            
            # Формируем ID задачи
            if subtask:
                if subtask.isalpha():
                    task_id = f"{main_num}.{subtask}"
                else:
                    task_id = f"{main_num}.{subtask}"
            else:
                task_id = main_num
            
            # Очищаем текст ответа
            answer_text = re.sub(r'[;.]$', '', answer_text).strip()
            # print(task_id, answer_text)
            answers.append({
                'id_tasks_book': task_id,
                'answer': answer_text
            })
    
    # Создаем DataFrame и сохраняем
    df = pd.DataFrame(answers)
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Answers')
    
    print(f"Ответы успешно сохранены в {output_file}")
    return df

if __name__ == "__main__":
    parse_answers("tekstovye_zadachi_po_matematike.docx")