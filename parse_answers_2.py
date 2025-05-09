import re
from docx import Document
import pandas as pd

def parse_answers(docx_path, output_file="answers.xlsx"):
    doc = Document(docx_path)
    answers = []
    found_answers = False
    
    # Улучшенное регулярное выражение
    answer_block_re = re.compile(r'(\d+)\.(.*?)(?=\d+\.|\Z)', re.DOTALL)
    answer_item_re = re.compile(
        r'([а-я]\)|\d+\))?\s*([^;.]*[;.]?)',
        re.DOTALL
    )
    
    full_text = "\n".join([para.text for para in doc.paragraphs])
    
    # Находим начало  и конец ответов
    answers_start = full_text.find("Ответы и советы")
    answers_end = full_text.find('Оглавление')
    if answers_start == -1:
        print("Раздел с ответами не найден!")
        return None
    
    answers_text = full_text[answers_start:answers_end]
    
    # Обрабатываем блоки ответов
    for block in answer_block_re.finditer(answers_text):
        main_num = block.group(1)
        content = block.group(2).strip()

        # Разделяем ответы внутри блока
        answer_items = [a.strip() for a in content.split(';') if a.strip()]
        
        for item in answer_items:

            # Обрабатываем каждый ответ
            match = answer_item_re.match(item)
            if not match:
                continue
                
            subtask = match.group(1)
            answer_text = match.group(2).strip()

            # Формируем ID задачи
            if subtask:
                subtask = subtask.replace(')', '')
                if subtask.isalpha():
                    task_id = f"{main_num}.{subtask}"
                else:
                    task_id = f"{main_num}.{subtask}"
            else:
                task_id = main_num
            
            # Очищаем ответ
            answer_text = re.sub(r'[.,;]$', '', answer_text).strip()
            
            answers.append({
                'id_tasks_book': task_id,
                'answer': answer_text
            })
    
    # Создаем DataFrame
    df = pd.DataFrame(answers)
    
    # Сохраняем в Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='answers')
    
    print(f"Ответы сохранены в {output_file}")
    return df

if __name__ == "__main__":
    parse_answers("tekstovye_zadachi_po_matematike.docx")