import os
import re

import pandas as pd
from docx import Document
from openpyxl import load_workbook

from constants import (AUTHOR_DATA, CLASSES, DOCX_PATH, LEVEL, OUTPUT_FILE,
                       TOPIC_ID)
from decorators import validate_docx_file
from filters import fix_degree_to_star, fix_difficult_tasks_symb
from ai_solution import add_ai_solution_to_excel


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


@validate_docx_file
def parse_toc_to_excel(input_file:str, output_file:str):
    """Парсинг оглавления в Excel."""
    doc = Document(input_file)
    sections = []
    last_main_section_id = 0
    found_toc = False

    section_pattern = re.compile(r'^(\d+)\.\s+(.*?)\s*\d*$')
    subsection_pattern = re.compile(r'^(\d+\.\d+)\.\s+(.*?)\s*\d*$')
    
    for para in doc.paragraphs:
        text = para.text.strip()

        if not found_toc:
            if text.lower() == "оглавление":
                found_toc = True
            continue
        
        if found_toc and not text:
            break

        main_section_match = section_pattern.match(text)
        subsection_match = subsection_pattern.match(text)
        
        if main_section_match and not '.' in main_section_match.group(1):
            section_num = main_section_match.group(1)
            section_name = main_section_match.group(2)
            
            sections.append({
                'id': len(sections) + 1,
                'name': f"{section_num}.{section_name}",
                'parent': 0
            })
            last_main_section_id = len(sections)
            
        elif subsection_match:
            subsection_num = subsection_match.group(1)
            subsection_name = subsection_match.group(2)
            
            sections.append({
                'id': len(sections) + 1,
                'name': f"{subsection_num}.{subsection_name}",
                'parent': last_main_section_id
            })

    save_to_excel(data=sections, output_file=output_file, sheet_name='table_of_contents')


@validate_docx_file
def parse_docx_to_excel(input_file:str, output_file:str):
    """Парсинг текста задач в Excel."""
    doc = Document(input_file)
    data = []
    toc = excel_to_dict(output_file)

    for para in doc.paragraphs:
        text = para.text.strip()
        text = fix_degree_to_star(text)
        if "Ответы и советы" in text:
            break

        cleaned_text = re.sub(r'(\d+\.\d*\.?)\s+', r'\1', text)
        if cleaned_text in toc:
            paragraph_id = toc.get(cleaned_text)
            continue

        if '\t' in text:
            parts = text.split('\t', 1)
        id_part = parts[0].strip()
        task_part = parts[1].strip()
        if '.' in id_part:
            main_num = id_part
            subtask_parts = task_part.split('\t', 1)
            if len(subtask_parts) == 1:
                main_num, subtask_parts = fix_difficult_tasks_symb(main_num, subtask_parts, 0)
                data.append({
                    'id_tasks_book': main_num,
                    'task': subtask_parts[0],
                    'answer': '',
                    'paragraph': paragraph_id,
                    'classes': CLASSES,
                    'topic_id': TOPIC_ID,
                    'level': LEVEL
                })
            if len(subtask_parts) == 2:
                slave_num = subtask_parts[0].replace(')', '')
                slave_num, subtask_parts = fix_difficult_tasks_symb(slave_num, subtask_parts, 1)
                data.append({
                'id_tasks_book': main_num + slave_num,
                'task': subtask_parts[1],
                'answer': '',
                'paragraph': paragraph_id,
                'classes': CLASSES,
                'topic_id': TOPIC_ID,
                'level': LEVEL
            })
                
        slave_num = id_part.replace(')', '')
        if main_num.strip() != slave_num.strip():
            slave_num, task_part = fix_difficult_tasks_symb(slave_num, task_part)
            data.append({
                'id_tasks_book': main_num + slave_num,
                'task': task_part,
                'answer': '',
                'paragraph': paragraph_id,
                'classes': CLASSES,
                'topic_id': TOPIC_ID,
                'level': LEVEL
            })

    save_to_excel(data=data, output_file=output_file, sheet_name='tasks')


@validate_docx_file
def parse_answers(docx_path: str, output_file: str):
    """Парсинг ответов в Excel."""
    doc = Document(docx_path)
    answers_dict = {}
    answer_block_re = re.compile(r'(\d+)\.(.*?)(?=\d+\.|\Z)', re.DOTALL)
    answer_item_re = re.compile(
        r'([а-я]\)|\d+\))?\s*([^;.]*[;.]?)',
        re.DOTALL
    )

    full_text = "\n".join([para.text for para in doc.paragraphs])

    answers_start = full_text.find("Ответы и советы")
    answers_end = full_text.find('Оглавление')
    answers_text = full_text[answers_start:answers_end]

    for block in answer_block_re.finditer(answers_text):
        main_num = block.group(1)
        content = block.group(2).strip()

        answer_items = [a.strip() for a in content.split(';') if a.strip()]
        
        for item in answer_items:
            match = answer_item_re.match(item)
            if not match:
                continue
                
            subtask = match.group(1)
            answer_text = match.group(2).strip()

            if subtask:
                subtask = subtask.replace(')', '')
                task_id = f"{main_num}.{subtask}" if subtask.isalpha() else f"{main_num}.{subtask}"
            else:
                task_id = f"{main_num}."

            answers_dict[task_id] = re.sub(r'[.,;]$', '', answer_text).strip()

    tasks_df = pd.read_excel(output_file, sheet_name='tasks')
    tasks_df['answer'] = tasks_df['id_tasks_book'].map(answers_dict).fillna('Отсутствует')
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        tasks_df.to_excel(writer, index=False, sheet_name='tasks')


def add_author(author_data: list[dict], output_file:str):
    """Добавление к Excel файлу листа авторов."""
    save_to_excel(data=author_data, output_file=output_file, sheet_name='author')


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

if __name__ == "__main__":

    parse_toc_to_excel(DOCX_PATH, OUTPUT_FILE)
    parse_docx_to_excel(DOCX_PATH, OUTPUT_FILE)
    add_author(AUTHOR_DATA, OUTPUT_FILE)
    parse_answers(DOCX_PATH, OUTPUT_FILE)
    add_ai_solution_to_excel(OUTPUT_FILE)
    