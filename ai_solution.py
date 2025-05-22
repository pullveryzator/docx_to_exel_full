import os
import time
from typing import Optional

import pandas as pd
from dotenv import load_dotenv
from mistralai import Mistral

from constants import (ADVICE, ID_TASK_COLUMN, MISTRAL_MODEL, SOLUTION_COLUMN,
                       TASK_COLUMN, TASK_SHEET_NAME, TASK_SLICE_LENGTH,
                       TIME_SLEEP)
from utils import save_to_excel

load_dotenv()
MISTRAL_API_KEY: Optional[str] = os.getenv("MISTRAL_API_KEY")

def get_ai_solution(task_number: int, task_text: str) -> str:
    """Отправляет задачу в Mistral API и возвращает решение."""
    if not MISTRAL_API_KEY:
        return "Ошибка: не задан API ключ"
    
    client: Mistral = Mistral(api_key=MISTRAL_API_KEY)

    try:
        chat_response = client.chat.complete(
             model= MISTRAL_MODEL,
             messages = [
                 {"role": "user",
                  "content": f"{ADVICE} Задача №{task_number}: {task_text}",
                  },
            ]
        )
        solution: str = chat_response.choices[0].message.content
        return solution
    except Exception as e:
        print(f"Ошибка при запросе к API: {e}")
        return "Ошибка: не удалось получить решение"

def add_ai_solution_to_excel(file_path: str) -> None:
    """Обновляет Excel-файл, используя pandas и безопасное сохранение"""
    try:
        # Читаем данные из Excel
        df: pd.DataFrame = pd.read_excel(
            file_path, 
            engine='openpyxl', 
            sheet_name=TASK_SHEET_NAME
        )

        if SOLUTION_COLUMN not in df.columns:
            df[SOLUTION_COLUMN] = None

        if TASK_COLUMN not in df.columns:
            raise ValueError(f"Колонка '{TASK_COLUMN}' не найдена!")
        
        # Обрабатываем только строки, где есть задача и нет решения
        mask: pd.Series = df[TASK_COLUMN].notna() & df[SOLUTION_COLUMN].isna()
        tasks_to_process: pd.DataFrame = df[mask]
        
        if tasks_to_process.empty:
            print("Нет задач для обработки.")
            return
        
        print(f"Найдено {len(tasks_to_process)} задач для обработки...")

        # Создаем словарь для хранения всех листов
        with pd.ExcelFile(file_path) as excel:
            all_sheets = {sheet: pd.read_excel(excel, sheet_name=sheet) 
                        for sheet in excel.sheet_names}
        
        for index, row in tasks_to_process.iterrows():
            task: str = row[TASK_COLUMN]
            task_number: int = row[ID_TASK_COLUMN]
            print(f"Обработка задачи: {task_number} {task[:TASK_SLICE_LENGTH]}...")
            
            solution: str = get_ai_solution(task_number, task)
            df.at[index, SOLUTION_COLUMN] = solution

            all_sheets[TASK_SHEET_NAME] = df

            if index % 5 == 0:
                save_to_excel(all_sheets[TASK_SHEET_NAME], file_path, TASK_SHEET_NAME)
            
            time.sleep(TIME_SLEEP)

        success = save_to_excel(all_sheets[TASK_SHEET_NAME], file_path, TASK_SHEET_NAME)
        if success:
            print(f"Все решения записаны в файл {file_path}.")
        else:
            print("Ошибка при финальном сохранении!")
        
    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")
