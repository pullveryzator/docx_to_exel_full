import json
import os
import pickle
import re
import warnings
from typing import Dict, List

import gdown
import numpy as np
import pandas as pd
import torch
from dotenv import load_dotenv
from transformers import AutoModel, AutoTokenizer

from constants import (BATCH_SIZE, DEST_FOLDER, GOOGLE_DRIVE_COMMON_PATH,
                       TASK_COLUMN, TASK_SHEET_NAME)
from decorators import validate_excel_file

warnings.filterwarnings("ignore", category=FutureWarning)

os.makedirs(DEST_FOLDER, exist_ok=True)
load_dotenv()
FOLDER_ID_ARTIFACTS = os.getenv("FOLDER_ID_ARTIFACTS")
FILE_URLS = {
    "confidence_thresholds.json": f"{GOOGLE_DRIVE_COMMON_PATH}={os.getenv('CONFIDENCE_THRESHOLDS_JSON_ID')}",
    "hierarchical_model_state.pt": f"{GOOGLE_DRIVE_COMMON_PATH}={os.getenv('HIERARCHICAL_MODEL_STATE_PT_ID')}",
    "tokenizer_config.json": f"{GOOGLE_DRIVE_COMMON_PATH}={os.getenv('TOKENIZER_CONFIG_JSON_ID')}",
    "special_tokens_map.json": f"{GOOGLE_DRIVE_COMMON_PATH}={os.getenv('SPECIAL_TOKENS_MAP_JSON_ID')}",
    "vocab.txt": f"{GOOGLE_DRIVE_COMMON_PATH}={os.getenv('VOCAB_TXT_ID')}",
    "tokenizer.json": f"{GOOGLE_DRIVE_COMMON_PATH}={os.getenv('TOKENIZER_JSON_ID')}",
    "label_maps.pkl": f"{GOOGLE_DRIVE_COMMON_PATH}={os.getenv('LABEL_MAPS_PKL_ID')}",
    "model_architecture_config.json": f"{GOOGLE_DRIVE_COMMON_PATH}={os.getenv('MODEL_ARCHITECTURE_CONFIG_JSON_ID')}",
    "topics.csv": f"{GOOGLE_DRIVE_COMMON_PATH}={os.getenv('TOPICS_CSV_ID')}",
}

def download_files():
    """Скачивает необходимые файлы"""
    print("Скачивание файлов модели...")
    for name, url in FILE_URLS.items():
        output_path = os.path.join(DEST_FOLDER, name)
        if not os.path.exists(output_path):
            print(f"Скачивание {name}...")
            gdown.download(url, output_path, quiet=False)
        else:
            print(f"Файл {name} уже существует, пропускаем")

download_files()

ART_PATH = lambda fn: os.path.join(DEST_FOLDER, fn)

print("\nЗагрузка конфигурации модели...")
with open(ART_PATH("model_architecture_config.json"), 'r', encoding='utf-8') as f:
    cfg = json.load(f)

IGNORE_INDEX = cfg.get("ignore_index")
num_classes_per_level = cfg.get("num_classes_per_level")
MODEL_NAME_HF = cfg.get("model_name")
TOKENIZER_MAX_LENGTH = cfg.get("tokenizer_max_length", 128)
MAX_LEVELS_CONFIG = cfg.get("max_levels_defined_in_script", 3)

print("Инициализация модели...")
tokenizer = AutoTokenizer.from_pretrained(DEST_FOLDER)

class HierarchicalClassifier(torch.nn.Module):
    def __init__(self, base_model_name: str, num_classes_list: List[int]) -> None:
        super().__init__()
        self.encoder = AutoModel.from_pretrained(base_model_name)
        self.classifiers = torch.nn.ModuleList()
        for n_classes in num_classes_list:
            if n_classes > 0:
                self.classifiers.append(
                    torch.nn.Linear(self.encoder.config.hidden_size, n_classes)
                )
            else:
                self.classifiers.append(None)

    def forward(self, input_ids, attention_mask) -> List[torch.Tensor]:
        pooled = self.encoder(input_ids=input_ids, attention_mask=attention_mask).pooler_output
        return [clf(pooled) for clf in self.classifiers if clf is not None]

model = HierarchicalClassifier(MODEL_NAME_HF, num_classes_per_level)
model.load_state_dict(torch.load(ART_PATH("hierarchical_model_state.pt"), map_location="cpu"))
model.eval()
DEVICE = torch.device("cuda" if torch.cuda.is_available() else "cpu")
model.to(DEVICE)

print("Загрузка дополнительных данных...")
with open(ART_PATH("label_maps.pkl"), "rb") as f:
    label_maps = pickle.load(f)

with open(ART_PATH("confidence_thresholds.json"), 'r', encoding='utf-8') as f:
    conf_thresholds = {int(k): v for k, v in json.load(f).items()}

TOPICS_DF = pd.read_csv(ART_PATH("topics.csv"))
ID2NAME = TOPICS_DF.set_index("id")["name"].to_dict()

# Функции предобработки и предсказания
def preprocess_latex_for_model(text: str) -> str:
    if not isinstance(text, str):
        return ""

    def process_formula(match):
        formula = match.group(1)
        substitutions = {
            r"\\frac": ' / ', r"\\cdot": ' · ', r"\\times": ' × ',
            r"\\div": ' ÷ ', r"\\leq": ' ≤ ', r"\\geq": ' ≥ ',
            r"\\neq": ' ≠ ', r"\\approx": ' ≈ ', r"\\rightarrow": ' → ',
            r"\\leftarrow": ' ← ', r"\\leftrightarrow": ' ↔ ',
            r"\\partial": ' ∂ ', r"\\infty": ' ∞ ', r"\\pi": ' π ',
            r"\\int": ' <INT> ', r"\\sum": ' <SUM> ', r"\\lim": ' <LIM> ',
            r"\\sqrt": ' <SQRT> ', r"[{}^_\\]": ' ', r"\s+": ' '
        }
        for pattern, replacement in substitutions.items():
            formula = re.sub(pattern, replacement, formula)
        return f" {formula.strip()} "

    processed_text = text
    patterns = [r"\\\((.*?)\\\)", r"\\\[(.*?)\\\]", r"\$(.*?)\$"]
    for pat in patterns:
        processed_text = re.sub(pat, process_formula, processed_text)
    return re.sub(r"\s+", ' ', processed_text).strip().lower()

def decode_prediction(pred_idx: int, pred_prob: float, original_level_idx: int) -> Dict:
    threshold = conf_thresholds.get(original_level_idx, 0.05)
    if pred_prob >= threshold and pred_idx != IGNORE_INDEX:
        if (original_level_idx in label_maps and 
            pred_idx in label_maps[original_level_idx]["index_to_id"]):
            raw_id = label_maps[original_level_idx]["index_to_id"][pred_idx]
            if isinstance(raw_id, str) and "NO_LABEL" in raw_id:
                return {"id": None, "name": "—"}
            try:
                topic_id = int(raw_id)
                return {
                    "id": topic_id,
                    "name": ID2NAME.get(topic_id, f"ID_{topic_id} (имя не найдено)")
                }
            except ValueError:
                return {"id": None, "name": f"ID_{raw_id} (некорректный формат)"}
    return {"id": None, "name": "—"}

@torch.inference_mode()
def predict_texts_hierarchical(texts: List[str]) -> List[List[Dict]]:
    if not texts:
        return []

    processed = [preprocess_latex_for_model(t) for t in texts]
    enc = tokenizer(
        processed,
        padding=True,
        truncation=True,
        max_length=TOKENIZER_MAX_LENGTH,
        return_tensors="pt"
    ).to(DEVICE)
    
    logits_list = model(input_ids=enc['input_ids'], attention_mask=enc['attention_mask'])
    
    results = []
    for i in range(len(texts)):
        preds = [{'id': None, 'name': '—'} for _ in range(MAX_LEVELS_CONFIG)]
        for mdl_idx, logits in enumerate(logits_list):
            lvl_idx = mdl_idx
            probs = torch.softmax(logits[i], dim=0)
            prob, idx = torch.max(probs, dim=0)
            preds[lvl_idx] = decode_prediction(idx.item(), prob.item(), lvl_idx)
        results.append(preds)
    return results


@validate_excel_file
def process_topics(output_file: str):
    print("\nИерархическая классификация математических задач...")
    try:
        with pd.ExcelFile(output_file) as xls:
            if TASK_SHEET_NAME not in xls.sheet_names:
                print(f"Лист '{TASK_SHEET_NAME}' не найден в файле!")
                
            df = pd.read_excel(xls, sheet_name=TASK_SHEET_NAME)
            
        if TASK_COLUMN not in df.columns:
            print(f"Колонка '{TASK_COLUMN}' не найдена. Доступные колонки: {list(df.columns)}")
            
        # Предсказание
        print(f"\nОбработка {len(df)} задач...")
        all_preds = []
        
        for i in range(0, len(df), BATCH_SIZE):
            batch = df[TASK_COLUMN].iloc[i:i+BATCH_SIZE].fillna("").tolist()
            all_preds.extend(predict_texts_hierarchical(batch))
            print(f"Обработано: {min(i+BATCH_SIZE, len(df))}/{len(df)}")
        
        # Добавление результатов в DataFrame
        for lvl in range(MAX_LEVELS_CONFIG):
            df[f'topic_id_lvl_{lvl+1}'] = [p[lvl]['id'] for p in all_preds]
            df[f'topic_name_{lvl+1}'] = [p[lvl]['name'] for p in all_preds]
        
        # Сохранение обратно в тот же файл
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=TASK_SHEET_NAME, index=False)
        
        print(f"\nРезультаты добавлены в файл: {output_file} (лист '{TASK_SHEET_NAME}')")
        print("Завершение работы...")
    except Exception as e:
        print(f"Ошибка: {str(e)}")
