"""
История генерации отчётов.
"""

import json
from datetime import datetime

from config import REPORT_HISTORY_FILE
from constants import ALL_GODY


def load_report_history() -> list:
    if not REPORT_HISTORY_FILE.exists():
        return []
    try:
        with open(REPORT_HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, Exception) as e:
        print(f"Ошибка загрузки истории: {e}")
        return []


def save_report_history(history: list):
    try:
        with open(REPORT_HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2, default=str)
    except Exception as e:
        print(f"Ошибка сохранения истории: {e}")


def add_report_to_history(teacher_fio: str, year: str, file_path: str,
                          generation_time: str, network_path: str = ""):
    history = load_report_history()
    history.append({
        "teacher_fio": teacher_fio,
        "year": year,
        "file_path": file_path,
        "network_path": network_path,
        "generation_time": generation_time,
        "date": datetime.now().strftime("%Y-%m-%d")
    })
    if len(history) > 1000:
        history = history[-1000:]
    save_report_history(history)


def get_last_report_for_teacher(teacher_fio: str, year: str = None) -> dict:
    history = load_report_history()
    filtered = [h for h in history if h.get("teacher_fio") == teacher_fio]
    if year and year != ALL_GODY:
        filtered = [h for h in filtered if h.get("year") == year]
    if filtered:
        filtered.sort(key=lambda x: x.get("generation_time", ""), reverse=True)
        return filtered[0]
    return None