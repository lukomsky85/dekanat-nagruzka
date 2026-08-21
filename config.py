"""
Пути к файлам и работа с конфигурацией.
"""

import sys
import json
from pathlib import Path


def get_app_dir() -> Path:
    """Возвращает папку приложения (работает и в .py, и в .exe)."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent


APP_DIR = get_app_dir()
CONFIG_FILE = APP_DIR / "config.json"
REPORT_HISTORY_FILE = APP_DIR / "report_history.json"


# ---------- Конфигурация ----------

def load_config() -> dict:
    default_config = {
        "server": r"DEKANAT100\SQLEXPRESS",
        "database": "Деканат",
        "auth_type": "windows",
        "username": "",
        "password": "",
        "save_password": False,
        "window_geometry": "820x680",
        "use_network_folder": False,
        "network_folder_path": r"\\server\reports",
        "local_folder_path": str(APP_DIR),
        "numbering": "original",
        "mode": "По семестрам (осенний + весенний)",
        # --- НОВЫЕ ПОЛЯ ДЛЯ ШАБЛОНА ---
        "template_institution_lines": "ОМСКИЙ ИНСТИТУТ ВОДНОГО ТРАНСПОРТА - ФИЛИАЛ\nФЕДЕРАЛЬНОГО ГОСУДАРСТВЕННОГО БЮДЖЕТНОГО ОБРАЗОВАТЕЛЬНОГО УЧРЕЖДЕНИЯ ВЫСШЕГО ОБРАЗОВАНИЯ\n«СИБИРСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ ВОДНОГО ТРАНСПОРТА»",
        "template_approver_title": "Заместитель директора по\nстратегическому развитию",
        "template_approver_name": "Д.А. Токарев",
        "template_umo_name": "Никишкин А.С.",
        # --- ПРЕСЕТЫ ФИЛЬТРОВ ВЫБОРКИ ДАННЫХ ---
        "query_templates": {},
        # --- ОБНОВЛЕНИЯ ---
        "update_manifest_url": None,  # None -> берётся UPDATE_MANIFEST_URL из constants.py
        "auto_check_updates": True,
    }
    if not CONFIG_FILE.exists():
        return default_config
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
            return {**default_config, **config}
    except (json.JSONDecodeError, Exception) as e:
        print(f"Ошибка загрузки конфига: {e}")
        return default_config


def save_config(config: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Ошибка сохранения конфига: {e}")