import os
from pathlib import Path
from dotenv import load_dotenv
from database import Database

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()

def parse_chat_id(value):
    try:
        if not value or value in ("0", "None", ""):
            return None
        return int(value)
    except:
        return None

CHAT_ID_DIRECTOR = parse_chat_id(os.getenv("CHAT_ID_DIRECTOR"))
CHAT_ID_FINDIRECTOR = parse_chat_id(os.getenv("CHAT_ID_FINDIRECTOR"))

CHECK_INTERVAL = int(os.getenv("CHECK_INTERVAL", "10"))
FILE_SETTLE_TIME = int(os.getenv("FILE_SETTLE_TIME", "5"))

_db = Database()

DEFAULT_PATHS = {
    "director_folder":    r"C:\Users\Yevhen\OneDrive\Документы\Облік замовлень\Директор",
    "findirector_folder": r"C:\Users\Yevhen\OneDrive\Документы\Облік замовлень\Фін директор",
    "accountant_folder":  r"C:\Users\Yevhen\OneDrive\Документы\Облік замовлень\Бухгалтер",
    "cashier_folder":     r"C:\Users\Yevhen\OneDrive\Документы\Облік замовлень\Касир",
    "rejected_folder":    r"C:\Users\Yevhen\OneDrive\Документы\Облік замовлень\Відхилені",
}

_PATHS = {}
for key, default in DEFAULT_PATHS.items():
    saved = _db.get_setting(key)
    _PATHS[key] = saved if saved is not None else default
    if saved is None:
        _db.set_setting(key, default)

PATHS = _PATHS

def get_path(key: str) -> str | None:
    return PATHS.get(key)

def update_path(key: str, new_path: str) -> bool:
    if key not in PATHS:
        return False

    new_path = new_path.strip()
    if not new_path:
        return False

    try:
        Path(new_path).mkdir(parents=True, exist_ok=True)
        PATHS[key] = new_path
        success = _db.set_setting(key, new_path)
        if success:
            print(f"Шлях збережено в БД: {key} → {new_path}")
        return success
    except Exception as e:
        print(f"Помилка збереження шляху {key}: {e}")
        return False

def ensure_folders_exist():
    for path in PATHS.values():
        if path:
            try:
                Path(path).mkdir(parents=True, exist_ok=True)
            except Exception as e:
                print(f"Не вдалося створити папку {path}: {e}")

ensure_folders_exist()

if __name__ != "__main__":
    print("\nНалаштування шляхів (з бази даних):")
    for k, v in PATHS.items():
        status = "OK" if Path(v).exists() else "НЕМАЄ"
        print(f"   {k:20} → {v} [{status}]")
    print()
