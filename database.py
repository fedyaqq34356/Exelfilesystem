# database.py
import sqlite3
import os
from datetime import datetime
from pathlib import Path

DATABASE_NAME = "processed_files.db"

class Database:
    def __init__(self):
        self.db_path = DATABASE_NAME
        self.init_db()
        self.init_settings_table()  


    def init_db(self):

        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS processed_files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT UNIQUE NOT NULL,
                    file_name TEXT,
                    file_hash TEXT,
                    processed_at TEXT,
                    approver TEXT,
                    status TEXT DEFAULT 'DETECTED'
                )
            """)


            cursor.execute("""
                CREATE TABLE IF NOT EXISTS actions_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_name TEXT,
                    action TEXT,
                    user TEXT,
                    timestamp TEXT,
                    details TEXT
                )
            """)


            cursor.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
            """)


            cursor.execute("CREATE INDEX IF NOT EXISTS idx_file_path ON processed_files(file_path)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_timestamp ON actions_log(timestamp DESC)")

            conn.commit()

    def init_settings_table(self):

        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
            """)


    def get_setting(self, key: str, default=None):

        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT value FROM settings WHERE key = ?", (key,))
                row = cursor.fetchone()
                return row[0] if row else default
        except Exception as e:
            print(f"Помилка читання налаштування {key}: {e}")
            return default

    def set_setting(self, key: str, value: str):

        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("""
                    INSERT INTO settings (key, value)
                    VALUES (?, ?)
                    ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """, (key, value))
                conn.commit()
            return True
        except Exception as e:
            print(f"Помилка збереження налаштування {key}: {e}")
            return False

    def get_all_settings(self):

        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute("SELECT key, value FROM settings")
                return {row["key"]: row["value"] for row in cursor.fetchall()}
        except:
            return {}

    def get_file_hash(self, file_path):
        try:
            stat = os.stat(file_path)
            return f"{file_path}_{stat.st_size}_{stat.st_mtime}"
        except:
            return file_path

    def is_file_processed(self, file_path):
        file_hash = self.get_file_hash(file_path)
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT 1 FROM processed_files WHERE file_path = ?", (file_path,))
            if cursor.fetchone():
                return True
            cursor.execute("SELECT 1 FROM processed_files WHERE file_hash = ?", (file_hash,))
            return cursor.fetchone() is not None

    def add_processed_file(self, file_path, approver):
        try:
            file_name = Path(file_path).name
            file_hash = self.get_file_hash(file_path)
            timestamp = datetime.now().isoformat()

            with sqlite3.connect(self.db_path) as conn:
                conn.execute("""
                    INSERT OR REPLACE INTO processed_files
                    (file_path, file_name, file_hash, processed_at, approver, status)
                    VALUES (?, ?, ?, ?, ?, 'DETECTED')
                """, (file_path, file_name, file_hash, timestamp, approver))
            return True
        except Exception as e:
            print(f"Помилка додавання файлу в БД: {e}")
            return False

    def update_file_status(self, file_path, status):
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("UPDATE processed_files SET status = ? WHERE file_path = ?", (status, file_path))
            return True
        except Exception as e:
            print(f"Помилка оновлення статусу: {e}")
            return False

    def log_action(self, file_name, action, user, details=""):
        try:
            timestamp = datetime.now().isoformat()
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("""
                    INSERT INTO actions_log (file_name, action, user, timestamp, details)
                    VALUES (?, ?, ?, ?, ?)
                """, (file_name, action, user, timestamp, details))
            return True
        except Exception as e:
            print(f"Помилка логування: {e}")
            return False

    def get_recent_actions(self, limit=20):
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT file_name, action, user,
                           datetime(timestamp) as timestamp, details
                    FROM actions_log
                    ORDER BY timestamp DESC
                    LIMIT ?
                """, (limit,))
                return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            print(f"Помилка отримання логів: {e}")
            return []

    def get_all_processed_files(self):
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT file_name, approver, status,
                           datetime(processed_at) as processed_at
                    FROM processed_files
                    ORDER BY processed_at DESC
                """)
                return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            print(f"Помилка отримання файлів: {e}")
            return []