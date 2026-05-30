import sqlite3
import os
import hashlib
import json
from typing import List, Dict, Optional

DB_PATH = os.path.join(os.getcwd(), "consolidation.db")

def get_connection():
    """Создает новое соединение для каждого вызова (безопасно для потоков)"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db():
    """Создает таблицы при первом запуске"""
    schema_path = os.path.join(os.getcwd(), "database_schema.sql")
    if os.path.exists(schema_path):
        with open(schema_path, "r", encoding="utf-8") as f:
            sql_script = f.read()
    else:
        raise FileNotFoundError("database_schema.sql не найден в директории проекта!")
    
    conn = get_connection()
    conn.executescript(sql_script)
    conn.close()

# ==================== CRUD: Пользователи ====================
class UserCRUD:
    @staticmethod
    def create(username: str, password: str, role: str = "operator") -> int:
        pwd_hash = hashlib.sha256(password.encode()).hexdigest()
        conn = get_connection()
        cur = conn.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)", (username, pwd_hash, role))
        conn.commit()
        uid = cur.lastrowid
        conn.close()
        return uid

    @staticmethod
    def authenticate(username: str, password: str) -> Optional[Dict]:
        pwd_hash = hashlib.sha256(password.encode()).hexdigest()
        conn = get_connection()
        cur = conn.execute("SELECT * FROM users WHERE username=? AND password_hash=?", (username, pwd_hash))
        row = cur.fetchone()
        conn.close()
        return dict(row) if row else None

    @staticmethod
    def update(user_id: int, **kwargs) -> bool:
        if not kwargs: return False
        fields = ", ".join(f"{k}=?" for k in kwargs.keys())
        values = list(kwargs.values()) + [user_id]
        conn = get_connection()
        conn.execute(f"UPDATE users SET {fields} WHERE id=?", values)
        conn.commit()
        conn.close()
        return True

    @staticmethod
    def delete(user_id: int) -> bool:
        conn = get_connection()
        conn.execute("DELETE FROM users WHERE id=?", (user_id,))
        conn.commit()
        conn.close()
        return True

    @staticmethod
    def seed_default():
        conn = get_connection()
        cur = conn.execute("SELECT COUNT(*) FROM users")
        if cur.fetchone()[0] == 0:
            UserCRUD.create("admin", "12345", "admin")
        conn.close()

# ==================== CRUD: Настройки порогов ====================
class ThresholdCRUD:
    @staticmethod
    def create(user_id: int, name: str, threshold: int, is_default: bool = False) -> int:
        conn = get_connection()
        cur = conn.execute("INSERT INTO threshold_settings (user_id, name, similarity_threshold, is_default) VALUES (?, ?, ?, ?)",
                           (user_id, name, threshold, is_default))
        conn.commit()
        tid = cur.lastrowid
        conn.close()
        return tid

    @staticmethod
    def get_all(user_id: int = None) -> List[Dict]:
        conn = get_connection()
        query = "SELECT * FROM threshold_settings"
        params = ()
        if user_id:
            query += " WHERE user_id=?"
            params = (user_id,)
        rows = conn.execute(query, params).fetchall()
        conn.close()
        return [dict(r) for r in rows]

    @staticmethod
    def update(threshold_id: int, **kwargs) -> bool:
        if not kwargs: return False
        fields = ", ".join(f"{k}=?" for k in kwargs.keys())
        values = list(kwargs.values()) + [threshold_id]
        conn = get_connection()
        conn.execute(f"UPDATE threshold_settings SET {fields} WHERE id=?", values)
        conn.commit()
        conn.close()
        return True

    @staticmethod
    def delete(threshold_id: int) -> bool:
        conn = get_connection()
        conn.execute("DELETE FROM threshold_settings WHERE id=?", (threshold_id,))
        conn.commit()
        conn.close()
        return True

# ==================== CRUD: Шаблоны сопоставления ====================
class TemplateCRUD:
    @staticmethod
    def create(name: str, category_key: str, synonyms: list, created_by: int) -> int:
        conn = get_connection()
        cur = conn.execute("INSERT INTO matching_templates (name, category_key, synonyms, created_by) VALUES (?, ?, ?, ?)",
                           (name, category_key, json.dumps(synonyms, ensure_ascii=False), created_by))
        conn.commit()
        tid = cur.lastrowid
        conn.close()
        return tid

    @staticmethod
    def get_all() -> List[Dict]:
        conn = get_connection()
        rows = conn.execute("SELECT * FROM matching_templates").fetchall()
        conn.close()
        res = []
        for r in rows:
            d = dict(r)
            d['synonyms'] = json.loads(d['synonyms']) if isinstance(d['synonyms'], str) else d['synonyms']
            res.append(d)
        return res

    @staticmethod
    def update(template_id: int, **kwargs) -> bool:
        if "synonyms" in kwargs and isinstance(kwargs["synonyms"], list):
            kwargs["synonyms"] = json.dumps(kwargs["synonyms"], ensure_ascii=False)
        if not kwargs: return False
        fields = ", ".join(f"{k}=?" for k in kwargs.keys())
        values = list(kwargs.values()) + [template_id]
        conn = get_connection()
        conn.execute(f"UPDATE matching_templates SET {fields} WHERE id=?", values)
        conn.commit()
        conn.close()
        return True

    @staticmethod
    def delete(template_id: int) -> bool:
        conn = get_connection()
        conn.execute("DELETE FROM matching_templates WHERE id=?", (template_id,))
        conn.commit()
        conn.close()
        return True

# ==================== CRUD: История обработок ====================
class HistoryCRUD:
    @staticmethod
    def create(user_id: int, folder_path: str, threshold: int, template_id: int = None) -> int:
        conn = get_connection()
        cur = conn.execute("INSERT INTO processing_history (user_id, folder_path, threshold_used, template_id) VALUES (?, ?, ?, ?)",
                           (user_id, folder_path, threshold, template_id))
        conn.commit()
        hid = cur.lastrowid
        conn.close()
        return hid

    @staticmethod
    def update_status(history_id: int, status: str, finished_at: str = None, result_path: str = None, error: str = None):
        conn = get_connection()
        conn.execute("UPDATE processing_history SET status=?, finished_at=?, result_file_path=?, error_message=? WHERE id=?",
                     (status, finished_at, result_path, error, history_id))
        conn.commit()
        conn.close()

    @staticmethod
    def get_all(user_id: int = None) -> List[Dict]:
        conn = get_connection()
        query = "SELECT * FROM processing_history"
        params = ()
        if user_id:
            query += " WHERE user_id=?"
            params = (user_id,)
        rows = conn.execute(query, params).fetchall()
        conn.close()
        return [dict(r) for r in rows]

    @staticmethod
    def delete(history_id: int) -> bool:
        conn = get_connection()
        conn.execute("DELETE FROM processing_history WHERE id=?", (history_id,))
        conn.commit()
        conn.close()
        return True