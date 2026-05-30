import sqlite3
import os
import hashlib
import json
from typing import List, Dict, Optional

DB_PATH = os.path.join(os.getcwd(), "consolidation.db")

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db():
    schema_path = os.path.join(os.getcwd(), "database_schema.sql")
    if os.path.exists(schema_path):
        with open(schema_path, "r", encoding="utf-8") as f:
            sql_script = f.read()
    else:
        raise FileNotFoundError("database_schema.sql не найден!")
    conn = get_connection()
    conn.executescript(sql_script)
    conn.close()

# ==================== CRUD: Пользователи ====================
class UserCRUD:
    @staticmethod
    def create(username: str, password: str, role: str = "user") -> int:
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
    def exists(username: str) -> bool:
        conn = get_connection()
        cur = conn.execute("SELECT 1 FROM users WHERE username=?", (username,))
        exists = cur.fetchone() is not None
        conn.close()
        return exists

    @staticmethod
    def seed_admin():
        conn = get_connection()
        cur = conn.execute("SELECT COUNT(*) FROM users WHERE role='admin'")
        if cur.fetchone()[0] == 0:
            UserCRUD.create("uli", "11037", "admin")
        conn.close()

    # ✅ НОВЫЕ МЕТОДЫ ДЛЯ АДМИНКИ
    @staticmethod
    def get_all() -> List[Dict]:
        conn = get_connection()
        rows = conn.execute("SELECT id, username, role, created_at FROM users ORDER BY id").fetchall()
        conn.close()
        return [dict(r) for r in rows]

    @staticmethod
    def update_password(user_id: int, new_password: str) -> bool:
        pwd_hash = hashlib.sha256(new_password.encode()).hexdigest()
        conn = get_connection()
        conn.execute("UPDATE users SET password_hash=? WHERE id=?", (pwd_hash, user_id))
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

# ==================== CRUD: Ключи доступа ====================
class KeyCRUD:
    @staticmethod
    def create_key(code: str, key_type: str, created_by: int) -> int:
        conn = get_connection()
        cur = conn.execute("INSERT INTO registration_keys (key_code, key_type, created_by) VALUES (?, ?, ?)", (code, key_type, created_by))
        conn.commit()
        kid = cur.lastrowid
        conn.close()
        return kid

    @staticmethod
    def validate_key(code: str) -> Optional[str]:
        conn = get_connection()
        cur = conn.execute("SELECT key_type, is_used FROM registration_keys WHERE key_code=?", (code,))
        row = cur.fetchone()
        conn.close()
        if row and row["is_used"] == 0:
            return row["key_type"]
        return None

    @staticmethod
    def consume_key(code: str, user_id: int):
        conn = get_connection()
        conn.execute("UPDATE registration_keys SET is_used=1, used_by_user_id=? WHERE key_code=?", (user_id, code))
        conn.commit()
        conn.close()

    @staticmethod
    def get_unused_keys(created_by: int) -> List[Dict]:
        conn = get_connection()
        rows = conn.execute("SELECT * FROM registration_keys WHERE created_by=? AND is_used=0 ORDER BY created_at DESC", (created_by,)).fetchall()
        conn.close()
        return [dict(r) for r in rows]

# ==================== CRUD: Настройки порогов ====================
class ThresholdCRUD:
    @staticmethod
    def create(user_id: int, name: str, threshold: int, is_default: bool = False) -> int:
        conn = get_connection()
        cur = conn.execute("INSERT INTO threshold_settings (user_id, name, similarity_threshold, is_default) VALUES (?, ?, ?, ?)", (user_id, name, threshold, is_default))
        conn.commit()
        tid = cur.lastrowid; conn.close(); return tid

    @staticmethod
    def get_all(user_id: int = None) -> List[Dict]:
        conn = get_connection()
        q, p = "SELECT * FROM threshold_settings", ()
        if user_id: q += " WHERE user_id=?"; p = (user_id,)
        rows = conn.execute(q, p).fetchall(); conn.close()
        return [dict(r) for r in rows]

# ==================== CRUD: Шаблоны ====================
class TemplateCRUD:
    @staticmethod
    def create(name: str, key: str, syns: list, uid: int) -> int:
        conn = get_connection()
        cur = conn.execute("INSERT INTO matching_templates (name, category_key, synonyms, created_by) VALUES (?, ?, ?, ?)", (name, key, json.dumps(syns, ensure_ascii=False), uid))
        conn.commit(); tid = cur.lastrowid; conn.close(); return tid

    @staticmethod
    def get_all() -> List[Dict]:
        conn = get_connection(); rows = conn.execute("SELECT * FROM matching_templates").fetchall(); conn.close()
        return [{**dict(r), "synonyms": json.loads(r["synonyms"])} for r in rows]

# ==================== CRUD: История ====================
class HistoryCRUD:
    @staticmethod
    def create(uid: int, fp: str, th: int, tid: int=None) -> int:
        conn = get_connection()
        cur = conn.execute("INSERT INTO processing_history (user_id, folder_path, threshold_used, template_id) VALUES (?, ?, ?, ?)", (uid, fp, th, tid))
        conn.commit(); hid = cur.lastrowid; conn.close(); return hid

    @staticmethod
    def update_status(hid: int, status: str, finished: str=None, res: str=None, err: str=None):
        conn = get_connection()
        conn.execute("UPDATE processing_history SET status=?, finished_at=?, result_file_path=?, error_message=? WHERE id=?", (status, finished, res, err, hid))
        conn.commit(); conn.close()

    @staticmethod
    def get_all(uid: int=None) -> List[Dict]:
        conn = get_connection()
        q, p = "SELECT * FROM processing_history", ()
        if uid: q += " WHERE user_id=?"; p = (uid,)
        rows = conn.execute(q, p).fetchall(); conn.close()
        return [dict(r) for r in rows]