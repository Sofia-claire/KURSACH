-- 1. Пользователи
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT DEFAULT 'user' CHECK(role IN ('user', 'admin', 'recovery')),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Одноразовые ключи доступа
CREATE TABLE IF NOT EXISTS registration_keys (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    key_code TEXT UNIQUE NOT NULL,
    key_type TEXT CHECK(key_type IN ('admin', 'user', 'recovery')) NOT NULL,
    is_used INTEGER DEFAULT 0,
    used_by_user_id INTEGER,
    created_by INTEGER REFERENCES users(id),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 3. Настройки порогов
CREATE TABLE IF NOT EXISTS threshold_settings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER REFERENCES users(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    similarity_threshold INTEGER DEFAULT 85 CHECK (similarity_threshold BETWEEN 50 AND 100),
    is_default INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 4. Шаблоны сопоставления
CREATE TABLE IF NOT EXISTS matching_templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    category_key TEXT NOT NULL,
    synonyms TEXT NOT NULL,
    created_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 5. История обработок
CREATE TABLE IF NOT EXISTS processing_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER REFERENCES users(id),
    folder_path TEXT NOT NULL,
    threshold_used INTEGER,
    template_id INTEGER REFERENCES matching_templates(id) ON DELETE SET NULL,
    status TEXT DEFAULT 'running',
    started_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    finished_at TIMESTAMP,
    result_file_path TEXT,
    error_message TEXT
);

-- 6. Логи загрузок файлов
CREATE TABLE IF NOT EXISTS upload_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    history_id INTEGER REFERENCES processing_history(id) ON DELETE CASCADE,
    filename TEXT NOT NULL,
    row_count INTEGER,
    status TEXT DEFAULT 'success',
    error_message TEXT,
    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 7. Результаты группировки
CREATE TABLE IF NOT EXISTS grouping_results (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    history_id INTEGER REFERENCES processing_history(id) ON DELETE CASCADE,
    canonical_name TEXT NOT NULL,
    total_amount REAL NOT NULL,
    operation_count INTEGER NOT NULL,
    departments TEXT,
    date_range TEXT,
    original_names TEXT
);

-- 8. Аудит действий
CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
    action TEXT NOT NULL,
    details TEXT,
    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);