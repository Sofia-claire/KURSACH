-- 1. Пользователи
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT DEFAULT 'operator',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Настройки порогов
CREATE TABLE IF NOT EXISTS threshold_settings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    name TEXT NOT NULL,
    similarity_threshold INTEGER DEFAULT 85 CHECK (similarity_threshold BETWEEN 50 AND 100),
    is_default INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
);

-- 3. Шаблоны сопоставления
CREATE TABLE IF NOT EXISTS matching_templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    category_key TEXT NOT NULL,
    synonyms TEXT NOT NULL,
    created_by INTEGER,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(created_by) REFERENCES users(id) ON DELETE SET NULL
);

-- 4. История обработок
CREATE TABLE IF NOT EXISTS processing_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    folder_path TEXT NOT NULL,
    threshold_used INTEGER,
    template_id INTEGER,
    status TEXT DEFAULT 'running',
    started_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    finished_at TIMESTAMP,
    result_file_path TEXT,
    error_message TEXT,
    FOREIGN KEY(user_id) REFERENCES users(id),
    FOREIGN KEY(template_id) REFERENCES matching_templates(id)
);

-- 5. Логи загрузок файлов
CREATE TABLE IF NOT EXISTS upload_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    history_id INTEGER,
    filename TEXT NOT NULL,
    row_count INTEGER,
    status TEXT DEFAULT 'success',
    error_message TEXT,
    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(history_id) REFERENCES processing_history(id) ON DELETE CASCADE
);

-- 6. Результаты группировки
CREATE TABLE IF NOT EXISTS grouping_results (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    history_id INTEGER,
    canonical_name TEXT NOT NULL,
    total_amount REAL NOT NULL,
    operation_count INTEGER NOT NULL,
    departments TEXT,
    date_range TEXT,
    original_names TEXT,
    FOREIGN KEY(history_id) REFERENCES processing_history(id) ON DELETE CASCADE
);

-- 7. Параметры обработок
CREATE TABLE IF NOT EXISTS processing_parameters (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    history_id INTEGER,
    param_name TEXT NOT NULL,
    param_value TEXT,
    FOREIGN KEY(history_id) REFERENCES processing_history(id) ON DELETE CASCADE
);

-- 8. Аудит действий
CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    action TEXT NOT NULL,
    details TEXT,
    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL
);