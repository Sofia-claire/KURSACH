import pandas as pd
from fuzzywuzzy import fuzz
import os
import random
import json
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import threading

# ==================== 0. КОНФИГУРАЦИЯ И МАППИНГ ====================
CONFIG_FILE = "config.json"

# Маппинг для перевода значений интерфейса в параметры customtkinter
APPEARANCE_MAP = {"Светлая": "Light", "Тёмная": "Dark", "Системная": "System"}
THEME_MAP = {"Синяя": "blue", "Зелёная": "green", "Тёмно-синяя": "dark-blue"}


def load_config():
    defaults = {"threshold": 85, "appearance": "Dark", "theme": "blue"}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
                return {**defaults, **cfg}
        except:
            pass
    return defaults


def save_config(cfg):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)


# Обратный маппинг (из конфига в русский текст для UI)
def get_rus_appearance(val):
    rev = {"Light": "Светлая", "Dark": "Тёмная", "System": "Системная"}
    return rev.get(val, "Тёмная")


def get_rus_theme(val):
    rev = {"blue": "Синяя", "green": "Зелёная", "dark-blue": "Тёмно-синяя"}
    return rev.get(val, "Синяя")


# ==================== 1. ТЕСТОВЫЕ ДАННЫЕ ====================
def create_realistic_files(folder_path='тестовые_отчёты_расширенные'):
    expense_categories = {
        'канцтовары': ['канцтовары', 'канцелярия', 'офисные товары', 'канц. принадлежности'],
        'бумага': ['бумага', 'бумага А4', 'бумага офисная', 'белая бумага'],
        'картриджи': ['картриджи', 'тонер', 'расходники принтера', 'картриджи для принтера'],
        'хозтовары': ['хозтовары', 'бытовая химия', 'средства уборки', 'салфетки'],
        'чай_кофе': ['чай', 'кофе', 'напитки', 'чай/кофе', 'печенье к чаю'],
        'мебель': ['мебель офисная', 'столы', 'стулья', 'кресла'],
        'обучение': ['обучение', 'курсы', 'тренинги', 'семинары'],
        'транспорт': ['такси', 'транспортные расходы', 'ГСМ', 'бензин'],
    }

    os.makedirs(folder_path, exist_ok=True)
    departments = ['Бухгалтерия', 'IT_отдел', 'Отдел_кадров', 'Юридический_отдел']
    all_files = []

    for dept in departments:
        num_transactions = random.randint(15, 30)
        data = []
        for _ in range(num_transactions):
            category = random.choice(list(expense_categories.keys()))
            name_variant = random.choice(expense_categories[category])
            amount = random.randint(500, 50000)
            date = datetime.now() - timedelta(days=random.randint(0, 30))
            data.append({
                '№ п/п': _ + 1,
                'Наименование (факт)': name_variant,
                'Категория (внутр)': category,
                'Сумма, руб': amount,
                'Дата операции': date.strftime('%d.%m.%Y'),
                'Подразделение': dept,
                'Статус': random.choice(['оплачено', 'ожидает', 'проведено']),
            })

        df = pd.DataFrame(data)
        filename = os.path.join(folder_path, f'{dept}_отчёт.xlsx')
        df.to_excel(filename, index=False)
        all_files.append(filename)
    return all_files


# ==================== 2. КОНСОЛИДАЦИЯ ====================
def smart_consolidate(folder_path, similarity_threshold=85, progress_callback=None):
    if progress_callback: progress_callback("🔄 Загружаем файлы из папки...", 0)

    all_data = []
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not excel_files: return None, None, "❌ Нет Excel файлов в папке!"

    for idx, file in enumerate(excel_files):
        df = pd.read_excel(os.path.join(folder_path, file))
        all_data.append(df)
        if progress_callback: progress_callback(f"📂 Загружен {file}", int((idx + 1) / len(excel_files) * 30))

    combined = pd.concat(all_data, ignore_index=True)
    if progress_callback: progress_callback(f"📊 Всего записей: {len(combined)}", 40)

    name_col = amount_col = None
    for col in combined.columns:
        cl = col.lower()
        if not name_col and any(k in cl for k in ['наимен', 'статья', 'назван', 'товар']): name_col = col
        if not amount_col and ('сумма' in cl or 'руб' in cl): amount_col = col

    if not name_col: name_col = combined.select_dtypes(include=['object']).columns[0]
    if not amount_col: amount_col = combined.select_dtypes(include=['number']).columns[0]

    if progress_callback: progress_callback(f"🎯 Группировка с порогом {similarity_threshold}%...", 50)

    groups = {}
    processed = [False] * len(combined)

    for i in range(len(combined)):
        if processed[i]: continue

        c_name = str(combined[name_col].iloc[i])
        c_amount = combined[amount_col].iloc[i]
        c_date = combined['Дата операции'].iloc[i] if 'Дата операции' in combined.columns else None
        c_dept = combined['Подразделение'].iloc[i] if 'Подразделение' in combined.columns else None

        groups[c_name] = {
            'каноническое_имя': c_name, 'сумма_общая': c_amount, 'количество_записей': 1,
            'оригинальные_названия': [c_name], 'отделы': set([c_dept]) if c_dept else set(),
            'суммы_по_отделам': {c_dept: c_amount} if c_dept else {},
            'диапазон_дат': [c_date] if c_date else []
        }
        processed[i] = True

        for j in range(i + 1, len(combined)):
            if processed[j]: continue
            compare_name = str(combined[name_col].iloc[j])
            if fuzz.ratio(c_name.lower(), compare_name.lower()) >= similarity_threshold:
                amount_j = combined[amount_col].iloc[j]
                groups[c_name]['сумма_общая'] += amount_j
                groups[c_name]['количество_записей'] += 1
                groups[c_name]['оригинальные_названия'].append(compare_name)

                dept_j = combined['Подразделение'].iloc[j] if 'Подразделение' in combined.columns else None
                if dept_j:
                    groups[c_name]['отделы'].add(dept_j)
                    groups[c_name]['суммы_по_отделам'][dept_j] = groups[c_name]['суммы_по_отделам'].get(dept_j,
                                                                                                        0) + amount_j

                date_j = combined['Дата операции'].iloc[j] if 'Дата операции' in combined.columns else None
                if date_j: groups[c_name]['диапазон_дат'].append(date_j)
                processed[j] = True

        if progress_callback and i % 20 == 0:
            progress_callback(f"🔄 Обработано {i}/{len(combined)} записей", 50 + int(i / len(combined) * 30))

    if progress_callback: progress_callback("📊 Формируем результаты...", 90)

    results = []
    for g in groups.values():
        date_range = "нет данных"
        if g['диапазон_дат']:
            try:
                dates = pd.to_datetime(g['диапазон_дат'], errors='coerce').dropna()
                if len(dates) > 0:
                    date_range = f"{dates.min().strftime('%d.%m.%Y')} - {dates.max().strftime('%d.%m.%Y')}"
            except:
                date_range = "разные даты"

        results.append({
            'Консолидированная статья': g['каноническое_имя'][:50] + ('...' if len(g['каноническое_имя']) > 50 else ''),
            'Общая сумма, руб': g['сумма_общая'],
            'Количество операций': g['количество_записей'],
            'Затронутые отделы': ', '.join(g['отделы']),
            'Макс. сумма по отделу': max(g['суммы_по_отделам'].values()) if g['суммы_по_отделам'] else 0,
            'Диапазон дат': date_range,
            'Варианты названий (пример)': ', '.join(g['оригинальные_названия'][:3]) + (
                '...' if len(g['оригинальные_названия']) > 3 else '')
        })

    return pd.DataFrame(results).sort_values('Общая сумма, руб', ascending=False), combined, None


# ==================== 3. GUI ====================
class ConsolidationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.config = load_config()
        # Применяем начальные настройки
        ctk.set_appearance_mode(self.config["appearance"])
        ctk.set_default_color_theme(self.config["theme"])

        self.title("Система консолидации отчётов")
        self.geometry("1200x700")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.result_df = self.raw_data = None
        self.folder_path = ctk.StringVar()
        self.similarity_threshold = ctk.IntVar(value=self.config["threshold"])
        self.status_text = ctk.StringVar(value="Готов к работе")

        # Русские значения для UI
        self.appearance_var = ctk.StringVar(value=get_rus_appearance(self.config["appearance"]))
        self.theme_var = ctk.StringVar(value=get_rus_theme(self.config["theme"]))

        self.create_navigation_frame()
        self.create_frames()
        self.select_frame_by_name("consolidator")

    def create_navigation_frame(self):
        self.navigation_frame = ctk.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(self.navigation_frame, text=" 📑 МЕНЮ", compound="left",
                     font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, padx=20, pady=20)

        self.home_button = ctk.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10,
                                         text="🏠 Главная", fg_color="transparent", text_color=("gray10", "gray90"),
                                         hover_color=("gray70", "gray30"), anchor="w",
                                         command=lambda: self.select_frame_by_name("home"))
        self.home_button.grid(row=1, column=0, sticky="ew")

        self.consolidator_button = ctk.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10,
                                                 text="📈 Консолидатор", fg_color="transparent",
                                                 text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                 anchor="w", command=lambda: self.select_frame_by_name("consolidator"))
        self.consolidator_button.grid(row=2, column=0, sticky="ew")

        self.settings_button = ctk.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10,
                                             text="⚙️ Настройки", fg_color="transparent",
                                             text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                             anchor="w", command=lambda: self.select_frame_by_name("settings"))
        self.settings_button.grid(row=3, column=0, sticky="ew")

        # Меню оформления (Только выбор, без ввода)
        self.appearance_mode_menu = ctk.CTkOptionMenu(
            self.navigation_frame,
            variable=self.appearance_var,
            values=["Светлая", "Тёмная", "Системная"],
            command=self.change_appearance_mode_event
        )
        self.appearance_mode_menu.grid(row=5, column=0, padx=20, pady=20, sticky="s")

    def select_frame_by_name(self, name):
        active = ("gray75", "gray25")
        self.home_button.configure(fg_color=active if name == "home" else "transparent")
        self.consolidator_button.configure(fg_color=active if name == "consolidator" else "transparent")
        self.settings_button.configure(fg_color=active if name == "settings" else "transparent")

        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()

        if name == "consolidator":
            self.consolidator_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.consolidator_frame.grid_forget()

        if name == "settings":
            self.settings_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.settings_frame.grid_forget()

    def create_frames(self):
        self.home_frame = self.create_home_frame()
        self.consolidator_frame = self.create_consolidator_frame()
        self.settings_frame = self.create_settings_frame()

    def create_home_frame(self):
        frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text="📖 О программе", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0,
                                                                                                 pady=(40, 10))
        info = ("Данная программа разработана для автоматической консолидации финансовых отчётов.\n\n"
                "🔹 Основные возможности:\n• Загрузка множества Excel-файлов.\n• Умная группировка строк (fuzzy matching).\n"
                "• Подсчёт сумм и статистика по отделам.\n• Экспорт результатов в Excel.\n\n🔹 Как работать:\n"
                "1. Перейдите во вкладку «Консолидатор».\n2. Выберите папку с файлами или создайте тестовые данные.\n"
                "3. Нажмите «Запустить консолидацию».\n4. Сохраните итоговый отчёт.")
        ctk.CTkLabel(frame, text=info, justify="left", wraplength=800, font=ctk.CTkFont(size=14)).grid(row=1, column=0,
                                                                                                       pady=30, padx=40,
                                                                                                       sticky="w")
        ctk.CTkLabel(frame, text="👤 Автор: Уливанова София", font=ctk.CTkFont(size=16, weight="bold"),
                     text_color="gray").grid(row=2, column=0, pady=40)
        return frame

    def create_consolidator_frame(self):
        frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")

        ctrl = ctk.CTkFrame(frame)
        ctrl.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(ctrl, text="📁 Папка с отчётами:", font=ctk.CTkFont(size=14)).pack(anchor="w")
        ctk.CTkEntry(ctrl, textvariable=self.folder_path).pack(fill="x", pady=5)
        ctk.CTkButton(ctrl, text="Выбрать папку", command=self.select_folder).pack(anchor="w", pady=5)

        btns = ctk.CTkFrame(frame)
        btns.pack(fill="x", padx=20, pady=10)

        # ✅ Убраны жёсткие HEX-цвета. Теперь кнопки автоматически
        # наследуют палитру выбранной темы (синяя/зелёная/тёмно-синяя)
        ctk.CTkButton(btns, text="📂 Тестовые данные", command=self.create_test_data).pack(side="left", padx=5)
        ctk.CTkButton(btns, text="🚀 Запуск", command=self.run_consolidation,
                      font=ctk.CTkFont(size=14, weight="bold")).pack(side="left", padx=5)

        st = ctk.CTkFrame(frame)
        st.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(st, text="Статус:", font=ctk.CTkFont(size=12)).pack(anchor="w")
        ctk.CTkLabel(st, textvariable=self.status_text, wraplength=900, justify="left").pack(anchor="w", pady=5)
        self.progress_bar = ctk.CTkProgressBar(st)
        self.progress_bar.pack(fill="x", pady=10)
        self.progress_bar.set(0)

        tbl = ctk.CTkFrame(frame)
        tbl.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tbl, show="headings", height=18)
        vsb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tbl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl.grid_rowconfigure(0, weight=1)
        tbl.grid_columnconfigure(0, weight=1)

        sb = ctk.CTkFrame(frame)
        sb.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(sb, text="💾 Сохранить", command=self.save_report, fg_color=("#e67e22", "#d35400"),
                      hover_color=("#c96a19", "#a84300")).pack(side="left", padx=5)
        ctk.CTkButton(sb, text="📊 Статистика", command=self.show_statistics, fg_color=("#9b59b6", "#8e44ad"),
                      hover_color=("#804999", "#6c3b8e")).pack(side="left", padx=5)
        return frame

    def create_settings_frame(self):
        frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        ctk.CTkLabel(frame, text="⚙️ Настройки системы", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=30)

        box = ctk.CTkFrame(frame, width=500)
        box.pack(pady=20)

        ctk.CTkLabel(box, text="Порог схожести группировки (%):").pack(anchor="w", padx=30, pady=(20, 5))
        ctk.CTkSlider(box, from_=60, to=100, variable=self.similarity_threshold, number_of_steps=40).pack(fill="x",
                                                                                                          padx=30,
                                                                                                          pady=5)
        ctk.CTkLabel(box, textvariable=self.similarity_threshold).pack()

        ctk.CTkLabel(box, text="Цветовая тема:").pack(anchor="w", padx=30, pady=(20, 5))
        # ✅ Исправление 3 & 4: state="readonly" запрещает ввод, значения на русском
        self.theme_combo = ctk.CTkComboBox(box, values=["Синяя", "Зелёная", "Тёмно-синяя"],
                                           variable=self.theme_var, state="readonly",
                                           command=self.change_theme)
        self.theme_combo.pack(fill="x", padx=30, pady=5)

        ctk.CTkLabel(box, text="💡 Настройки сохраняются автоматически при закрытии.", text_color="gray").pack(pady=40)
        return frame

    # ==================== ЛОГИКА ИНТЕРФЕЙСА ====================
    # ✅ Исправление 2: Мгновенное применение темы оформления
    def change_appearance_mode_event(self, choice):
        mode = APPEARANCE_MAP.get(choice, choice)
        ctk.set_appearance_mode(mode)
        self.config["appearance"] = mode
        save_config(self.config)

        # ✅ Принудительно перерисовываем интерфейс для мгновенного применения
        self.update()
        self.update_idletasks()

    def change_theme(self, choice):
        theme = THEME_MAP.get(choice, choice)
        ctk.set_default_color_theme(theme)
        self.config["theme"] = theme
        save_config(self.config)
        # Пересоздаём фреймы, чтобы новая тема применилась ко всем элементам
        self.rebuild_ui()

    def rebuild_ui(self):
        """Полностью пересоздаёт интерфейс для применения новой цветовой темы"""
        # 1. Запоминаем, какая вкладка сейчас открыта
        current = "consolidator"
        if hasattr(self, 'home_frame') and self.home_frame.winfo_ismapped():
            current = "home"
        elif hasattr(self, 'settings_frame') and self.settings_frame.winfo_ismapped():
            current = "settings"

        # 2. Удаляем ВСЕ фреймы, включая боковое меню (раньше оно не удалялось)
        for attr in ['navigation_frame', 'home_frame', 'consolidator_frame', 'settings_frame']:
            if hasattr(self, attr):
                getattr(self, attr).destroy()

        # 3. Пересоздаём их заново (они автоматически подхватят новую тему)
        self.create_navigation_frame()
        self.create_frames()
        self.select_frame_by_name(current)

        # 4. Принудительно обновляем отрисовку окна
        self.update()

    def on_closing(self):
        self.config["threshold"] = self.similarity_threshold.get()
        # Сохраняем английские значения в конфиг
        self.config["appearance"] = APPEARANCE_MAP.get(self.appearance_var.get(), "Dark")
        self.config["theme"] = THEME_MAP.get(self.theme_var.get(), "blue")
        save_config(self.config)
        self.destroy()

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder: self.folder_path.set(folder)

    def create_test_data(self):
        try:
            self.status_text.set("Создание тестовых данных...")
            self.progress_bar.set(0.3)
            target_dir = filedialog.askdirectory(title="Выберите место для папки с отчётами")
            if not target_dir: return
            folder_path = os.path.join(target_dir, "тестовые_отчёты")
            files = create_realistic_files(folder_path)
            self.progress_bar.set(1.0)
            self.status_text.set(f"✅ Создано {len(files)} файлов в '{folder_path}'")
            messagebox.showinfo("Успех", f"Файлы созданы!\n{folder_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.progress_bar.set(0)

    def run_consolidation(self):
        if not self.folder_path.get():
            messagebox.showwarning("Внимание", "Выберите папку!")
            return
        threading.Thread(target=self._consolidate_thread, daemon=True).start()

    def _consolidate_thread(self):
        try:
            self.after(0, lambda: self.status_text.set("🚀 Запуск..."))
            self.after(0, lambda: self.progress_bar.set(0.1))

            def update(msg, pct):
                self.after(0, lambda: self.status_text.set(msg))
                self.after(0, lambda: self.progress_bar.set(pct / 100))

            res, raw, err = smart_consolidate(self.folder_path.get(), self.similarity_threshold.get(), update)
            if err:
                self.after(0, lambda: messagebox.showerror("Ошибка", err))
                return
            self.result_df, self.raw_data = res, raw
            self.after(0, self.display_results)
            self.after(0, lambda: self.progress_bar.set(1.0))
            self.after(0, lambda: self.status_text.set(f"✅ Готово! {len(raw)} записей → {len(res)} групп."))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.after(0, lambda: self.progress_bar.set(0))

    def display_results(self):
        if self.result_df is None: return
        for item in self.tree.get_children(): self.tree.delete(item)
        cols = list(self.result_df.columns)
        self.tree["columns"] = cols
        widths = {'Консолидированная статья': 250, 'Общая сумма, руб': 150, 'Количество операций': 120,
                  'Затронутые отделы': 200, 'Макс. сумма по отделу': 150, 'Диапазон дат': 180,
                  'Варианты названий (пример)': 300}
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=widths.get(col, 150), anchor="w" if col != 'Общая сумма, руб' else "e")
        for _, row in self.result_df.head(100).iterrows():
            vals = [row[c] for c in cols]
            for i, c in enumerate(cols):
                if 'сумма' in c.lower(): vals[i] = f"{vals[i]:,.2f}"
            self.tree.insert("", "end", values=vals)

    def save_report(self):
        if self.result_df is None: messagebox.showwarning("Внимание", "Нет данных!")
        try:
            fp = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="отчёт.xlsx")
            if fp:
                with pd.ExcelWriter(fp, engine='openpyxl') as writer:
                    self.result_df.to_excel(writer, sheet_name='Итог', index=False)
                    if self.raw_data is not None: self.raw_data.to_excel(writer, sheet_name='Исходные', index=False)
                messagebox.showinfo("Успех", f"Сохранено в {fp}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def show_statistics(self):
        if self.result_df is None: messagebox.showwarning("Внимание", "Нет данных!")
        win = ctk.CTkToplevel(self)
        win.title("Статистика")
        win.geometry("500x400")
        txt = (f"📊 СТАТИСТИКА\nИсходных записей: {len(self.raw_data) if self.raw_data is not None else 0}\n"
               f"После консолидации: {len(self.result_df)}\n"
               f"Общая сумма: {self.result_df['Общая сумма, руб'].sum():,.2f} руб.")
        t = ctk.CTkTextbox(win)
        t.pack(fill="both", expand=True, padx=20, pady=20)
        t.insert("1.0", txt)
        t.configure(state="disabled")
        ctk.CTkButton(win, text="Закрыть", command=win.destroy).pack(pady=10)


if __name__ == "__main__":
    app = ConsolidationApp()
    app.mainloop()