import pandas as pd
from fuzzywuzzy import fuzz
import os
import random
import json
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import threading

from db_manager import (init_db, UserCRUD, ThresholdCRUD,
                        TemplateCRUD, HistoryCRUD)

# ==================== 0. КОНФИГУРАЦИЯ ====================
DEFAULT_CATEGORIES = {
    'канцтовары': ['канцтовары', 'канцелярия', 'офисные товары'],
    'бумага': ['бумага', 'бумага А4', 'офисная бумага'],
    'картриджи': ['картриджи', 'тонер', 'расходники принтера'],
    'хозтовары': ['хозтовары', 'бытовая химия', 'салфетки'],
    'чай_кофе': ['чай', 'кофе', 'напитки', 'печенье к чаю'],
    'мебель': ['столы', 'стулья', 'кресла', 'офисная мебель'],
    'обучение': ['курсы', 'тренинги', 'семинары'],
    'транспорт': ['такси', 'ГСМ', 'бензин', 'проезд']
}


# ==================== 1. ТЕСТОВЫЕ ДАННЫЕ ====================
def create_realistic_files(folder_path='тестовые_отчёты', categories=None):
    if categories is None: categories = DEFAULT_CATEGORIES
    os.makedirs(folder_path, exist_ok=True)
    departments = ['Бухгалтерия', 'IT_отдел', 'Отдел_кадров', 'Юридический_отдел']
    all_files = []
    for dept in departments:
        data = []
        for _ in range(random.randint(15, 30)):
            category = random.choice(list(categories.keys()))
            data.append({
                '№ п/п': _ + 1,
                'Наименование (факт)': random.choice(categories[category]),
                'Категория (внутр)': category,
                'Сумма, руб': random.randint(500, 50000),
                'Дата операции': (datetime.now() - timedelta(days=random.randint(0, 30))).strftime('%d.%m.%Y'),
                'Подразделение': dept,
                'Статус': random.choice(['оплачено', 'ожидает'])
            })
        df = pd.DataFrame(data)
        filename = os.path.join(folder_path, f'{dept}_отчёт.xlsx')
        df.to_excel(filename, index=False)
        all_files.append(filename)
    return all_files


# ==================== 2. КОНСОЛИДАЦИЯ ====================
def smart_consolidate(folder_path, similarity_threshold=85, progress_callback=None):
    if progress_callback: progress_callback("🔄 Загружаем файлы...", 0)
    all_data = []
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not excel_files: return None, None, "❌ Нет Excel файлов!"

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

    if progress_callback: progress_callback(f"🎯 Группировка ({similarity_threshold}%)...", 50)
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
            'суммы_по_отделам': {c_dept: c_amount} if c_dept else {}, 'диапазон_дат': [c_date] if c_date else []
        }
        processed[i] = True

        for j in range(i + 1, len(combined)):
            if processed[j]: continue
            if fuzz.ratio(c_name.lower(), str(combined[name_col].iloc[j]).lower()) >= similarity_threshold:
                amount_j = combined[amount_col].iloc[j]
                groups[c_name]['сумма_общая'] += amount_j
                groups[c_name]['количество_записей'] += 1
                groups[c_name]['оригинальные_названия'].append(str(combined[name_col].iloc[j]))
                dept_j = combined['Подразделение'].iloc[j] if 'Подразделение' in combined.columns else None
                if dept_j:
                    groups[c_name]['отделы'].add(dept_j)
                    groups[c_name]['суммы_по_отделам'][dept_j] = groups[c_name]['суммы_по_отделам'].get(dept_j,
                                                                                                        0) + amount_j
                date_j = combined['Дата операции'].iloc[j] if 'Дата операции' in combined.columns else None
                if date_j: groups[c_name]['диапазон_дат'].append(date_j)
                processed[j] = True

        if progress_callback and i % 20 == 0:
            progress_callback(f"🔄 Обработано {i}/{len(combined)}", 50 + int(i / len(combined) * 30))

    results = []
    for g in groups.values():
        date_range = "нет данных"
        if g['диапазон_дат']:
            try:
                dates = pd.to_datetime(g['диапазон_дат'], dayfirst=True, errors='coerce').dropna()
                if len(
                    dates) > 0: date_range = f"{dates.min().strftime('%d.%m.%Y')} - {dates.max().strftime('%d.%m.%Y')}"
            except:
                date_range = "разные даты"
        results.append({
            'Консолидированная статья': g['каноническое_имя'][:50],
            'Общая сумма, руб': g['сумма_общая'],
            'Количество операций': g['количество_записей'],
            'Затронутые отделы': ', '.join(g['отделы']),
            'Диапазон дат': date_range,
            'Варианты названий': ', '.join(g['оригинальные_названия'][:3])
        })
    return pd.DataFrame(results).sort_values('Общая сумма, руб', ascending=False), combined, None


# ==================== 3. GUI ====================
class ConsolidationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        init_db()
        UserCRUD.seed_default()

        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.title("Система консолидации отчётов (БД)")
        self.geometry("1250x750")

        self.current_user = None
        self.result_df = self.raw_data = None
        self.folder_path = ctk.StringVar()
        self.similarity_threshold = ctk.IntVar(value=85)
        self.status_text = ctk.StringVar(value="Готов к работе")
        self.expense_categories = DEFAULT_CATEGORIES.copy()
        self.cat_widgets = []

        self.show_login()

    def show_login(self):
        self.withdraw()
        win = ctk.CTkToplevel(self)
        win.title("🔐 Вход в систему")
        win.geometry("350x250")
        win.grab_set()
        win.transient(self)

        ctk.CTkLabel(win, text="Введите данные", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        usr = ctk.CTkEntry(win, placeholder_text="Логин (admin)", width=250)
        usr.pack(pady=5)
        pwd = ctk.CTkEntry(win, placeholder_text="Пароль (12345)", show="*", width=250)
        pwd.pack(pady=5)

        def try_login():
            u, p = usr.get(), pwd.get()
            res = UserCRUD.authenticate(u, p)
            if res:
                self.current_user = res
                win.destroy()
                self.deiconify()
                self.init_main_app()
            else:
                messagebox.showerror("Ошибка", "Неверный логин или пароль!")

        ctk.CTkButton(win, text="Войти", command=try_login).pack(pady=20)
        win.protocol("WM_DELETE_WINDOW", lambda: self.quit())

    def init_main_app(self):
        # Загрузка порогов и шаблонов из БД
        self.thresholds = ThresholdCRUD.get_all(self.current_user["id"])
        self.templates = TemplateCRUD.get_all()

        if self.thresholds:
            default_t = next((t for t in self.thresholds if t.get("is_default")), self.thresholds[0])
            self.similarity_threshold.set(default_t["similarity_threshold"])
        if self.templates:
            self.expense_categories = {t["category_key"]: t["synonyms"] for t in self.templates}

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.create_navigation_frame()
        self.create_frames()
        self.select_frame_by_name("consolidator")

    def create_navigation_frame(self):
        self.nav = ctk.CTkFrame(self, corner_radius=0)
        self.nav.grid(row=0, column=0, sticky="nsew")
        self.nav.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(self.nav, text=" 📑 МЕНЮ ", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, padx=20,
                                                                                               pady=20)

        self.home_btn = ctk.CTkButton(self.nav, text="🏠 Главная", fg_color="transparent",
                                      hover_color=("gray70", "gray30"), anchor="w",
                                      command=lambda: self.select_frame_by_name("home"))
        self.home_btn.grid(row=1, column=0, sticky="ew")

        self.con_btn = ctk.CTkButton(self.nav, text="📈 Консолидатор", fg_color="transparent",
                                     hover_color=("gray70", "gray30"), anchor="w",
                                     command=lambda: self.select_frame_by_name("consolidator"))
        self.con_btn.grid(row=2, column=0, sticky="ew")

        self.set_btn = ctk.CTkButton(self.nav, text="⚙️ Настройки", fg_color="transparent",
                                     hover_color=("gray70", "gray30"), anchor="w",
                                     command=lambda: self.select_frame_by_name("settings"))
        self.set_btn.grid(row=3, column=0, sticky="ew")

        self.hist_btn = ctk.CTkButton(self.nav, text="📜 История БД", fg_color="transparent",
                                      hover_color=("gray70", "gray30"), anchor="w",
                                      command=lambda: self.select_frame_by_name("history"))
        self.hist_btn.grid(row=4, column=0, sticky="ew")

        ctk.CTkLabel(self.nav, text=f"👤 {self.current_user['username']}", text_color="gray").grid(row=5, column=0,
                                                                                                  pady=10)
        ctk.CTkButton(self.nav, text="🚪 Выход", command=self.on_closing, fg_color="red", hover_color="darkred").grid(
            row=6, column=0, pady=10, padx=20, sticky="ew")

    def select_frame_by_name(self, name):
        # 1. Обновляем стиль кнопок
        btn_map = {"home": self.home_btn, "consolidator": self.con_btn, "settings": self.set_btn,
                   "history": self.hist_btn}
        for key, btn in btn_map.items():
            btn.configure(fg_color=("gray75", "gray25") if name == key else "transparent")

        # 2. Показываем/скрываем фреймы (безопасная логика)
        frame_map = {"home": self.home_frame, "consolidator": self.con_frame, "settings": self.set_frame,
                     "history": self.hist_frame}
        for key, frame in frame_map.items():
            if frame is not None:
                if name == key:
                    frame.grid(row=0, column=1, sticky="nsew")
                else:
                    frame.grid_forget()

        # 3. Доп. логика для вкладки истории
        if name == "history" and self.hist_frame is not None:
            self.load_history_ui()

    def create_frames(self):
        self.home_frame = self.create_home_frame()
        self.con_frame = self.create_con_frame()
        self.set_frame = self.create_settings_frame()
        self.hist_frame = self.create_history_frame()

    def create_home_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        f.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(f, text="📖 О программе", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0,
                                                                                             pady=(40, 10))
        info = "Автоматическая консолидации финансовых отчётов с нечётким сравнением.\n\n🔹 Хранение данных перенесено в SQLite.\n🔹 Реализованы CRUD для пользователей, порогов, шаблонов и истории."
        ctk.CTkLabel(f, text=info, justify="left", wraplength=800, font=ctk.CTkFont(size=14)).grid(row=1, column=0,
                                                                                                   pady=30, padx=40,
                                                                                                   sticky="w")
        return f

    def create_con_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        ctrl = ctk.CTkFrame(f)
        ctrl.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(ctrl, text="📁 Папка:").pack(anchor="w")
        ctk.CTkEntry(ctrl, textvariable=self.folder_path).pack(fill="x", pady=5)
        ctk.CTkButton(ctrl, text="Выбрать папку", command=lambda: self.folder_path.set(filedialog.askdirectory())).pack(
            anchor="w", pady=5)

        btns = ctk.CTkFrame(f)
        btns.pack(fill="x", padx=20, pady=10)
        ctk.CTkButton(btns, text="📂 Тестовые данные", command=self.create_test_data).pack(side="left", padx=5)
        ctk.CTkButton(btns, text="🚀 Запуск", command=self.run_consolidation,
                      font=ctk.CTkFont(size=14, weight="bold")).pack(side="left", padx=5)

        st = ctk.CTkFrame(f)
        st.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(st, text="Статус:").pack(anchor="w")
        ctk.CTkLabel(st, textvariable=self.status_text, wraplength=900).pack(anchor="w", pady=5)
        self.progress_bar = ctk.CTkProgressBar(st)
        self.progress_bar.pack(fill="x", pady=10)
        self.progress_bar.set(0)

        tbl = ctk.CTkFrame(f)
        tbl.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tbl, show="headings", height=18)
        ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview).grid(row=0, column=1, sticky="ns")
        self.tree.grid(row=0, column=0, sticky="nsew")
        tbl.grid_rowconfigure(0, weight=1)
        tbl.grid_columnconfigure(0, weight=1)

        sb = ctk.CTkFrame(f)
        sb.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(sb, text="💾 Сохранить", command=self.save_report).pack(side="left", padx=5)
        return f

    def create_settings_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        ctk.CTkLabel(f, text="⚙️ Настройки (БД)", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        box = ctk.CTkFrame(f, width=600)
        box.pack(pady=10, fill="both", expand=True)

        ctk.CTkLabel(box, text="Порог схожести (%):").pack(anchor="w", padx=30, pady=(10, 5))
        ctk.CTkSlider(box, from_=60, to=100, variable=self.similarity_threshold, number_of_steps=40).pack(fill="x",
                                                                                                          padx=30,
                                                                                                          pady=5)
        ctk.CTkLabel(box, textvariable=self.similarity_threshold).pack()

        ctk.CTkLabel(box, text="📂 Категории/Шаблоны:").pack(anchor="w", padx=30, pady=(20, 5))
        self.cat_scroll = ctk.CTkScrollableFrame(box, height=160)
        self.cat_scroll.pack(fill="x", padx=30, pady=5)
        ctk.CTkButton(box, text="➕ Добавить", command=self.add_cat_ui).pack(pady=(5, 10))
        for k, v in self.expense_categories.items(): self.add_cat_ui(k, ", ".join(v))
        return f

    def create_history_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        ctk.CTkLabel(f, text="📜 История обработок из БД", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        self.hist_tree = ttk.Treeview(f, show="headings", height=20)
        cols = ["ID", "Пользователь", "Папка", "Порог", "Статус", "Старт", "Файл результата"]
        self.hist_tree["columns"] = cols
        for c in cols:
            self.hist_tree.heading(c, text=c)
            self.hist_tree.column(c, width=150)
        self.hist_tree.pack(fill="both", expand=True, padx=20, pady=10)
        ttk.Scrollbar(f, orient="vertical", command=self.hist_tree.yview).pack(side="right", fill="y")
        ctk.CTkButton(f, text="🔄 Обновить", command=self.load_history_ui).pack(pady=10)
        return f

    def load_history_ui(self):
        if not self.hist_tree: return
        for i in self.hist_tree.get_children(): self.hist_tree.delete(i)
        rows = HistoryCRUD.get_all()
        for r in rows:
            vals = [r.get(k, "") for k in
                    ["id", "user_id", "folder_path", "threshold_used", "status", "started_at", "result_file_path"]]
            self.hist_tree.insert("", "end", values=vals)

    # ==================== UI Helpers ====================
    def add_cat_ui(self, name="", keywords=""):
        row = ctk.CTkFrame(self.cat_scroll)
        row.pack(fill="x", pady=2)
        ent = ctk.CTkEntry(row, placeholder_text="Категория", width=120);
        ent.pack(side="left", padx=2);
        ent.insert(0, name)
        txt = ctk.CTkTextbox(row, height=35, width=250);
        txt.pack(side="left", fill="x", expand=True, padx=2);
        txt.insert("1.0", keywords);
        txt.configure(wrap="word")
        ctk.CTkButton(row, text="🗑️", width=35, command=lambda r=row: self.remove_cat_ui(r, ent, txt)).pack(side="left")
        self.cat_widgets.append((ent, txt, row))

    def remove_cat_ui(self, row, ent, txt):
        row.destroy()
        self.cat_widgets = [(e, t, r) for e, t, r in self.cat_widgets if r != row]

    def save_settings_to_db(self):
        if not self.current_user: return
        ThresholdCRUD.create(self.current_user["id"], "Default", self.similarity_threshold.get(), is_default=True)
        for e, t, _ in self.cat_widgets:
            n = e.get().strip().lower().replace(" ", "_")
            kws = [x.strip() for x in t.get("1.0", "end-1c").split(",") if x.strip()]
            if n and kws:
                TemplateCRUD.create(n, n, kws, self.current_user["id"])

    # ==================== Логика ====================
    def create_test_data(self):
        try:
            target = filedialog.askdirectory(title="Куда создать тестовые данные?")
            if not target: return
            folder = os.path.join(target, "тестовые_отчёты")
            cats = {e.get().strip().lower().replace(" ", "_"): [x.strip() for x in t.get("1.0", "end-1c").split(",") if
                                                                x.strip()] for e, t, _ in self.cat_widgets}
            files = create_realistic_files(folder, cats or DEFAULT_CATEGORIES)
            messagebox.showinfo("Успех", f"Создано {len(files)} файлов в {folder}")
        except Exception as ex:
            messagebox.showerror("Ошибка", str(ex))

    def run_consolidation(self):
        if not self.folder_path.get(): return messagebox.showwarning("Внимание", "Выберите папку!")
        threading.Thread(target=self._consolidate_thread, daemon=True).start()

    def _consolidate_thread(self):
        try:
            self.after(0, lambda: self.status_text.set("🚀 Запуск..."))
            self.after(0, lambda: self.progress_bar.set(0.1))

            history_id = HistoryCRUD.create(self.current_user["id"], self.folder_path.get(),
                                            self.similarity_threshold.get())
            self.after(0, lambda: self.status_text.set("🔄 Обработка..."))

            def update(msg, pct):
                self.after(0, lambda: self.status_text.set(msg))
                self.after(0, lambda: self.progress_bar.set(pct / 100))

            res, raw, err = smart_consolidate(self.folder_path.get(), self.similarity_threshold.get(), update)
            if err: raise Exception(err)

            self.result_df, self.raw_data = res, raw
            HistoryCRUD.update_status(history_id, "completed", datetime.now().isoformat(), "результат.xlsx")
            self.after(0, self.display_results)
            self.after(0, lambda: self.status_text.set(f"✅ Готово! {len(raw)} → {len(res)} групп."))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
            self.after(0, lambda: self.progress_bar.set(0))

    def display_results(self):
        if self.result_df is None: return
        for i in self.tree.get_children(): self.tree.delete(i)
        cols = list(self.result_df.columns)
        self.tree["columns"] = cols
        for c in cols: self.tree.heading(c, text=c); self.tree.column(c, width=180)
        for _, r in self.result_df.head(100).iterrows():
            vals = [r[c] for c in cols]
            for i, c in enumerate(cols):
                if 'сумма' in c.lower(): vals[i] = f"{vals[i]:,.2f}"
            self.tree.insert("", "end", values=vals)

    def save_report(self):
        if self.result_df is None: return messagebox.showwarning("Внимание", "Нет данных!")
        try:
            fp = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="итог.xlsx")
            if fp:
                with pd.ExcelWriter(fp, engine='openpyxl') as w:
                    self.result_df.to_excel(w, sheet_name='Итог', index=False)
                messagebox.showinfo("Успех", f"Сохранено в {fp}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def on_closing(self):
        self.save_settings_to_db()
        self.destroy()


if __name__ == "__main__":
    app = ConsolidationApp()
    app.mainloop()