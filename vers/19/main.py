import pandas as pd
from fuzzywuzzy import fuzz
import os, random, json, secrets
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import threading
from PIL import Image
from pystray import Icon, Menu, MenuItem
from db_manager import (init_db, UserCRUD, ThresholdCRUD,
                        TemplateCRUD, HistoryCRUD, KeyCRUD)

# ==================== 0. КОНФИГУРАЦИЯ И МАППИНГ ====================
APPEARANCE_MAP = {"Светлая": "Light", "Тёмная": "Dark", "Системная": "System"}
THEME_MAP = {"Синяя": "blue", "Зелёная": "green", "Тёмно-синяя": "dark-blue"}
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


def get_rus_appearance(val):
    rev = {"Light": "Светлая", "Dark": "Тёмная", "System": "Системная"}
    return rev.get(val, "Тёмная")


def get_rus_theme(val):
    rev = {"blue": "Синяя", "green": "Зелёная", "dark-blue": "Тёмно-синяя"}
    return rev.get(val, "Синяя")


# ==================== 1. ТЕСТОВЫЕ ДАННЫЕ ====================
def create_realistic_files(folder_path='тестовые_отчёты', categories=None):
    if categories is None: categories = DEFAULT_CATEGORIES
    os.makedirs(folder_path, exist_ok=True)
    depts = ['Бухгалтерия', 'IT_отдел', 'Отдел_кадров', 'Юридический_отдел']
    files = []
    for d in depts:
        data = []
        for _ in range(random.randint(15, 30)):
            cat = random.choice(list(categories.keys()))
            data.append({
                '№ п/п': _ + 1, 'Наименование (факт)': random.choice(categories[cat]),
                'Категория (внутр)': cat, 'Сумма, руб': random.randint(500, 50000),
                'Дата операции': (datetime.now() - timedelta(days=random.randint(0, 30))).strftime('%d.%m.%Y'),
                'Подразделение': d, 'Статус': random.choice(['оплачено', 'ожидает'])
            })
        df = pd.DataFrame(data)
        fn = os.path.join(folder_path, f'{d}_отчёт.xlsx')
        df.to_excel(fn, index=False)
        files.append(fn)
    return files


# ==================== 2. КОНСОЛИДАЦИЯ ====================
def smart_consolidate(folder_path, threshold=85, cb=None):
    if cb: cb("🔄 Загрузка...", 0)
    data = []
    ex = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    if not ex: return None, None, "❌ Нет Excel файлов!"
    for i, f in enumerate(ex):
        data.append(pd.read_excel(os.path.join(folder_path, f)))
        if cb: cb(f"📂 {f}", int((i + 1) / len(ex) * 30))
    comb = pd.concat(data, ignore_index=True)
    if cb: cb(f"📊 Записей: {len(comb)}", 40)
    nc = ac = None
    for c in comb.columns:
        cl = c.lower()
        if not nc and any(k in cl for k in ['наимен', 'статья', 'назван', 'товар']): nc = c
        if not ac and ('сумма' in cl or 'руб' in cl): ac = c
    if not nc: nc = comb.select_dtypes(include=['object']).columns[0]
    if not ac: ac = comb.select_dtypes(include=['number']).columns[0]
    if cb: cb(f"🎯 Группировка ({threshold}%)...", 50)
    grps, proc = {}, [False] * len(comb)
    for i in range(len(comb)):
        if proc[i]: continue
        cn, ca = str(comb[nc].iloc[i]), comb[ac].iloc[i]
        cd = comb['Дата операции'].iloc[i] if 'Дата операции' in comb.columns else None
        cp = comb['Подразделение'].iloc[i] if 'Подразделение' in comb.columns else None
        grps[cn] = {'каноническое_имя': cn, 'сумма_общая': ca, 'количество_записей': 1,
                    'оригинальные_названия': [cn], 'отделы': set([cp]) if cp else set(),
                    'суммы_по_отделам': {cp: ca} if cp else {}, 'диапазон_дат': [cd] if cd else []}
        proc[i] = True
        for j in range(i + 1, len(comb)):
            if proc[j]: continue
            if fuzz.ratio(cn.lower(), str(comb[nc].iloc[j]).lower()) >= threshold:
                aj = comb[ac].iloc[j]
                grps[cn]['сумма_общая'] += aj
                grps[cn]['количество_записей'] += 1
                grps[cn]['оригинальные_названия'].append(str(comb[nc].iloc[j]))
                dp = comb['Подразделение'].iloc[j] if 'Подразделение' in comb.columns else None
                if dp: grps[cn]['отделы'].add(dp); grps[cn]['суммы_по_отделам'][dp] = grps[cn]['суммы_по_отделам'].get(
                    dp, 0) + aj
                dt = comb['Дата операции'].iloc[j] if 'Дата операции' in comb.columns else None
                if dt: grps[cn]['диапазон_дат'].append(dt)
                proc[j] = True
        if cb and i % 20 == 0: cb(f"🔄 {i}/{len(comb)}", 50 + int(i / len(comb) * 30))
    res = []
    for g in grps.values():
        dr = "нет данных"
        if g['диапазон_дат']:
            try:
                ds = pd.to_datetime(g['диапазон_дат'], dayfirst=True, errors='coerce').dropna()
                if len(ds) > 0: dr = f"{ds.min().strftime('%d.%m.%Y')} - {ds.max().strftime('%d.%m.%Y')}"
            except:
                dr = "разные даты"
        res.append({'Консолидированная статья': g['каноническое_имя'][:50], 'Общая сумма, руб': g['сумма_общая'],
                    'Количество операций': g['количество_записей'], 'Затронутые отделы': ', '.join(g['отделы']),
                    'Диапазон дат': dr, 'Варианты названий': ', '.join(g['оригинальные_названия'][:3])})
    return pd.DataFrame(res).sort_values('Общая сумма, руб', ascending=False), comb, None


# ==================== ОКНО АВТОРИЗАЦИИ ====================
class AuthDialog(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.result = None
        self.title("🔐 Вход / Регистрация")
        self.geometry("380x480")
        self.transient(master);
        self.grab_set()
        self.mode = "login"
        self.build_ui()
        self.wait_window()

    def build_ui(self):
        ctk.CTkLabel(self, text="Система консолидации отчётов", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=15)
        self.usr_ent = ctk.CTkEntry(self, placeholder_text="Логин", width=300)
        self.usr_ent.pack(pady=5)
        self.pwd_ent = ctk.CTkEntry(self, placeholder_text="Пароль", show="*", width=300)
        self.pwd_ent.pack(pady=5)

        self.frame_reg_fields = ctk.CTkFrame(self)
        self.key_ent = ctk.CTkEntry(self.frame_reg_fields, placeholder_text="🔑 Ключ доступа", width=300)
        self.key_ent.pack(pady=5)
        self.pwd2_ent = ctk.CTkEntry(self.frame_reg_fields, placeholder_text="Повторите пароль", show="*", width=300)
        self.pwd2_ent.pack(pady=5)
        self.frame_reg_fields.pack_forget()

        self.btn_submit = ctk.CTkButton(self, text="Войти", command=self.submit)
        self.btn_submit.pack(pady=15, fill="x", padx=40)
        self.toggle_btn = ctk.CTkButton(self, text="Нет аккаунта? Регистрация", fg_color="transparent",
                                        text_color=("gray50", "gray80"), hover_color=("gray70", "gray30"),
                                        command=self.toggle_mode)
        self.toggle_btn.pack(pady=5)
        ctk.CTkLabel(self, text="По умолчанию: admin / admin123", text_color="gray").pack(pady=10)

    def toggle_mode(self):
        self.mode = "register" if self.mode == "login" else "login"
        if self.mode == "login":
            self.frame_reg_fields.pack_forget()
            self.toggle_btn.configure(text="Нет аккаунта? Регистрация")
            self.title("🔐 Вход / Регистрация")
            self.btn_submit.configure(text="Войти")
        else:
            self.frame_reg_fields.pack(fill="x", padx=20, pady=5)
            self.toggle_btn.configure(text="Уже есть аккаунт? Вход")
            self.title("📝 Регистрация")
            self.btn_submit.configure(text="Зарегистрироваться")

    def submit(self):
        u, p = self.usr_ent.get().strip(), self.pwd_ent.get()
        if self.mode == "login":
            res = UserCRUD.authenticate(u, p)
            if res:
                self.result = res; self.destroy()
            else:
                messagebox.showerror("Ошибка", "Неверный логин или пароль!")
        else:
            p2 = self.pwd2_ent.get()
            k = self.key_ent.get().strip().upper()
            if not u or not p or not p2 or not k: return messagebox.showwarning("Внимание", "Заполните все поля!")
            if p != p2: return messagebox.showerror("Ошибка", "Пароли не совпадают!")
            if UserCRUD.exists(u): return messagebox.showerror("Ошибка", "Логин занят!")
            k_type = KeyCRUD.validate_key(k)
            if not k_type: return messagebox.showerror("Ошибка", "Ключ недействителен или уже использован!")
            uid = UserCRUD.create(u, p, k_type)
            KeyCRUD.consume_key(k, uid)
            messagebox.showinfo("Успех", "Аккаунт создан! Выполняется вход...")
            self.result = {"id": uid, "username": u, "role": k_type}
            self.destroy()


# ==================== GUI ПРИЛОЖЕНИЯ ====================
class ConsolidationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        init_db()
        UserCRUD.seed_admin()
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.title("Система консолидации отчётов (БД)")
        self.geometry("1250x750")

        # ✅ ИКОНКА ПРИЛОЖЕНИЯ
        self.update_window_icon()

        # ✅ ТРЕЙ
        self.tray_icon = None
        self.tray_thread = None
        self.setup_tray()

        self.auth_win = AuthDialog(self)
        if not self.auth_win.result: self.quit(); return
        self.current_user = self.auth_win.result

        self.result_df = self.raw_data = None
        self.folder_path = ctk.StringVar()
        self.similarity_threshold = ctk.IntVar(value=85)
        self.status_text = ctk.StringVar(value="Готов к работе")
        self.expense_categories = DEFAULT_CATEGORIES.copy()
        self.cat_widgets = []

        self.appearance_var = ctk.StringVar(value="Тёмная")
        self.theme_var = ctk.StringVar(value="Синяя")

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.init_main_app()

    def update_window_icon(self):
        """Обновление иконки окна в зависимости от темы"""
        try:
            appearance = ctk.get_appearance_mode()
            if appearance == "Dark":
                icon_path = os.path.join("ico", "dark.ico")
            else:
                icon_path = os.path.join("ico", "light.ico")

            if os.path.exists(icon_path):
                if os.name == 'nt':  # Windows
                    self.iconbitmap(icon_path)
                else:  # Linux/Mac
                    img = Image.open(icon_path)
                    photo = ctk.CTkImage(light_image=img, dark_image=img, size=(32, 32))
                    self.iconphoto(True, photo)
        except Exception as e:
            print(f"⚠️ Не удалось установить иконку: {e}")

    def setup_tray(self):
        """Настройка трея"""
        try:
            # Определяем иконку для трея
            appearance = ctk.get_appearance_mode()
            if appearance == "Dark":
                tray_icon_path = os.path.join("ico", "dark.ico")
            else:
                tray_icon_path = os.path.join("ico", "light.ico")

            if os.path.exists(tray_icon_path):
                tray_icon_image = Image.open(tray_icon_path)
            else:
                # Создаём заглушку если иконки нет
                tray_icon_image = Image.new('RGB', (64, 64), color='blue')

            # Меню трея
            menu = Menu(
                MenuItem('📂 Развернуть', self.show_from_tray, default=True),
                MenuItem('⚙️ Настройки', self.open_settings_from_tray),
                MenuItem('─', None),  # Разделитель
                MenuItem('❌ Выход', self.exit_app)
            )

            self.tray_icon = Icon("ConsolidationApp", tray_icon_image, "Система консолидации", menu)

        except Exception as e:
            print(f"⚠️ Не удалось создать трей: {e}")
            self.tray_icon = None

    def start_tray(self):
        """Запуск трея в отдельном потоке"""
        if self.tray_icon:
            self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
            self.tray_thread.start()

    def hide_to_tray(self):
        """Скрыть окно в трей"""
        self.withdraw()

    def show_from_tray(self, icon=None, item=None):
        """Показать окно из трея"""
        self.deiconify()
        self.lift()
        self.focus_force()

    def open_settings_from_tray(self, icon=None, item=None):
        """Открыть настройки из трея"""
        self.show_from_tray()
        self.select_frame_by_name("settings")

    def exit_app(self, icon=None, item=None):
        """Полный выход из приложения"""
        if self.tray_icon:
            self.tray_icon.stop()
        self.save_settings_to_db()
        self.quit()
        self.destroy()

    def on_closing(self):
        """Обработка закрытия окна - сворачивание в трей"""
        if messagebox.askyesno("Подтверждение", "Закрыть программу?\n(Отмена - свернуть в трей)"):
            self.exit_app()
        else:
            self.hide_to_tray()

    def init_main_app(self):
        self.thresholds = ThresholdCRUD.get_all(self.current_user["id"])
        self.templates = TemplateCRUD.get_all()
        if self.thresholds:
            def_t = next((t for t in self.thresholds if t.get("is_default")), self.thresholds[0])
            self.similarity_threshold.set(def_t["similarity_threshold"])
        if self.templates:
            self.expense_categories = {t["category_key"]: t["synonyms"] for t in self.templates}

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.create_navigation_frame()
        self.create_frames()
        self.select_frame_by_name("consolidator")

        # ✅ ЗАПУСК ТРЕЯ
        self.after(1000, self.start_tray)

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
        self.hist_btn = ctk.CTkButton(self.nav, text="📜 История", fg_color="transparent",
                                      hover_color=("gray70", "gray30"), anchor="w",
                                      command=lambda: self.select_frame_by_name("history"))
        self.hist_btn.grid(row=4, column=0, sticky="ew")

        self.admin_btn = self.users_btn = None
        if self.current_user["role"] == "admin":
            self.admin_btn = ctk.CTkButton(self.nav, text="🔑 Ключи", fg_color="transparent",
                                           hover_color=("gray70", "gray30"), anchor="w",
                                           command=lambda: self.select_frame_by_name("admin"))
            self.admin_btn.grid(row=5, column=0, sticky="ew")
            self.users_btn = ctk.CTkButton(self.nav, text="👥 Пользователи", fg_color="transparent",
                                           hover_color=("gray70", "gray30"), anchor="w",
                                           command=lambda: self.select_frame_by_name("users"))
            self.users_btn.grid(row=6, column=0, sticky="ew")

        row_idx = 7 if self.current_user["role"] == "admin" else 5
        ctk.CTkLabel(self.nav, text=f"👤 {self.current_user['username']} ({self.current_user['role']})",
                     text_color="gray").grid(row=row_idx, column=0, pady=10)
        ctk.CTkButton(self.nav, text="🚪 Выход", command=self.on_closing, fg_color="red", hover_color="darkred").grid(
            row=row_idx + 1, column=0, pady=10, padx=20, sticky="ew")

    def select_frame_by_name(self, name):
        btn_map = {"home": self.home_btn, "consolidator": self.con_btn, "settings": self.set_btn,
                   "history": self.hist_btn}
        if self.admin_btn: btn_map["admin"] = self.admin_btn
        if self.users_btn: btn_map["users"] = self.users_btn
        for k, b in btn_map.items():
            if b: b.configure(fg_color=("gray75", "gray25") if name == k else "transparent")

        frame_map = {"home": self.home_frame, "consolidator": self.con_frame, "settings": self.set_frame,
                     "history": self.hist_frame}
        if self.admin_btn: frame_map["admin"] = self.admin_frame
        if self.users_btn: frame_map["users"] = self.users_frame
        for k, f in frame_map.items():
            if f: f.grid(row=0, column=1, sticky="nsew") if name == k else f.grid_forget()

        if name == "history": self.load_history_ui()
        if name == "admin": self.load_admin_ui()
        if name == "users": self.load_users_ui()

    def create_frames(self):
        self.home_frame = self.create_home_frame()
        self.con_frame = self.create_con_frame()
        self.set_frame = self.create_settings_frame()
        self.hist_frame = self.create_history_frame()
        self.admin_frame = self.create_admin_frame() if self.current_user["role"] == "admin" else None
        self.users_frame = self.create_users_frame() if self.current_user["role"] == "admin" else None

    def create_home_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        f.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(f, text="📖 О программе", font=ctk.CTkFont(size=28, weight="bold")).grid(row=0, column=0,
                                                                                             pady=(40, 20))
        info = ("Данная программа разработана для автоматической консолидации финансовых отчётов.\n\n"
                "🔹 Основные возможности:\n• Загрузка множества Excel-файлов из папки\n"
                "• Умная группировка строк (fuzzy matching)\n• Подсчёт сумм и статистика по отделам\n"
                "• Экспорт результатов в Excel\n\n🔹 Как работать:\n"
                "1. Перейдите во вкладку «Консолидатор»\n2. Выберите папку с файлами или создайте тестовые данные\n"
                "3. Нажмите «Запустить консолидацию»\n4. Сохраните итоговый отчёт")
        ctk.CTkLabel(f, text=info, justify="left", wraplength=850, font=ctk.CTkFont(size=14)).grid(row=1, column=0,
                                                                                                   pady=30, padx=50,
                                                                                                   sticky="w")
        separator = ctk.CTkFrame(f, height=2, fg_color=("gray70", "gray30"))
        separator.grid(row=2, column=0, sticky="ew", padx=50, pady=20)
        author_info = ("👤 Автор: Уливанова София\n\n🔐 Система включает:\n"
                       "• Авторизацию по ключам доступа\n• Ролевую модель (Admin/User)\n"
                       "• Админ-панель управления пользователями\n• Историю обработок в SQLite")
        ctk.CTkLabel(f, text=author_info, justify="left", wraplength=850, font=ctk.CTkFont(size=13),
                     text_color=("gray20", "gray70")).grid(row=3, column=0, pady=(0, 40), padx=50, sticky="w")
        return f

    def create_con_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        ctrl = ctk.CTkFrame(f);
        ctrl.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(ctrl, text="📁 Папка с отчётами:", font=ctk.CTkFont(size=14)).pack(anchor="w")
        ctk.CTkEntry(ctrl, textvariable=self.folder_path).pack(fill="x", pady=5)
        ctk.CTkButton(ctrl, text="Выбрать папку", command=lambda: self.folder_path.set(filedialog.askdirectory())).pack(
            anchor="w", pady=5)
        btns = ctk.CTkFrame(f);
        btns.pack(fill="x", padx=20, pady=10)
        ctk.CTkButton(btns, text="📂 Тестовые данные", command=self.create_test_data).pack(side="left", padx=5)
        ctk.CTkButton(btns, text="🚀 Запуск", command=self.run_consolidation,
                      font=ctk.CTkFont(size=14, weight="bold")).pack(side="left", padx=5)
        st = ctk.CTkFrame(f);
        st.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(st, text="Статус:", font=ctk.CTkFont(size=12)).pack(anchor="w")
        ctk.CTkLabel(st, textvariable=self.status_text, wraplength=900, justify="left").pack(anchor="w", pady=5)
        self.progress_bar = ctk.CTkProgressBar(st);
        self.progress_bar.pack(fill="x", pady=10);
        self.progress_bar.set(0)
        tbl = ctk.CTkFrame(f);
        tbl.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tbl, show="headings", height=18)
        vsb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tbl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl.grid_rowconfigure(0, weight=1);
        tbl.grid_columnconfigure(0, weight=1)
        sb = ctk.CTkFrame(f);
        sb.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(sb, text="💾 Сохранить", command=self.save_report).pack(side="left", padx=5)
        ctk.CTkButton(sb, text="📊 Статистика", command=self.show_statistics).pack(side="left", padx=5)
        return f

    def create_settings_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        ctk.CTkLabel(f, text="⚙️ Настройки системы", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        box = ctk.CTkFrame(f, width=600);
        box.pack(pady=10, fill="both", expand=True)
        ctk.CTkLabel(box, text="Порог схожести группировки (%):").pack(anchor="w", padx=30, pady=(10, 5))
        ctk.CTkSlider(box, from_=60, to=100, variable=self.similarity_threshold, number_of_steps=40).pack(fill="x",
                                                                                                          padx=30,
                                                                                                          pady=5)
        ctk.CTkLabel(box, textvariable=self.similarity_threshold).pack()
        ctk.CTkLabel(box,
                     text="💡 Совет: 85% - оптимальное значение\n70-80% - более широкая группировка\n90-95% - строгая группировка",
                     text_color="gray", justify="left", wraplength=400).pack(pady=(5, 15))
        ctk.CTkLabel(box, text="Цветовая тема:").pack(anchor="w", padx=30, pady=(20, 5))
        self.theme_combo = ctk.CTkComboBox(box, values=["Синяя", "Зелёная", "Тёмно-синяя"], variable=self.theme_var,
                                           state="readonly", command=self.change_theme)
        self.theme_combo.pack(fill="x", padx=30, pady=5)
        ctk.CTkLabel(box, text="Режим отображения:").pack(anchor="w", padx=30, pady=(20, 5))
        self.appearance_menu = ctk.CTkOptionMenu(box, variable=self.appearance_var,
                                                 values=["Светлая", "Тёмная", "Системная"],
                                                 command=self.change_appearance_mode_event)
        self.appearance_menu.pack(fill="x", padx=30, pady=5)
        ctk.CTkLabel(box, text="📂 Категории расходов:", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w",
                                                                                                       padx=30,
                                                                                                       pady=(20, 5))
        self.cat_scroll = ctk.CTkScrollableFrame(box, height=160);
        self.cat_scroll.pack(fill="x", padx=30, pady=5)
        ctk.CTkButton(box, text="➕ Добавить категорию", command=lambda: self.add_cat_ui(), width=180).pack(pady=(5, 10))
        for k, v in self.expense_categories.items(): self.add_cat_ui(k, ", ".join(v))
        return f

    def create_history_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        ctk.CTkLabel(f, text="📜 История обработок", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        self.hist_tree = ttk.Treeview(f, show="headings", height=20)
        cols = ["ID", "Пользователь", "Папка", "Порог", "Статус", "Старт"]
        self.hist_tree["columns"] = cols
        for c in cols: self.hist_tree.heading(c, text=c); self.hist_tree.column(c, width=150)
        self.hist_tree.pack(fill="both", expand=True, padx=20, pady=10)
        ttk.Scrollbar(f, orient="vertical", command=self.hist_tree.yview).pack(side="right", fill="y")
        ctk.CTkButton(f, text="🔄 Обновить", command=self.load_history_ui).pack(pady=10)
        return f

    def create_admin_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        ctk.CTkLabel(f, text="🔑 Управление ключами доступа", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        ctrl = ctk.CTkFrame(f, width=600);
        ctrl.pack(pady=10)
        ctk.CTkLabel(ctrl, text="Тип ключа:").pack(anchor="w", padx=20)
        self.key_type_var = ctk.StringVar(value="user")
        ctk.CTkComboBox(ctrl, values=["user", "admin", "recovery"], variable=self.key_type_var, width=150).pack(
            side="left", padx=20, pady=10)
        ctk.CTkButton(ctrl, text="Сгенерировать", command=self.generate_key).pack(side="left", padx=10)
        self.generated_key_lbl = ctk.CTkLabel(f, text="", font=ctk.CTkFont(size=16, weight="bold"))
        self.generated_key_lbl.pack(pady=5)
        self.copy_btn = ctk.CTkButton(f, text="📋 Копировать", command=self.copy_key, state="disabled")
        self.copy_btn.pack(pady=5)
        ctk.CTkLabel(f, text="Активные ключи:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=40, pady=(20, 5))
        self.key_tree = ttk.Treeview(f, show="headings", height=10)
        self.key_tree["columns"] = ["ID", "Код", "Тип", "Создан"]
        for c in self.key_tree["columns"]: self.key_tree.heading(c, text=c); self.key_tree.column(c, width=150)
        self.key_tree.pack(fill="both", expand=True, padx=40, pady=10)
        ctk.CTkButton(f, text="🗑️ Удалить выбранный", command=self.delete_key).pack(pady=5)
        return f

    def create_users_frame(self):
        f = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        ctk.CTkLabel(f, text="👥 Управление учетными записями", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)
        self.user_tree = ttk.Treeview(f, show="headings", height=15)
        cols = ["ID", "Логин", "Роль", "Создан"]
        self.user_tree["columns"] = cols
        for c in cols: self.user_tree.heading(c, text=c); self.user_tree.column(c, width=150)
        self.user_tree.pack(fill="both", expand=True, padx=20, pady=10)
        ttk.Scrollbar(f, orient="vertical", command=self.user_tree.yview).pack(side="right", fill="y")
        pwd_frame = ctk.CTkFrame(f);
        pwd_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(pwd_frame, text="Новый пароль:").pack(side="left", padx=5)
        self.new_pwd_ent = ctk.CTkEntry(pwd_frame, placeholder_text="Введите новый пароль", show="*", width=200)
        self.new_pwd_ent.pack(side="left", padx=5)
        ctk.CTkButton(pwd_frame, text="🔒 Сменить пароль", command=self.change_user_password).pack(side="left", padx=5)
        ctk.CTkButton(f, text="🗑️ Удалить выбранного", command=self.delete_user, fg_color="red",
                      hover_color="darkred").pack(pady=5)
        ctk.CTkButton(f, text="🔄 Обновить список", command=self.load_users_ui).pack(pady=5)
        return f

    def rebuild_ui(self):
        current = "consolidator"
        frames_map = [("home", self.home_frame), ("consolidator", self.con_frame),
                      ("settings", self.set_frame), ("history", self.hist_frame)]
        if self.admin_btn: frames_map.extend([("admin", self.admin_frame), ("users", self.users_frame)])
        for name, frm in frames_map:
            if frm and frm.winfo_ismapped(): current = name; break

        for attr in ['nav', 'home_frame', 'con_frame', 'set_frame', 'hist_frame', 'admin_frame', 'users_frame']:
            if hasattr(self, attr) and getattr(self, attr) is not None:
                getattr(self, attr).destroy()

        # ✅ ОЧИСТКА СПИСКА УДАЛЁННЫХ ВИДЖЕТОВ
        self.cat_widgets = []

        self.create_navigation_frame()
        self.create_frames()
        self.select_frame_by_name(current)
        self.update()
        self.update_window_icon()

    def change_appearance_mode_event(self, choice):
        mode = APPEARANCE_MAP.get(choice, choice)
        ctk.set_appearance_mode(mode)
        self.rebuild_ui()

    def change_theme(self, choice):
        theme = THEME_MAP.get(choice, choice)
        ctk.set_default_color_theme(theme)
        self.rebuild_ui()

    def load_history_ui(self):
        for i in self.hist_tree.get_children(): self.hist_tree.delete(i)
        for r in HistoryCRUD.get_all(self.current_user["id"] if self.current_user["role"] != "admin" else None):
            self.hist_tree.insert("", "end", values=[r.get(k, "") for k in
                                                     ["id", "user_id", "folder_path", "threshold_used", "status",
                                                      "started_at"]])

    def load_admin_ui(self):
        if not hasattr(self, 'key_tree'): return
        for i in self.key_tree.get_children(): self.key_tree.delete(i)
        for r in KeyCRUD.get_unused_keys(self.current_user["id"]):
            self.key_tree.insert("", "end", values=[r["id"], r["key_code"], r["key_type"], r["created_at"]])

    def load_users_ui(self):
        if not hasattr(self, 'user_tree'): return
        for i in self.user_tree.get_children(): self.user_tree.delete(i)
        for r in UserCRUD.get_all():
            self.user_tree.insert("", "end", values=[r["id"], r["username"], r["role"], r["created_at"]])

    def generate_key(self):
        tp = self.key_type_var.get()
        code = f"KEY-{tp.upper()}-{secrets.token_hex(4).upper()}"
        KeyCRUD.create_key(code, tp, self.current_user["id"])
        self.generated_key_lbl.configure(text=code, text_color="lightgreen")
        self.copy_btn.configure(state="normal")
        self.load_admin_ui()

    def copy_key(self):
        self.clipboard_clear();
        self.clipboard_append(self.generated_key_lbl.cget("text"))

    def delete_key(self):
        sel = self.key_tree.selection()
        if not sel: return
        kid = self.key_tree.item(sel[0])["values"][0]
        if messagebox.askyesno("Удаление", f"Удалить ключ ID {kid}?"):
            from db_manager import get_connection
            conn = get_connection()
            conn.execute("DELETE FROM registration_keys WHERE id=?", (kid,));
            conn.commit();
            conn.close()
            self.load_admin_ui()

    def change_user_password(self):
        sel = self.user_tree.selection()
        if not sel: return messagebox.showwarning("Внимание", "Выберите пользователя в таблице!")
        uid = self.user_tree.item(sel[0])["values"][0]
        new_pwd = self.new_pwd_ent.get().strip()
        if not new_pwd: return messagebox.showerror("Ошибка", "Введите новый пароль!")
        if uid == self.current_user["id"]: return messagebox.showerror("Ошибка",
                                                                       "Нельзя менять пароль текущей сессии через эту форму!")
        if messagebox.askyesno("Подтверждение", f"Сменить пароль для пользователя ID {uid}?"):
            UserCRUD.update_password(uid, new_pwd)
            messagebox.showinfo("Успех", "Пароль успешно изменён!")
            self.new_pwd_ent.delete(0, "end")

    def delete_user(self):
        sel = self.user_tree.selection()
        if not sel: return messagebox.showwarning("Внимание", "Выберите пользователя в таблице!")
        uid = self.user_tree.item(sel[0])["values"][0]
        username = self.user_tree.item(sel[0])["values"][1]
        if uid == self.current_user["id"]: return messagebox.showerror("Ошибка", "Нельзя удалить себя!")
        if messagebox.askyesno("Подтверждение",
                               f"Удалить пользователя {username} (ID {uid})?\nВсе его данные будут потеряны."):
            UserCRUD.delete(uid)
            self.load_users_ui()
            messagebox.showinfo("Успех", f"Пользователь {username} удалён.")

    def add_cat_ui(self, name="", keywords=""):
        row = ctk.CTkFrame(self.cat_scroll);
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
        row.destroy();
        self.cat_widgets = [(e, t, r) for e, t, r in self.cat_widgets if r != row]

    def save_settings_to_db(self):
        ThresholdCRUD.create(self.current_user["id"], "Default", self.similarity_threshold.get(), is_default=True)

        for e, t, _ in self.cat_widgets:
            try:
                # Если виджет уже уничтожен (например, при смене темы), пропускаем его
                if not e.winfo_exists() or not t.winfo_exists():
                    continue

                n = e.get().strip().lower().replace(" ", "_")
                kws = [x.strip() for x in t.get("1.0", "end-1c").split(", ") if x.strip()]
                if n and kws:
                    TemplateCRUD.create(n, n, kws, self.current_user["id"])
            except Exception:
                # На случай, если виджет находится в процессе уничтожения при закрытии
                continue

    def create_test_data(self):
        try:
            target = filedialog.askdirectory()
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
            self.after(0, lambda: self.status_text.set("🚀 Запуск..."));
            self.after(0, lambda: self.progress_bar.set(0.1))
            hid = HistoryCRUD.create(self.current_user["id"], self.folder_path.get(), self.similarity_threshold.get())
            self.after(0, lambda: self.status_text.set("🔄 Обработка..."))

            def update(msg, pct):
                self.after(0, lambda: self.status_text.set(msg)); self.after(0,
                                                                             lambda: self.progress_bar.set(pct / 100))

            res, raw, err = smart_consolidate(self.folder_path.get(), self.similarity_threshold.get(), update)
            if err: raise Exception(err)
            self.result_df, self.raw_data = res, raw
            HistoryCRUD.update_status(hid, "completed", datetime.now().isoformat(), "результат.xlsx")
            self.after(0, self.display_results)
            self.after(0, lambda: self.status_text.set(f"✅ Готово! {len(raw)} → {len(res)} групп."))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)));
            self.after(0, lambda: self.progress_bar.set(0))

    def display_results(self):
        if self.result_df is None: return
        for i in self.tree.get_children(): self.tree.delete(i)
        cols = list(self.result_df.columns);
        self.tree["columns"] = cols
        widths = {'Консолидированная статья': 250, 'Общая сумма, руб': 150, 'Количество операций': 120,
                  'Затронутые отделы': 200, 'Диапазон дат': 180, 'Варианты названий': 300}
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths.get(c, 150), anchor="w" if c != 'Общая сумма, руб' else "e")
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
                    if self.raw_data is not None: self.raw_data.to_excel(w, sheet_name='Исходные', index=False)
                messagebox.showinfo("Успех", f"Сохранено в {fp}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def show_statistics(self):
        if self.result_df is None: return messagebox.showwarning("Внимание", "Нет данных!")
        win = ctk.CTkToplevel(self)
        win.title("📊 Статистика")
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