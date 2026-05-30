import pandas as pd
from fuzzywuzzy import fuzz
import os
import random
import json
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import threading

# ==================== 0. РАБОТА С НАСТРОЙКАМИ ====================
CONFIG_FILE = "config.json"


def load_config():
    defaults = {"threshold": 85, "appearance": "dark", "theme": "blue"}
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


# ==================== 1. ТЕСТОВЫЕ ДАННЫЕ ====================
def create_realistic_files(folder_path='тестовые_отчёты_расширенные'):
    """Создаёт 5-10 файлов со сложной структурой"""
    expense_categories = {
        'канцтовары': ['канцтовары', 'канцелярия', 'офисные товары', 'канц. принадлежности', 'писчебумажные товары'],
        'бумага': ['бумага', 'бумага А4', 'бумага офисная', 'белая бумага', 'листы А4'],
        'картриджи': ['картриджи', 'тонер', 'печатающие устройства', 'расходники принтера', 'картриджи для принтера'],
        'хозтовары': ['хозтовары', 'хозяйственные товары', 'бытовая химия', 'средства уборки', 'салфетки',
                      'чистящие средства'],
        'чай_кофе': ['чай', 'кофе', 'напитки', 'чай/кофе', 'снэки', 'печенье к чаю'],
        'мебель': ['мебель офисная', 'столы', 'стулья', 'кресла', 'шкафы', 'оргтехника мебель'],
        'обучение': ['обучение', 'курсы', 'тренинги', 'повышение квалификации', 'вебинары', 'семинары'],
        'транспорт': ['такси', 'транспортные расходы', 'проезд', 'ГСМ', 'бензин', 'командировочные'],
        'связь': ['интернет', 'сотовая связь', 'телефония', 'мобильная связь', 'доступ в сеть'],
        'софт': ['софт', 'программное обеспечение', 'лицензии ПО', 'ПО', 'антивирус', 'облачные сервисы'],
        'обслуживание': ['обслуживание оргтехники', 'ремонт', 'сервисное обслуживание', 'заправка картриджей'],
        'канцелярия': ['канцелярия', 'ручки', 'карандаши', 'маркеры', 'стикеры', 'папки', 'скоросшиватели'],
    }

    os.makedirs(folder_path, exist_ok=True)
    departments = ['Бухгалтерия', 'IT_отдел', 'Отдел_кадров', 'Юридический_отдел', 'Служба_безопасности']
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
                'Статус': random.choice(['оплачено', 'ожидает', 'проведено', 'завершено']),
                'Примечание': random.choice(['', 'срочно', 'по договору', f'счёт №{random.randint(1000, 9999)}']),
                'Кол-во': random.randint(1, 10),
                'Ед.изм': random.choice(['шт', 'уп', 'компл', 'л', 'мес'])
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

    # Поиск колонок
    name_col = amount_col = None
    for col in combined.columns:
        cl = col.lower()
        if not name_col and any(k in cl for k in ['наимен', 'статья', 'назван', 'товар']): name_col = col
        if not amount_col and ('сумма' in cl or 'руб' in cl): amount_col = col
    if not name_col:
        name_col = combined.select_dtypes(include=['object']).columns[0]
    if not amount_col:
        amount_col = combined.select_dtypes(include=['number']).columns[0]

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
class ConsolidationApp:
    def __init__(self):
        self.config = load_config()
        ctk.set_appearance_mode(self.config["appearance"])
        ctk.set_default_color_theme(self.config["theme"])

        self.root = ctk.CTk()
        self.root.title("Система консолидации отчётов")
        self.root.geometry("1200x700")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)  # Сохранение при выходе

        self.result_df = self.raw_data = None
        self.folder_path = ctk.StringVar()
        self.similarity_threshold = ctk.IntVar(value=self.config["threshold"])
        self.status_text = ctk.StringVar(value="Готов к работе")

        self.appearance_var = ctk.StringVar(value=self.config["appearance"])
        self.theme_var = ctk.StringVar(value=self.config["theme"])

        self.setup_main_layout()
        self.create_frames()
        self.show_frame("consolidator")  # По умолчанию открываем рабочий раздел

    def setup_main_layout(self):
        self.main_container = ctk.CTkFrame(self.root)
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)

        self.sidebar = ctk.CTkFrame(self.main_container, width=220)
        self.sidebar.pack(side="left", fill="y", padx=(0, 10))
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(self.sidebar, text="📑 МЕНЮ", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(30, 20))
        ctk.CTkButton(self.sidebar, text="🏠 Главная", command=lambda: self.show_frame("home")).pack(fill="x", padx=15,
                                                                                                    pady=8)
        ctk.CTkButton(self.sidebar, text="📈 Консолидатор отчётов",
                      command=lambda: self.show_frame("consolidator")).pack(fill="x", padx=15, pady=8)
        ctk.CTkButton(self.sidebar, text="⚙️ Настройки", command=lambda: self.show_frame("settings")).pack(fill="x",
                                                                                                           padx=15,
                                                                                                           pady=8)

        self.content_area = ctk.CTkFrame(self.main_container)
        self.content_area.pack(side="right", fill="both", expand=True)

    def show_frame(self, frame_name):
        if self.current_frame: self.current_frame.pack_forget()
        self.frames[frame_name].pack(fill="both", expand=True)
        self.current_frame = self.frames[frame_name]

    def create_frames(self):
        self.frames = {}
        self.frames["home"] = self.create_home_frame()
        self.frames["consolidator"] = self.create_consolidator_frame()
        self.frames["settings"] = self.create_settings_frame()
        self.current_frame = self.frames["consolidator"]

    def create_home_frame(self):
        frame = ctk.CTkFrame(self.content_area)
        ctk.CTkLabel(frame, text="📖 О программе", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(40, 10))

        info = ("Данная программа разработана для автоматической консолидации финансовых и операционных отчётов.\n\n"
                "🔹 Основные возможности:\n"
                "• Загрузка множества Excel-файлов из выбранной папки.\n"
                "• Умная группировка строк с использованием нечёткого сравнения (fuzzy matching).\n"
                "• Автоматический подсчёт сумм, объединение диапазонов дат и статистика по отделам.\n"
                "• Экспорт результатов в единый Excel-отчёт.\n\n"
                "🔹 Как работать:\n"
                "1. Перейдите во вкладку «Консолидатор отчётов».\n"
                "2. Выберите папку с исходными файлами или создайте тестовые данные.\n"
                "3. При необходимости настройте порог схожести во вкладке «Настройки».\n"
                "4. Нажмите «Запустить консолидацию» и дождитесь результата.\n"
                "5. Сохраните итоговый отчёт или просмотрите статистику.")

        ctk.CTkLabel(frame, text=info, justify="left", wraplength=750, font=ctk.CTkFont(size=14)).pack(pady=30, padx=40)
        ctk.CTkLabel(frame, text="👤 Автор: Уливанова София", font=ctk.CTkFont(size=16, weight="bold"),
                     text_color="gray").pack(side="bottom", pady=40)
        return frame

    def create_consolidator_frame(self):
        frame = ctk.CTkFrame(self.content_area)

        # Управление папкой
        ctrl = ctk.CTkFrame(frame)
        ctrl.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(ctrl, text="📁 Папка с отчётами:", font=ctk.CTkFont(size=14)).pack(anchor="w")
        ctk.CTkEntry(ctrl, textvariable=self.folder_path).pack(fill="x", pady=5)
        ctk.CTkButton(ctrl, text="Выбрать папку", command=self.select_folder).pack(anchor="w", pady=5)

        # Кнопки действий
        btns = ctk.CTkFrame(frame)
        btns.pack(fill="x", padx=20, pady=10)
        ctk.CTkButton(btns, text="📂 Создать тестовые данные", command=self.create_test_data, fg_color="#2ecc71",
                      hover_color="#27ae60").pack(side="left", padx=5)
        ctk.CTkButton(btns, text="🚀 Запустить консолидацию", command=self.run_consolidation, fg_color="#3498db",
                      hover_color="#2980b9", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left", padx=5)

        # Статус и прогресс
        st = ctk.CTkFrame(frame)
        st.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(st, text="Статус:", font=ctk.CTkFont(size=12)).pack(anchor="w")
        ctk.CTkLabel(st, textvariable=self.status_text, wraplength=650, justify="left").pack(anchor="w", pady=5)
        self.progress_bar = ctk.CTkProgressBar(st)
        self.progress_bar.pack(fill="x", pady=10)
        self.progress_bar.set(0)

        # Таблица
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

        # Нижние кнопки
        sb = ctk.CTkFrame(frame)
        sb.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(sb, text="💾 Сохранить отчёт", command=self.save_report, fg_color="#e67e22",
                      hover_color="#d35400").pack(side="left", padx=5)
        ctk.CTkButton(sb, text="📊 Показать статистику", command=self.show_statistics, fg_color="#9b59b6",
                      hover_color="#8e44ad").pack(side="left", padx=5)
        return frame

    def create_settings_frame(self):
        frame = ctk.CTkFrame(self.content_area)
        ctk.CTkLabel(frame, text="⚙️ Настройки системы", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=30)

        box = ctk.CTkFrame(frame, width=450, height=350)
        box.pack(pady=20)
        box.pack_propagate(False)

        ctk.CTkLabel(box, text="Порог схожести группировки (%):").pack(anchor="w", padx=30, pady=(20, 5))
        ctk.CTkSlider(box, from_=60, to=100, variable=self.similarity_threshold, number_of_steps=40).pack(fill="x",
                                                                                                          padx=30,
                                                                                                          pady=5)
        ctk.CTkLabel(box, textvariable=self.similarity_threshold).pack()

        ctk.CTkLabel(box, text="Режим оформления:").pack(anchor="w", padx=30, pady=(30, 5))
        ctk.CTkComboBox(box, values=["dark", "light", "system"], variable=self.appearance_var,
                        command=lambda v: ctk.set_appearance_mode(v)).pack(fill="x", padx=30, pady=5)

        ctk.CTkLabel(box, text="Цветовая тема:").pack(anchor="w", padx=30, pady=(20, 5))
        ctk.CTkComboBox(box, values=["blue", "green", "dark-blue"], variable=self.theme_var,
                        command=lambda v: ctk.set_default_color_theme(v)).pack(fill="x", padx=30, pady=5)

        ctk.CTkLabel(box, text="💡 Настройки сохраняются автоматически при закрытии программы.", text_color="gray",
                     wraplength=380).pack(pady=40, padx=30)
        return frame

    # ==================== ЛОГИКА ====================
    def on_closing(self):
        self.config["threshold"] = self.similarity_threshold.get()
        self.config["appearance"] = self.appearance_var.get()
        self.config["theme"] = self.theme_var.get()
        save_config(self.config)
        self.root.destroy()

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
            self.status_text.set(f"Выбрана папка: {folder}")

    def create_test_data(self):
        try:
            self.status_text.set("Создание тестовых данных...")
            self.progress_bar.set(0.3)

            target_dir = filedialog.askdirectory(title="Выберите место для создания тестовых отчётов")
            if not target_dir:
                self.status_text.set("Выбор папки отменён")
                return

            folder_path = os.path.join(target_dir, "тестовые_отчёты_расширенные")
            files = create_realistic_files(folder_path)

            self.progress_bar.set(1.0)
            self.status_text.set(f"✅ Создано {len(files)} тестовых файлов в папке '{folder_path}'")
            messagebox.showinfo("Успех", f"Создано {len(files)} тестовых отчётов!\nПуть: {folder_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать тестовые данные:\n{str(e)}")
            self.status_text.set(f"❌ Ошибка: {str(e)}")
        finally:
            self.progress_bar.set(0)

    def run_consolidation(self):
        if not self.folder_path.get():
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите папку с отчётами!")
            return
        thread = threading.Thread(target=self._consolidate_thread)
        thread.start()

    def _consolidate_thread(self):
        try:
            self.root.after(0, lambda: self.status_text.set("🚀 Запуск консолидации..."))
            self.root.after(0, lambda: self.progress_bar.set(0.1))

            def update(msg, pct):
                self.root.after(0, lambda: self.status_text.set(msg))
                self.root.after(0, lambda: self.progress_bar.set(pct / 100))

            res, raw, err = smart_consolidate(self.folder_path.get(), self.similarity_threshold.get(), update)
            if err:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", err))
                return

            self.result_df, self.raw_data = res, raw
            self.root.after(0, self.display_results)
            self.root.after(0, lambda: self.progress_bar.set(1.0))
            self.root.after(0, lambda: self.status_text.set(
                f"✅ Готово! {len(raw)} записей → {len(res)} групп. "
                f"Общая сумма: {res['Общая сумма, руб'].sum():,.2f} руб."
            ))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Ошибка при консолидации:\n{str(e)}"))
        finally:
            self.root.after(0, lambda: self.progress_bar.set(0))

    def display_results(self):
        if self.result_df is None or self.result_df.empty: return
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
        if self.result_df is None:
            messagebox.showwarning("Предупреждение", "Нет данных для сохранения!\nСначала запустите консолидацию.")
            return
        try:
            fp = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                              initialfile="консолидированный_отчёт.xlsx")
            if fp:
                with pd.ExcelWriter(fp, engine='openpyxl') as writer:
                    self.result_df.to_excel(writer, sheet_name='Консолидация', index=False)
                    if self.raw_data is not None: self.raw_data.to_excel(writer, sheet_name='Исходные_данные',
                                                                         index=False)
                self.status_text.set(f"💾 Отчёт сохранён: {fp}")
                messagebox.showinfo("Успех", f"Отчёт успешно сохранён!\n{fp}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить отчёт:\n{str(e)}")

    def show_statistics(self):
        if self.result_df is None:
            messagebox.showwarning("Предупреждение", "Нет данных для анализа!\nСначала запустите консолидацию.")
            return

        win = ctk.CTkToplevel(self.root)
        win.title("Статистика консолидации")
        win.geometry("500x400")
        win.grab_focus()

        raw_len = len(self.raw_data) if self.raw_data is not None else 0
        txt = (f"📊 СТАТИСТИКА КОНСОЛИДАЦИИ\n{'=' * 40}\n"
               f"📥 Исходных записей: {raw_len}\n📤 После консолидации: {len(self.result_df)}\n"
               f"✅ Сокращение: {raw_len - len(self.result_df)} ручных операций\n"
               f"💰 Общая сумма расходов: {self.result_df['Общая сумма, руб'].sum():,.2f} руб.\n"
               f"📈 Средняя сумма по статье: {self.result_df['Общая сумма, руб'].mean():,.2f} руб.\n"
               f"🏆 Максимальная сумма по статье: {self.result_df['Общая сумма, руб'].max():,.2f} руб.\n\n"
               f"📋 ТОП-5 СТАТЕЙ РАСХОДОВ:\n{'=' * 40}\n")
        for i, row in self.result_df.head(5).iterrows():
            txt += f"\n{i + 1}. {row['Консолидированная статья'][:40]}\n   Сумма: {row['Общая сумма, руб']:,.2f} руб. | {row['Количество операций']} оп."

        t = ctk.CTkTextbox(win, font=ctk.CTkFont(size=12))
        t.pack(fill="both", expand=True, padx=20, pady=20)
        t.insert("1.0", txt)
        t.configure(state="disabled")
        ctk.CTkButton(win, text="Закрыть", command=win.destroy).pack(pady=(0, 20))

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = ConsolidationApp()
    app.run()