import pandas as pd
from fuzzywuzzy import fuzz
import os
import random
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import threading
from pathlib import Path

# Настройка внешнего вида customtkinter
ctk.set_appearance_mode("dark")  # Режимы: "dark", "light", "system"
ctk.set_default_color_theme("blue")  # Темы: "blue", "green", "dark-blue"

# ============ 1. СОЗДАЁМ РЕАЛИСТИЧНЫЕ ТЕСТОВЫЕ ДАННЫЕ ============

def create_realistic_files(folder_path='тестовые_отчёты_расширенные'):
    """Создаёт 5-10 файлов со сложной структурой"""
    
    # Реальные статьи расходов (разные варианты названий)
    expense_categories = {
        'канцтовары': ['канцтовары', 'канцелярия', 'офисные товары', 'канц. принадлежности', 'писчебумажные товары'],
        'бумага': ['бумага', 'бумага А4', 'бумага офисная', 'белая бумага', 'листы А4'],
        'картриджи': ['картриджи', 'тонер', 'печатающие устройства', 'расходники принтера', 'картриджи для принтера'],
        'хозтовары': ['хозтовары', 'хозяйственные товары', 'бытовая химия', 'средства уборки', 'салфетки', 'чистящие средства'],
        'чай_кофе': ['чай', 'кофе', 'напитки', 'чай/кофе', 'снэки', 'печенье к чаю'],
        'мебель': ['мебель офисная', 'столы', 'стулья', 'кресла', 'шкафы', 'оргтехника мебель'],
        'обучение': ['обучение', 'курсы', 'тренинги', 'повышение квалификации', 'вебинары', 'семинары'],
        'транспорт': ['такси', 'транспортные расходы', 'проезд', 'ГСМ', 'бензин', 'командировочные'],
        'связь': ['интернет', 'сотовая связь', 'телефония', 'мобильная связь', 'доступ в сеть'],
        'софт': ['софт', 'программное обеспечение', 'лицензии ПО', 'ПО', 'антивирус', 'облачные сервисы'],
        'обслуживание': ['обслуживание оргтехники', 'ремонт', 'сервисное обслуживание', 'заправка картриджей'],
        'канцелярия': ['канцелярия', 'ручки', 'карандаши', 'маркеры', 'стикеры', 'папки', 'скоросшиватели'],
    }
    
    # Создаём папку
    os.makedirs(folder_path, exist_ok=True)
    
    # Список отделов
    departments = ['Бухгалтерия', 'IT_отдел', 'Отдел_кадров', 'Юридический_отдел', 'Служба_безопасности']
    
    all_files = []
    
    for dept in departments:
        # Каждый отдел делает разное количество закупок
        num_transactions = random.randint(15, 30)
        
        data = []
        for _ in range(num_transactions):
            # Выбираем случайную категорию
            category = random.choice(list(expense_categories.keys()))
            # Берём случайный вариант названия из этой категории
            variants = expense_categories[category]
            name_variant = random.choice(variants)
            # Сумма от 500 до 50000
            amount = random.randint(500, 50000)
            # Дата в течение последних 30 дней
            date = datetime.now() - timedelta(days=random.randint(0, 30))
            
            # Добавляем дополнительную информацию
            data.append({
                '№ п/п': _ + 1,
                'Наименование (факт)': name_variant,
                'Категория (внутр)': category,
                'Сумма, руб': amount,
                'Дата операции': date.strftime('%d.%m.%Y'),
                'Подразделение': dept,
                'Статус': random.choice(['оплачено', 'ожидает', 'проведено', 'завершено']),
                'Примечание': random.choice(['', 'срочно', 'по договору', 'счёт №' + str(random.randint(1000, 9999))]),
                'Кол-во': random.randint(1, 10),
                'Ед.изм': random.choice(['шт', 'уп', 'компл', 'л', 'мес'])
            })
        
        df = pd.DataFrame(data)
        filename = f'{folder_path}/{dept}_отчёт.xlsx'
        df.to_excel(filename, index=False)
        all_files.append(filename)
    
    return all_files

# ============ 2. ПРОДВИНУТАЯ ГРУППИРОВКА ============

def smart_consolidate(folder_path, similarity_threshold=85, progress_callback=None):
    """
    Умная консолидация с несколькими алгоритмами
    """
    if progress_callback:
        progress_callback("🔄 Загружаем файлы из папки...", 0)
    
    # Собираем все Excel файлы
    all_data = []
    file_names = []
    
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    if not excel_files:
        return None, None, "❌ Нет Excel файлов в папке!"
    
    for idx, file in enumerate(excel_files):
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)
        all_data.append(df)
        file_names.append(file)
        if progress_callback:
            progress_callback(f"📂 Загружен {file}", int((idx + 1) / len(excel_files) * 30))
    
    # Объединяем все данные
    combined = pd.concat(all_data, ignore_index=True)
    
    if progress_callback:
        progress_callback(f"📊 Всего записей: {len(combined)}", 40)
    
    # Определяем колонку с названиями
    possible_name_columns = ['Наименование (факт)', 'статья', 'название', 'товар', 'услуга']
    name_col = None
    for col in combined.columns:
        if any(keyword in col.lower() for keyword in ['наимен', 'статья', 'назван', 'товар']):
            name_col = col
            break
    
    if name_col is None:
        for col in combined.columns:
            if combined[col].dtype == 'object':
                name_col = col
                break
    
    # Находим колонку с суммой
    amount_col = None
    for col in combined.columns:
        if 'сумма' in col.lower() or 'руб' in col.lower():
            amount_col = col
            break
    
    if amount_col is None:
        amount_col = combined.select_dtypes(include=['number']).columns[0]
    
    if progress_callback:
        progress_callback(f"🎯 Группировка с порогом {similarity_threshold}%...", 50)
    
    # Консолидация с нечёткой группировкой
    groups = {}
    processed = [False] * len(combined)
    
    for i in range(len(combined)):
        if processed[i]:
            continue
        
        current_name = str(combined[name_col].iloc[i])
        current_amount = combined[amount_col].iloc[i]
        current_date = combined['Дата операции'].iloc[i] if 'Дата операции' in combined.columns else None
        current_dept = combined['Подразделение'].iloc[i] if 'Подразделение' in combined.columns else None
        
        # Начинаем новую группу
        group_key = current_name
        groups[group_key] = {
            'каноническое_имя': current_name,
            'сумма_общая': current_amount,
            'количество_записей': 1,
            'оригинальные_названия': [current_name],
            'отделы': set([current_dept]) if current_dept else set(),
            'суммы_по_отделам': {},
            'диапазон_дат': [current_date] if current_date else []
        }
        
        if current_dept:
            groups[group_key]['суммы_по_отделам'][current_dept] = current_amount
        
        processed[i] = True
        
        # Ищем похожие строки
        for j in range(i + 1, len(combined)):
            if processed[j]:
                continue
            
            compare_name = str(combined[name_col].iloc[j])
            similarity = fuzz.ratio(current_name.lower(), compare_name.lower())
            
            if similarity >= similarity_threshold:
                amount_j = combined[amount_col].iloc[j]
                groups[group_key]['сумма_общая'] += amount_j
                groups[group_key]['количество_записей'] += 1
                groups[group_key]['оригинальные_названия'].append(compare_name)
                
                dept_j = combined['Подразделение'].iloc[j] if 'Подразделение' in combined.columns else None
                if dept_j:
                    groups[group_key]['отделы'].add(dept_j)
                    groups[group_key]['суммы_по_отделам'][dept_j] = groups[group_key]['суммы_по_отделам'].get(dept_j, 0) + amount_j
                
                date_j = combined['Дата операции'].iloc[j] if 'Дата операции' in combined.columns else None
                if date_j:
                    groups[group_key]['диапазон_дат'].append(date_j)
                
                processed[j] = True
        
        if progress_callback and i % 20 == 0:
            progress_callback(f"🔄 Обработано {i}/{len(combined)} записей", 50 + int(i / len(combined) * 30))
    
    if progress_callback:
        progress_callback("📊 Формируем результаты...", 90)
    
    # Формируем итоговый DataFrame
    results = []
    for group in groups.values():
        date_range = "нет данных"
        if group['диапазон_дат']:
            try:
                # Преобразуем в datetime, заменяя ошибки/пустоты на NaT
                dates = pd.to_datetime(group['диапазон_дат'], errors='coerce')
                # Оставляем только корректные даты
                valid_dates = dates.dropna()

                if len(valid_dates) > 0:
                    date_range = f"{valid_dates.min().strftime('%d.%m.%Y')} - {valid_dates.max().strftime('%d.%m.%Y')}"
                else:
                    date_range = "разные даты"
            except Exception:
                date_range = "разные даты"
        
        results.append({
            'Консолидированная статья': group['каноническое_имя'][:50] + ('...' if len(group['каноническое_имя']) > 50 else ''),
            'Общая сумма, руб': group['сумма_общая'],
            'Количество операций': group['количество_записей'],
            'Затронутые отделы': ', '.join(group['отделы']),
            'Макс. сумма по отделу': max(group['суммы_по_отделам'].values()) if group['суммы_по_отделам'] else 0,
            'Диапазон дат': date_range,
            'Варианты названий (пример)': ', '.join(group['оригинальные_названия'][:3]) + ('...' if len(group['оригинальные_названия']) > 3 else '')
        })
    
    result_df = pd.DataFrame(results)
    result_df = result_df.sort_values('Общая сумма, руб', ascending=False)
    
    return result_df, combined, None

# ============ 3. ГРАФИЧЕСКИЙ ИНТЕРФЕЙС ============

class ConsolidationApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Система консолидации отчётов")
        self.root.geometry("1200x700")
        
        # Переменные
        self.folder_path = ctk.StringVar()
        self.similarity_threshold = ctk.IntVar(value=85)
        self.status_text = ctk.StringVar(value="Готов к работе")
        
        # Создаём интерфейс
        self.setup_ui()
        
    def setup_ui(self):
        # Основной контейнер
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Левая панель с настройками
        left_panel = ctk.CTkFrame(main_frame, width=300)
        left_panel.pack(side="left", fill="y", padx=(0, 10))
        left_panel.pack_propagate(False)
        
        # Заголовок
        title_label = ctk.CTkLabel(left_panel, text="📊 Консолидатор отчётов", 
                                   font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=20)
        
        # Выбор папки
        folder_frame = ctk.CTkFrame(left_panel)
        folder_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(folder_frame, text="📁 Папка с отчётами:", 
                    font=ctk.CTkFont(size=14)).pack(anchor="w")
        
        folder_entry = ctk.CTkEntry(folder_frame, textvariable=self.folder_path)
        folder_entry.pack(fill="x", pady=(5, 5))
        
        ctk.CTkButton(folder_frame, text="Выбрать папку", 
                     command=self.select_folder).pack(pady=5)
        
        # Настройки группировки
        settings_frame = ctk.CTkFrame(left_panel)
        settings_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(settings_frame, text="⚙️ Настройки группировки:", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", pady=(0, 10))
        
        ctk.CTkLabel(settings_frame, text="Порог схожести (%):").pack(anchor="w")
        threshold_slider = ctk.CTkSlider(settings_frame, from_=60, to=100, 
                                        variable=self.similarity_threshold,
                                        number_of_steps=40)
        threshold_slider.pack(fill="x", pady=5)
        
        threshold_value = ctk.CTkLabel(settings_frame, textvariable=self.similarity_threshold)
        threshold_value.pack()
        
        ctk.CTkLabel(settings_frame, text="\n💡 Совет: 85% - оптимальное значение\n"
                                         "   70-80% - более широкая группировка\n"
                                         "   90-95% - строгая группировка", 
                    font=ctk.CTkFont(size=12), text_color="gray").pack(pady=10)
        
        # Кнопки действий
        actions_frame = ctk.CTkFrame(left_panel)
        actions_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkButton(actions_frame, text="📂 Создать тестовые данные", 
                     command=self.create_test_data,
                     fg_color="#2ecc71", hover_color="#27ae60").pack(fill="x", pady=5)
        
        ctk.CTkButton(actions_frame, text="🚀 Запустить консолидацию", 
                     command=self.run_consolidation,
                     fg_color="#3498db", hover_color="#2980b9",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", pady=10)
        
        # Статус
        status_frame = ctk.CTkFrame(left_panel)
        status_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(status_frame, text="Статус:", font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.status_label = ctk.CTkLabel(status_frame, textvariable=self.status_text, 
                                        wraplength=250, justify="left")
        self.status_label.pack(anchor="w", pady=5)
        
        self.progress_bar = ctk.CTkProgressBar(status_frame)
        self.progress_bar.pack(fill="x", pady=10)
        self.progress_bar.set(0)
        
        # Правая панель с результатами
        right_panel = ctk.CTkFrame(main_frame)
        right_panel.pack(side="right", fill="both", expand=True)
        
        # Таблица результатов
        ctk.CTkLabel(right_panel, text="📈 Результаты консолидации:", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(0, 10))
        
        # Фрейм для таблицы с прокруткой
        table_frame = ctk.CTkFrame(right_panel)
        table_frame.pack(fill="both", expand=True)
        
        # Создаём Treeview
        self.tree = ttk.Treeview(table_frame, show="headings", height=20)
        
        # Скроллбары
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Кнопка сохранения
        save_frame = ctk.CTkFrame(right_panel)
        save_frame.pack(fill="x", pady=(10, 0))
        
        ctk.CTkButton(save_frame, text="💾 Сохранить отчёт в Excel", 
                     command=self.save_report,
                     fg_color="#e67e22", hover_color="#d35400").pack(side="left", padx=5)
        
        ctk.CTkButton(save_frame, text="📊 Показать статистику", 
                     command=self.show_statistics,
                     fg_color="#9b59b6", hover_color="#8e44ad").pack(side="left", padx=5)
        
        # Хранилище результатов
        self.result_df = None
        self.raw_data = None
        
    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
            self.status_text.set(f"Выбрана папка: {folder}")

    def create_test_data(self):
        try:
            self.status_text.set("Создание тестовых данных...")
            self.progress_bar.set(0.3)

            # 📂 Открываем системный диалог выбора директории
            target_dir = filedialog.askdirectory(title="Выберите место для создания тестовых отчётов")
            if not target_dir:  # Пользователь нажал "Отмена"
                self.status_text.set("Выбор папки отменён")
                return

            # Формируем путь: создаём подпапку внутри выбранной директории
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
        """Запускает консолидацию в отдельном потоке"""
        if not self.folder_path.get():
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите папку с отчётами!")
            return
        
        # Запускаем в потоке, чтобы UI не зависал
        thread = threading.Thread(target=self._consolidate_thread)
        thread.start()
    
    def _consolidate_thread(self):
        """Поток для консолидации"""
        try:
            self.status_text.set("🚀 Запуск консолидации...")
            self.progress_bar.set(0.1)
            
            # Функция обновления прогресса
            def update_progress(message, percent):
                self.root.after(0, lambda: self.status_text.set(message))
                self.root.after(0, lambda: self.progress_bar.set(percent / 100))
            
            # Запускаем консолидацию
            result_df, raw_data, error = smart_consolidate(
                self.folder_path.get(), 
                self.similarity_threshold.get(),
                update_progress
            )
            
            if error:
                self.root.after(0, lambda: messagebox.showerror("Ошибка", error))
                self.root.after(0, lambda: self.status_text.set(f"❌ {error}"))
                return
            
            self.result_df = result_df
            self.raw_data = raw_data
            
            # Отображаем результаты в таблице
            self.root.after(0, self.display_results)
            
            self.root.after(0, lambda: self.progress_bar.set(1.0))
            self.root.after(0, lambda: self.status_text.set(
                f"✅ Готово! {len(raw_data)} записей → {len(result_df)} групп. "
                f"Общая сумма: {result_df['Общая сумма, руб'].sum():,.2f} руб."
            ))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Ошибка при консолидации:\n{str(e)}"))
            self.root.after(0, lambda: self.status_text.set(f"❌ Ошибка: {str(e)}"))
        finally:
            self.root.after(0, lambda: self.progress_bar.set(0))
    
    def display_results(self):
        """Отображает результаты в таблице"""
        if self.result_df is None or self.result_df.empty:
            return
        
        # Очищаем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Определяем колонки
        columns = list(self.result_df.columns)
        self.tree["columns"] = columns
        
        # Настраиваем колонки
        column_widths = {
            'Консолидированная статья': 250,
            'Общая сумма, руб': 150,
            'Количество операций': 120,
            'Затронутые отделы': 200,
            'Макс. сумма по отделу': 150,
            'Диапазон дат': 180,
            'Варианты названий (пример)': 300
        }
        
        for col in columns:
            width = column_widths.get(col, 150)
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor="w" if col != 'Общая сумма, руб' else "e")
        
        # Заполняем данными
        for _, row in self.result_df.head(100).iterrows():  # Показываем первые 100 строк
            values = [row[col] for col in columns]
            # Форматируем суммы
            for i, col in enumerate(columns):
                if 'сумма' in col.lower():
                    values[i] = f"{values[i]:,.2f}"
            self.tree.insert("", "end", values=values)
    
    def save_report(self):
        """Сохраняет отчёт в Excel"""
        if self.result_df is None:
            messagebox.showwarning("Предупреждение", "Нет данных для сохранения!\nСначала запустите консолидацию.")
            return
        
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="консолидированный_отчёт.xlsx"
            )
            
            if file_path:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    self.result_df.to_excel(writer, sheet_name='Консолидация', index=False)
                    if self.raw_data is not None:
                        self.raw_data.to_excel(writer, sheet_name='Исходные_данные', index=False)
                
                self.status_text.set(f"💾 Отчёт сохранён: {file_path}")
                messagebox.showinfo("Успех", f"Отчёт успешно сохранён!\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить отчёт:\n{str(e)}")
    
    def show_statistics(self):
        """Показывает статистику в отдельном окне"""
        if self.result_df is None:
            messagebox.showwarning("Предупреждение", "Нет данных для анализа!\nСначала запустите консолидацию.")
            return
        
        # Создаём окно статистики
        stats_window = ctk.CTkToplevel(self.root)
        stats_window.title("Статистика консолидации")
        stats_window.geometry("500x400")
        stats_window.grab_focus()
        
        # Статистика
        stats_text = f"""
📊 СТАТИСТИКА КОНСОЛИДАЦИИ
{'='*40}

📥 Исходных записей: {len(self.raw_data) if self.raw_data is not None else 0}
📤 После консолидации: {len(self.result_df)}
✅ Сокращение: {len(self.raw_data) - len(self.result_df) if self.raw_data is not None else 0} ручных операций

💰 Общая сумма расходов: {self.result_df['Общая сумма, руб'].sum():,.2f} руб.
📈 Средняя сумма по статье: {self.result_df['Общая сумма, руб'].mean():,.2f} руб.
🏆 Максимальная сумма по статье: {self.result_df['Общая сумма, руб'].max():,.2f} руб.

📋 ТОП-5 СТАТЕЙ РАСХОДОВ:
{'='*40}
"""
        
        for i, row in self.result_df.head(5).iterrows():
            stats_text += f"\n{i+1}. {row['Консолидированная статья'][:40]}\n"
            stats_text += f"   Сумма: {row['Общая сумма, руб']:,.2f} руб. | {row['Количество операций']} оп."
        
        # Отображаем статистику
        text_widget = ctk.CTkTextbox(stats_window, font=ctk.CTkFont(size=12))
        text_widget.pack(fill="both", expand=True, padx=20, pady=20)
        text_widget.insert("1.0", stats_text)
        text_widget.configure(state="disabled")
        
        # Кнопка закрытия
        ctk.CTkButton(stats_window, text="Закрыть", command=stats_window.destroy).pack(pady=(0, 20))
    
    def run(self):
        self.root.mainloop()

# ============ 4. ЗАПУСК ПРИЛОЖЕНИЯ ============

if __name__ == "__main__":
    app = ConsolidationApp()
    app.run()