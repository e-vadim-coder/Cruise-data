# -*- coding: utf-8 -*-
# main.py
import tkinter as tk  # (0)
from tkinter import filedialog, font, messagebox, ttk  # (0)
import pandas as pd  # (0)

from helpers import (  # (0)
    FIELDS_CONFIG,  # (4)
    DICTIONARIES,
    TABS_CONFIG,
    TABLE_HEADINGS_SHORT,
    create_backup,
    clean_duplicate_backups,
    get_file_hash,
    delete_row_from_df,  # (4)
    export_to_txt,  # (4)
    is_exact_match,  # (4)
    load_data,  # (4)
    reset_form_fields,  # (4)
    save_data,  # (4)
    validate_type,  # (4)
    get_card_context  # (0)
)  # (0)

FILE_NAME = "voyage_data.xlsx"  # (0)

class VoyageAppTabs:  # (0)

    def __init__(self, root):  # (4)
        self.root = root  # (8)
        self.root.title(" - - - - - Учет рейсовых данных - - - - -  ")  # (8)
        self.root.geometry("1400x980")  # (8)

        self.df = load_data(FILE_NAME)  # (8)
        self.selected_index = None  # (8)
        self.template_index = None  # (8)
        self.inputs = {}  # (8)

        self.btn_colors = {  # (8)
            "search"  : "#2980b9",  # (12)
            "add"     : "#27ae60",  # (12)
            "template": "#8e44ad",  # (12)
            "edit"    : "#d35400",  # (12)
            "delete"  : "#c0392b",  # (12)
            "clear"   : "#7f8c8d",  # (12)
            "export"  : "#16a085",  # (12)
        }  # (8)

        self.create_widgets()  # (8)
        self.refresh_table()  # (8)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)  # (8)

    def create_widgets(self):  # (4)
        # --- ВКЛАДКИ ДЛЯ ВВОДА ДАННЫХ ---  # (8)
        self.notebook = ttk.Notebook(self.root)  # (8)
        self.notebook.pack(fill="x", padx=10, pady=5)  # (8)

        for tab_name, fields in TABS_CONFIG.items():  # (8)
            tab_frame = tk.Frame(self.notebook, padx=15, pady=10)  # (12)
            self.notebook.add(tab_frame, text=f"  {tab_name}  ")  # (12)

            row, col = 0, 0  # (12)
            for field in fields:  # (12)
                lbl = tk.Label(  # (16)
                    tab_frame, text=field + ":", font=("Arial", 10)  # (20)
                )  # (16)
                lbl.grid(row=row, column=col, sticky="w", padx=5, pady=5)  # (16)

                if field in DICTIONARIES:  # (16)
                    widget = ttk.Combobox(  # (20)
                        tab_frame, values=DICTIONARIES[field], width=25  # (24)
                    )  # (20)
                else:  # (16)
                    widget = tk.Entry(tab_frame, width=28)  # (20)

                widget.grid(row=row, column=col + 1, padx=10, pady=5)  # (16)
                self.inputs[field] = widget  # (16)

                col += 2  # (16)
                if col >= 4:  # (16)
                    col = 0  # (20)
                    row += 1  # (20)

        # --- ОБНОВЛЕННАЯ ПАНЕЛЬ ЦВЕТНЫХ КНОПОК ---  # (8)
        btn_frame = tk.Frame(self.root, pady=10)  # (8)
        btn_frame.pack(fill="x", padx=10)  # (8)

        tk.Button(  # (8)
            btn_frame,  # (12)
            text="🔍 Поиск",  # (12)
            bg=self.btn_colors["search"],  # (12)
            fg="white",  # (12)
            font=("Arial", 10, "bold"),  # (12)
            padx=10,  # (12)
            command=self.search_data,  # (12)
        ).pack(side="left", padx=3)  # (8)

        tk.Button(  # (8)
            btn_frame,  # (12)
            text="➕ Создать новую",  # (12)
            bg=self.btn_colors["add"],  # (12)
            fg="white",  # (12)
            font=("Arial", 10, "bold"),  # (12)
            padx=10,  # (12)
            command=self.add_data,  # (12)
        ).pack(side="left", padx=3)  # (8)

        tk.Button(  # (8)
            btn_frame,  # (12)
            text="📋 Из шаблона",  # (12)
            bg=self.btn_colors["template"],  # (12)
            fg="white",  # (12)
            font=("Arial", 10, "bold"),  # (12)
            padx=10,  # (12)
            command=self.prepare_template,  # (12)
        ).pack(side="left", padx=3)  # (8)

        tk.Button(  # (8)
            btn_frame,  # (12)
            text="💾 Сохранить изменения",  # (12)
            bg=self.btn_colors["edit"],  # (12)
            fg="white",  # (12)
            font=("Arial", 10, "bold"),  # (12)
            padx=10,  # (12)
            command=self.update_data,  # (12)
        ).pack(side="left", padx=3)  # (8)

        tk.Button(  # (8)
            btn_frame,  # (12)
            text="❌ Удалить запись",  # (12)
            bg=self.btn_colors["delete"],  # (12)
            fg="white",  # (12)
            font=("Arial", 10, "bold"),  # (12)
            padx=10,  # (12)
            command=self.delete_data,  # (12)
        ).pack(side="left", padx=3)  # (8)

        # Кнопка перехода в карточный режим «Детальный просмотр»
        self.btn_detail_view = tk.Button(  # (8)
            btn_frame,  # (12)
            text="Детальный просмотр",  # (12)
            bg="#d2b48c",  # (12) Цвет кнопок просмотра (светло-коричневый) в вашем стиле
            font=("Arial", 10),  # (12)
            command=self.open_detail_view  # (12)
        )  # (8)
        self.btn_detail_view.pack(side=tk.LEFT, padx=2)  # (8)


        tk.Button(  # (8)
            btn_frame,  # (12)
            text="🧹 Сбросить форму",  # (12)
            bg=self.btn_colors["clear"],  # (12)
            fg="white",  # (12)
            font=("Arial", 10, "bold"),  # (12)
            padx=10,  # (12)
            command=self.clear_form,  # (12)
        ).pack(side="left", padx=3)  # (8)

        # в самом конце ряда:
        tk.Button(
            btn_frame,
            text="📄 Экспорт отчета",
            bg=self.btn_colors["export"],
            fg="white",
            font=("Arial", 10, "bold"),
            command=self.export_report
        ).pack(side="left", padx=5)  # (8)

        # --- ТАБЛИЦА ПРОСМОТРА ---  # (8)
        table_frame = tk.LabelFrame(  # (8)
            self.root,  # (12)
            text=" Просмотр базы данных рейсов ",  # (12)
            font=("Arial", 10),  # (12)
        )  # (8)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)  # (8)

        self.tree = ttk.Treeview(  # (8)
            table_frame,  # (12)
            columns=list(FIELDS_CONFIG.keys()),  # (12)
            show="headings",  # (12)
        )  # (8)

        scroll_x = tk.Scrollbar(  # (8)
            table_frame, orient="horizontal", command=self.tree.xview  # (12)
        )  # (8)
        scroll_y = tk.Scrollbar(  # (8)
            table_frame, orient="vertical", command=self.tree.yview  # (12)
        )  # (8)
        self.tree.configure(  # (8)
            xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set  # (12)
        )  # (8)

        scroll_x.pack(side="bottom", fill="x")  # (8)
        scroll_y.pack(side="right", fill="y")  # (8)
        self.tree.pack(fill="both", expand=True)  # (8)

        for field in FIELDS_CONFIG.keys():  # (8)
            # Берем короткое имя из словаря, если его там нет — оставляем длинное
            short_text = TABLE_HEADINGS_SHORT.get(field, field)
            self.tree.heading(field, text=short_text)
            self.tree.column(field, width=125, anchor="center")  # (12)

    def get_form_values(self) -> dict:  # (4)
        return {  # (8)
            field: widget.get().strip()  # (12)
            for field, widget in self.inputs.items()  # (12)
        }  # (8)

    def validate_form_data(self, data: dict) -> bool:  # (4)
        if not isinstance(data, dict):  # (8)
            raise TypeError("Параметр data должен быть словарем.")  # (12)

        for field, value in data.items():  # (8)
            expected_type = FIELDS_CONFIG[field]  # (12)
            if not validate_type(value, expected_type):  # (12)
                messagebox.showerror(  # (16)
                    "Ошибка валидации",  # (20)
                    f"Поле '{field}' принимает только целые числа!",  # (20)
                )  # (16)
                return False  # (16)
        return True  # (8)

    def refresh_table(self, dataframe: pd.DataFrame = None):  # (4)
        for item in self.tree.get_children():  # (8)
            self.tree.delete(item)  # (12)

        display_df = self.df if dataframe is None else dataframe  # (8)

        # 1. Заполняем таблицу данными
        for idx, row in display_df.iterrows():  # (8)
            values = ["" if pd.isna(val) else val for val in row]  # (12)
            self.tree.insert("", "end", iid=str(idx), values=values)  # (12)

        # 2. Настраиваем автоподбор ширины колонок
        # Берем стандартный шрифт интерфейса для точного расчета пикселей
        tk_font = font.Font(font="TkDefaultFont")  # (8)

        for field in FIELDS_CONFIG.keys():  # (8)
            # Измеряем длину заголовка колонки (+ запас на стрелочку сортировки)
            # max_len = tk_font.measure(str(field)) + 25  # (12)
            short_text = TABLE_HEADINGS_SHORT.get(field, field)
            max_len = tk_font.measure(str(short_text)) + 25

            # Если в таблице есть данные, ищем самую длинную строку в текущей колонке
            if not display_df.empty:  # (12)
                # Переводим все значения колонки в текст и находим макс. длину в пикселях
                col_lens = [  # (16)
                    tk_font.measure(str("" if pd.isna(val) else val))  # (20)
                    for val in display_df[field]  # (20)
                ]  # (16)
                max_len = max(max_len, max(col_lens) + 15)  # (16)

            # Ограничиваем минимальную и максимальную ширину для красоты (от 70 до 350 пикселей)
            final_width = min(max(max_len, 50), 113)  # (12)

            # Применяем вычисленную ширину к колонке
            self.tree.column(field, width=final_width, anchor="center", stretch=False, minwidth=final_width)  # (12)

    def on_row_select(self, event):  # (4)
        selected_items = self.tree.selection()  # (8)
        if not selected_items:  # (8)
            return  # (12)

        # Явно берем индекс первой выбранной строки в Treeview
        tree_id = selected_items[0]  # (8)
        self.selected_index = int(tree_id)  # (8)
        self.template_index = None  # Сбрасываем шаблон при обычном редактировании  # (8)

        # Получаем данные строки из нашего DataFrame
        row_data = self.df.loc[self.selected_index]  # (8)

        reset_form_fields(self.inputs)  # (8)

        for field, widget in self.inputs.items():  # (8)
            val = row_data[field]  # (12)
            if pd.isna(val):  # (12)
                val = ""  # (16)
            if isinstance(widget, ttk.Combobox):  # (12)
                widget.set(str(val))  # (16)
            else:  # (12)
                widget.insert(0, str(val))  # (16)

    def prepare_template(self):  # (4)
        selected_items = self.tree.selection()  # (8)
        if not selected_items:  # (8)
            messagebox.showwarning("Внимание", "Пожалуйста, выберите запись-шаблон в таблице!")  # (12)
            return  # (12)

        # Безопасно извлекаем ID выделенной строки из кортежа
        if isinstance(selected_items, (tuple, list)):  # (8)
            selected_id = selected_items[0]  # (12)
        else:  # (8)
            selected_id = selected_items  # (12)

        # Выводим диалоговое окно с тремя вариантами действий  # (8)
        choice = messagebox.askyesnocancel(  # (8)
            "Создание по шаблону",  # (12)
            "Вы хотите отредактировать новую запись после создания?\n\n"  # (12)
            "«Да» — создать и сразу открыть карточку «Детальный просмотр» для исправления\n"  # (12)
            "«Нет» — просто создать копию (изменится только № п/п)\n"  # (12)
            "«Отмена» — выйти и ничего не менять в таблице"  # (12)
        )  # (8)

        # Сценарий 1: Пользователь нажал «Отмена» (Выходим без изменений)  # (8)
        if choice is None:  # (8)
            return  # (12)

        # Сценарий 2: Пользователь нажал «Да» (Создать и открыть детальное окно)
        elif choice is True:  # (8)
            # Извлекаем строку-шаблон в словарь  # (12)
            template_row = self.df.loc[int(selected_id)].to_dict()  # (12)

            # Вычисляем уникальный № п/п (исторический максимум во всей базе + 1)  # (12)
            if "№ п/п" in template_row:  # (12)
                max_pp = self.df["№ п/п"].max()  # (16)
                template_row["№ п/п"] = int(max_pp) + 1 if pd.notna(max_pp) else 1  # (16)

            # Добавляем новую строку в Pandas DataFrame  # (12)
            new_index = len(self.df)  # (12)
            self.df.loc[new_index] = template_row  # (12)

            # Вызываем ваше родное обновление таблицы
            self.refresh_table()  # (12)

            # ИСПРАВЛЕНО: Запускаем окно локально из этого же файла main.py
            DetailViewWindow(self, new_index)  # (12)
            
        # Сценарий 3: Пользователь нажал «Нет» (Просто копия на том же диске)  # (8)
        elif choice is False:  # (8)
            template_row = self.df.loc[int(selected_id)].to_dict()  # (12)

            # Пересчитываем строго системный № п/п, диски копируются 1 в 1  # (12)
            if "№ п/п" in template_row:  # (12)
                max_pp = self.df["№ п/п"].max()  # (16)
                template_row["№ п/п"] = int(max_pp) + 1 if pd.notna(max_pp) else 1  # (16)

            new_index = len(self.df)  # (12)
            self.df.loc[new_index] = template_row  # (12)

            # Обновляем таблицу вашим методом
            self.refresh_table()  # (12)

            messagebox.showinfo("Успех", f"Запись успешно скопирована! Новый № п/п: {template_row['№ п/п']}")  # (12)

    def clear_form(self):  # (4)
        reset_form_fields(self.inputs)  # (8)
        self.selected_index = None  # (8)
        self.template_index = None  # (8)
        self.refresh_table()  # (8)
    def add_data(self):  # (4)
        data = self.get_form_values()  # (8)
        if not self.validate_form_data(data):  # (8)
            return  # (12)
        if self.template_index is not None:  # (8)
            template_row = self.df.loc[self.template_index]  # (12)
            if is_exact_match(data, template_row):  # (12)
                messagebox.showerror(  # (16)
                    "Ошибка дубликата",  # (20)
                    "Новая запись полностью совпадает с шаблоном! Измените поле.",  # (20)
                )  # (16)
                return  # (16)
        processed_data = {}  # (8)
        for k, v in data.items():  # (8)
            if FIELDS_CONFIG[k] is int and v != "":  # (12)
                processed_data[k] = int(v)  # (16)
            else:  # (12)
                processed_data[k] = v  # (16)
        self.df = pd.concat(  # (8)
            [self.df, pd.DataFrame([processed_data])], ignore_index=True  # (12)
        )  # (8)
        save_data(self.df, FILE_NAME)  # (8)
        self.refresh_table()  # (8)
        self.clear_form()  # (8)
        messagebox.showinfo("Готово", "Запись успешно создана!")  # (8)

    def update_data(self):  # (4)
        if self.selected_index is None:  # (8)
            messagebox.showwarning(  # (12)
                "Внимание",  # (16)
                "Сначала выберите строку в таблице кликом мышки.",  # (16)
            )  # (12)
            return  # (12)

        data = self.get_form_values()  # (8)
        if not self.validate_form_data(data):  # (8)
            return  # (12)

        # Записываем измененные поля из формы обратно в self.df
        for k, v in data.items():  # (8)
            if FIELDS_CONFIG[k] is int:  # (12)
                self.df.at[self.selected_index, k] = int(v) if v != "" else 0  # (16)
            else:  # (12)
                self.df.at[self.selected_index, k] = v  # (16)

        # Сохраняем обновленный DataFrame в Excel-файл
        save_data(self.df, FILE_NAME)  # (8)

        # ВАЖНО: Перечитываем данные из файла, чтобы обновить оперативную память
        self.df = load_data(FILE_NAME)  # (8)

        # Перерисовываем таблицу и очищаем форму ввода
        self.refresh_table()  # (8)
        self.clear_form()  # (8)
        messagebox.showinfo("Готово", "Изменения успешно сохранены в таблицу и файл!")  # (8)
    def delete_data(self):  # (4)
        selected_item = self.tree.selection()  # (8)
        if not selected_item:  # (8)
            messagebox.showwarning(  # (12)
                "Удаление", "Сначала выберите строку в таблице для удаления."  # (16)
            )  # (12)
            return  # (12)
        target_index = int(selected_item[0])  # (8)
        if messagebox.askyesno(  # (8)
            "Подтверждение",  # (12)
            "Вы уверены, что хотите безвозвратно удалить эту запись?",  # (12)
        ):  # (8)
            self.df = delete_row_from_df(self.df, target_index)  # (12)
            save_data(self.df, FILE_NAME)  # (12)
            self.refresh_table()  # (12)
            self.clear_form()  # (12)
            messagebox.showinfo("Успех", "Запись успешно удалена.")  # (12)
    def search_data(self):  # (4)
        search_criteria = self.get_form_values()  # (8)
        filtered_df = self.df.copy()  # (8)
        for field, value in search_criteria.items():  # (8)
            if value != "":  # (12)
                if FIELDS_CONFIG[field] is int:  # (12)
                    filtered_df = filtered_df[  # (16)
                        filtered_df[field].astype(str) == value  # (20)
                    ]  # (16)
                else:  # (12)
                    filtered_df = filtered_df[  # (16)
                        filtered_df[field]  # (20)
                            .astype(str)  # (24)
                            .str.contains(value, case=False, na=False)  # (24)
                    ]  # (16)
        self.refresh_table(filtered_df)  # (8)

    def open_detail_view(self):  # (4)
        """Проверяет выделение строки и открывает карточный режим."""  # (8)
        selected_item = self.tree.selection()  # (8)

        if not selected_item:  # (8)
            messagebox.showwarning("Внимание",
                                   "Пожалуйста, выберите запись в таблице для перехода в режим «Детальный просмотр»!")  # (12)
            return  # (12)

        # Получаем iid строки, который в вашем коде соответствует индексу строки в DataFrame  # (8)
        pandas_index = int(selected_item[0])  # (8)

        # Запускаем окно карточек  # (8)
        DetailViewWindow(self, pandas_index)  # (8)

    def export_report(self):  # (4)
        # Определяем, какие данные сейчас отображаются в таблице
        # Если применен поиск, экспортируем отфильтрованный результат, иначе — всю базу
        search_criteria = self.get_form_values()  # (8)
        export_df = self.df.copy()  # (8)

        is_filtered = False  # (8)
        for field, value in search_criteria.items():  # (8)
            if value != "":  # (12)
                export_df = export_df[export_df[field].astype(str).str.contains(value, case=False, na=False)]  # (12)
                is_filtered = True  # (12)

        if export_df.empty:  # (8)
            messagebox.showwarning("Экспорт", "Нет данных для экспорта.")  # (12)
            return  # (12)

        # Открываем стандартное диалоговое окно Windows для сохранения файла
        file_path = filedialog.asksaveasfilename(  # (8)
            defaultextension=".txt",  # (12)
            filetypes=[("Текстовый файл", "*.txt"), ("Все файлы", "*.*")],  # (12)
            title="Сохранить отчет как..."  # (12)
        )  # (8)

        if file_path:  # (8)
            # Вызываем процедуру экспорта из helpers.py
            export_to_txt(export_df, file_path)  # (12)

            status = "отфильтрованных" if is_filtered else "всех"  # (12)
            messagebox.showinfo("Успех",
                                f"Отчет по форме ({status} записей: {len(export_df)}) успешно сохранен!")  # (12)

class DetailViewWindow:  # (0)
    def __init__(self, app_instance, start_index: int):  # (4)
        if not isinstance(start_index, int):  # (8)
            raise TypeError("Параметр start_index должен быть целым числом (int)")  # (12)

        self.app = app_instance  # (8)
        self.current_index = start_index  # (8)

        # Настройка модального окна  # (8)
        self.window = tk.Toplevel(self.app.root)  # (8)
        self.window.title("Режим: Детальный просмотр")  # (8)
        self.window.geometry("1400x980")  # (8)
        self.window.grab_set()  # (8)

        self.center_entries = {}  # (8)

        # Порядок сборки интерфейса важен для Pack: сначала верх и низ, потом центр  # (8)
        self.create_navigation_bar()  # (8)
        self.create_action_bar()  # (8)
        self.create_cards_layout()  # (8)

        # Первичное заполнение данными  # (8)
        self.refresh_cards()  # (8)
        # Перехватываем закрытие окна на крестик, чтобы тоже обновлять таблицу
        self.window.protocol("WM_DELETE_WINDOW", self.exit_and_refresh)  # (8)

    def create_navigation_bar(self):  # (4)
        nav_frame = tk.Frame(self.window, pady=10)  # (8)
        nav_frame.pack(fill=tk.X, side=tk.TOP)  # (8)

        self.btn_first = tk.Button(nav_frame, text="<< В начало", command=lambda: self.move_focus("first"))  # (8)
        self.btn_first.pack(side=tk.LEFT, padx=15)  # (8)

        self.btn_prev = tk.Button(nav_frame, text="< Назад", command=lambda: self.move_focus("prev"))  # (8)
        self.btn_prev.pack(side=tk.LEFT, padx=5)  # (8)

        self.lbl_position = tk.Label(nav_frame, text="", font=("Arial", 11, "bold"))  # (8)
        self.lbl_position.pack(side=tk.LEFT, expand=True)  # (8)

        self.btn_next = tk.Button(nav_frame, text="Вперед >", command=lambda: self.move_focus("next"))  # (8)
        self.btn_next.pack(side=tk.LEFT, padx=5)  # (8)

        self.btn_last = tk.Button(nav_frame, text="В конец >>", command=lambda: self.move_focus("last"))  # (8)
        self.btn_last.pack(side=tk.LEFT, padx=15)  # (8)

    def create_action_bar(self):  # (4)
        action_frame = tk.Frame(self.window, pady=12, bg="#f5f5f5")  # (8)
        action_frame.pack(fill=tk.X, side=tk.BOTTOM)  # (8)

        self.btn_edit = tk.Button(action_frame, text="Редактировать", width=15, bg="#FFC107", fg="black", font=("Arial", 10, "bold"), command=self.activate_editing)  # (8)
        self.btn_edit.pack(side=tk.LEFT, padx=20)  # (8)

        self.btn_save = tk.Button(action_frame, text="Сохранить", width=15, bg="#2196F3", fg="white", font=("Arial", 10, "bold"), state=tk.DISABLED, command=self.accept_card_changes)  # (8)
        self.btn_save.pack(side=tk.LEFT, padx=5)  # (8)

        self.btn_exit = tk.Button(action_frame, text="Выйти", width=15, bg="#f44336", fg="white", font=("Arial", 10, "bold"), command=self.exit_and_refresh)  # (8)
        self.btn_exit.pack(side=tk.RIGHT, padx=20)  # (8)

    def create_cards_layout(self):  # (4)
        cards_frame = tk.Frame(self.window)  # (8)
        # ИСПРАВЛЕНО: Добавлен expand=True, чтобы контейнер карточек рос по высоте окна
        cards_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)  # (8)

        # Настраиваем сетку: 1 строка во всю высоту и 3 равные колонки
        cards_frame.rowconfigure(0, weight=1)  # (8)
        cards_frame.columnconfigure(0, weight=1, uniform="cards")  # (8)
        cards_frame.columnconfigure(1, weight=1, uniform="cards")  # (8)
        cards_frame.columnconfigure(2, weight=1, uniform="cards")  # (8)

        self.frame_left = tk.LabelFrame(cards_frame, text="Предыдущая запись", padx=10, pady=10, fg="gray")  # (8)
        self.frame_right = tk.LabelFrame(cards_frame, text="Следующая запись", padx=10, pady=10, fg="gray")  # (8)

        self.frame_center = tk.LabelFrame(  # (8)
            cards_frame,  # (12)
            text="---Центральная запись--- (Редактирование)",  # (12)
            bd=2, relief=tk.SOLID, padx=5, pady=5, fg="black"  # (12)
        )  # (8)

        self.frame_left.grid(row=0, column=0, sticky="nsew", padx=5)  # (8)
        self.frame_center.grid(row=0, column=1, sticky="nsew", padx=5)  # (8)
        self.frame_right.grid(row=0, column=2, sticky="nsew", padx=5)  # (8)

    def refresh_cards(self):
        context = get_card_context(self.app.df, self.current_index)  # (8)
        self.lbl_position.config(text=context["current_pos_text"])  # (8)
        self.btn_first.config(state=tk.NORMAL if context["has_prev"] else tk.DISABLED)  # (8)
        self.btn_prev.config(state=tk.NORMAL if context["has_prev"] else tk.DISABLED)  # (8)
        self.btn_next.config(state=tk.NORMAL if context["has_next"] else tk.DISABLED)  # (8)
        self.btn_last.config(state=tk.NORMAL if context["has_next"] else tk.DISABLED)  # (8)

        for frame in (self.frame_left, self.frame_right):  # (8)
            for w in frame.winfo_children(): w.destroy()  # (12)
        for w in self.frame_center.winfo_children(): w.destroy()  # (8)
        self.center_entries.clear()  # (8)

        # Используем общий метод сетки для всех трех карточек
        if context["left_data"] == "START":  # (8)
            tk.Label(self.frame_left, text="[ НАЧАЛО БАЗЫ ]", font=("Arial", 12, "bold"), fg="gray").pack(expand=True)  # (12)
        else:  # (8)
            # Передаем False для боковых, чтобы они были read-only
            self.draw_scrollable_editable_view(self.frame_left, context["left_data"], is_center=False)  # (12)

        self.draw_scrollable_editable_view(self.frame_center, context["center_data"], is_center=True)  # (8)

        if context["right_data"] == "END":  # (8)
            tk.Label(self.frame_right, text="[ КОНЕЦ БАЗЫ ]", font=("Arial", 12, "bold"), fg="gray").pack(expand=True)  # (12)
        else:  # (8)
            self.draw_scrollable_editable_view(self.frame_right, context["right_data"], is_center=False)  # (12)

        self.btn_save.config(state=tk.DISABLED)  # (8)
        self.btn_edit.config(state=tk.NORMAL)  # (8)


    def draw_scrollable_editable_view(self, parent_frame, data_dict, is_center=True):
        # Настраиваем цвет фона: для центра белый, для боков — мягкий серый
        bg_color = "#ffffff" if is_center else "#f5f5f5"
        parent_frame.config(bg=bg_color)

        canvas = tk.Canvas(parent_frame, highlightthickness=0, bg=bg_color)
        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=bg_color)

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))  # (8)
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")  # (8)
        canvas.configure(yscrollcommand=scrollbar.set)  # (8)

        # Безопасная функция прокрутки: проверяем, существует ли еще виджет  # (8)
        def _on_mousewheel(event):  # (8)
            try:  # (12)
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")  # (16)
            except tk.TclError:  # (12)
                pass  # (16) Защита на случай, если виджет удален

        # ИСПРАВЛЕНО: Привязываем колесико мыши только при входе курсора в зону карточки  # (8)
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))  # (8)
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))  # (8)
        scroll_frame.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))  # (8)
        scroll_frame.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))  # (8)

        canvas.pack(side="left", fill="both", expand=True)  # (8)
        scrollbar.pack(side="right", fill="y")  # (8)

        # Конфигурируем 6 микро-колонок для идеального выравнивания любых комбинаций по ширине  # (8)
        for c in range(6):  # (8)
            scroll_frame.columnconfigure(c, weight=1, minsize=60)  # (12)

        # Словарь попарных полей для 2 и 3 секций  # (8)
        paired_fields = {  # (8)
            "№ этапа рейса": "Тип данных",  # (12)
            "№ диска": "Инвентарный № диска"  # (12)
        }  # (8)

        # Полный список полей, которые мы пропускаем при обычном обходе  # (8)
        skip_fields = [  # (8)
            "Год", "Рейс", "№ рейса", "Дата окончания рейса",  # (12)
            "Тип данных", "Инвентарный № диска"  # (12)
        ]  # (8)

        grid_row = 0  # (8)

        for group_name, fields in TABS_CONFIG.items():  # (8)
            group_lbl = tk.Label(scroll_frame, text=f"--------------- {group_name.upper()} ---------------", font=("Arial", 10, "bold"), fg="black", bg=bg_color, pady=8)  # (12)
            group_lbl.grid(row=grid_row, column=0, columnspan=6, sticky="w", padx=5)  # (12)
            grid_row += 1  # (12)

            for field in fields:  # (12)
                if field in skip_fields:  # (16)
                    continue  # (20)

                # Строка 1: Трио [ № п/п ] [ Год ] [ Рейс ]  # (16)
                if field == "№ п/п":  # (16)
                    t_fields = ["№ п/п", "Год", "Рейс"]  # (20)
                    for col_idx, t_field in enumerate(t_fields):  # (20)
                        sub_frame = tk.Frame(scroll_frame, bg="#ffffff")  # (24)
                        sub_frame.grid(row=grid_row, column=col_idx*2, columnspan=2, sticky="ew", padx=4, pady=3)  # (24)
                        lbl = tk.Label(sub_frame, text=t_field, font=("Arial", 9, "italic"), bg=bg_color, fg="#333333")  # (24)
                        lbl.pack(anchor="w")  # (24)
                        # текст в полях вертикальных карточек
                        entry = tk.Entry(sub_frame, font=("Arial", 10, "bold"), bd=1, relief=tk.GROOVE, fg="#000055")  # (24)

                        val = data_dict.get(t_field, "")  # (24)
                        entry.insert(0, str(val) if pd.notna(val) else "")  # (24)
                        entry.config(state="readonly")  # (24)
                        entry.pack(fill=tk.X, pady=(2, 0))  # (24)
                        if is_center:
                            self.center_entries[t_field] = entry  # (24)
                    grid_row += 1  # (20)

                # Строка 2: Пара [ Судно ] [ № рейса ]  # (16)
                elif field == "Судно":  # (16)
                    p_fields = ["Судно", "№ рейса"]  # (20)
                    for col_idx, p_field in enumerate(p_fields):  # (20)
                        sub_frame = tk.Frame(scroll_frame, bg="#ffffff")  # (24)
                        sub_frame.grid(row=grid_row, column=col_idx*3, columnspan=3, sticky="ew", padx=4, pady=3)  # (24)
                        lbl = tk.Label(sub_frame, text=p_field, font=("Arial", 9, "italic"), bg=bg_color, fg="#333333")  # (24)
                        lbl.pack(anchor="w")  # (24)
                        # текст в полях вертикальных карточек
                        entry = tk.Entry(sub_frame, font=("Arial", 10, "bold"), bd=1, relief=tk.GROOVE, fg="#000055")  # (24)

                        val = data_dict.get(p_field, "")  # (24)
                        entry.insert(0, str(val) if pd.notna(val) else "")  # (24)
                        entry.config(state="readonly")  # (24)
                        entry.pack(fill=tk.X, pady=(2, 0))  # (24)
                        if is_center:
                            self.center_entries[p_field] = entry # (24)
                    grid_row += 1  # (20)

                # Строка 3: Пара [ Дата начала рейса ] [ Дата окончания рейса ]  # (16)
                elif field == "Дата начала рейса":  # (16)
                    d_fields = ["Дата начала рейса", "Дата окончания рейса"]  # (20)
                    for col_idx, d_field in enumerate(d_fields):  # (20)
                        sub_frame = tk.Frame(scroll_frame, bg="#ffffff")  # (24)
                        sub_frame.grid(row=grid_row, column=col_idx*3, columnspan=3, sticky="ew", padx=4, pady=3)  # (24)
                        lbl = tk.Label(sub_frame, text=d_field, font=("Arial", 9, "italic"), bg=bg_color, fg="#333333")  # (24)
                        lbl.pack(anchor="w")  # (24)
                        # текст в полях вертикальных карточек
                        entry = tk.Entry(sub_frame, font=("Arial", 10, "bold"), bd=1, relief=tk.GROOVE, fg="#000055")  # (24)

                        val = data_dict.get(d_field, "")  # (24)
                        entry.insert(0, str(val) if pd.notna(val) else "")  # (24)
                        entry.config(state="readonly")  # (24)
                        entry.pack(fill=tk.X, pady=(2, 0))  # (24)
                        if is_center:
                            self.center_entries[d_field] = entry # (24)
                    grid_row += 1  # (20)

                # Секции 2 и 3: попарная компоновка  # (16)
                elif field in paired_fields:  # (16)
                    partner_field = paired_fields[field]  # (20)
                    sub_frame_left = tk.Frame(scroll_frame, bg="#ffffff")  # (20)
                    sub_frame_left.grid(row=grid_row, column=0, columnspan=3, sticky="ew", padx=4, pady=3)  # (20)
                    lbl_l = tk.Label(sub_frame_left, text=field, font=("Arial", 9, "italic"), bg=bg_color, fg="#333333")  # (20)
                    lbl_l.pack(anchor="w")  # (20)
                    # текст в полях вертикальных карточек
                    entry_l = tk.Entry(sub_frame_left, font=("Arial", 10, "bold"), bd=1, relief=tk.GROOVE, fg="#000055")  # (20)

                    val_l = data_dict.get(field, "")  # (20)
                    entry_l.insert(0, str(val_l) if pd.notna(val_l) else "")  # (20)
                    entry_l.config(state="readonly")  # (20)
                    entry_l.pack(fill=tk.X, pady=(2, 0))  # (20)
                    # ИСПРАВЛЕНИЕ: Записываем левое поле только если это центр
                    if is_center:
                        self.center_entries[field] = entry_l  # (20)

                    sub_frame_right = tk.Frame(scroll_frame, bg="#ffffff")  # (20)
                    sub_frame_right.grid(row=grid_row, column=3, columnspan=3, sticky="ew", padx=4, pady=3)  # (20)
                    lbl_r = tk.Label(sub_frame_right, text=partner_field, font=("Arial", 9, "italic"), bg=bg_color, fg="#333333")  # (20)
                    lbl_r.pack(anchor="w")  # (20)
                    # текст в полях вертикальных карточек
                    entry_r = tk.Entry(sub_frame_right, font=("Arial", 10, "bold"), bd=1, relief=tk.GROOVE, fg="#000055")  # (20)

                    val_r = data_dict.get(partner_field, "")  # (20)
                    entry_r.insert(0, str(val_r) if pd.notna(val_r) else "")  # (20)
                    entry_r.config(state="readonly")  # (20)
                    entry_r.pack(fill=tk.X, pady=(2, 0))  # (20)
                    # ИСПРАВЛЕНИЕ: Записываем правое поле только если это центр
                    if is_center:
                        self.center_entries[partner_field] = entry_r  # (20)
                    grid_row += 1  # (20)

                # Одиночные длинные поля  # (16)
                else:  # (16)
                    sub_frame_full = tk.Frame(scroll_frame, bg="#ffffff")  # (20)
                    sub_frame_full.grid(row=grid_row, column=0, columnspan=6, sticky="ew", padx=4, pady=3)  # (20)
                    lbl = tk.Label(sub_frame_full, text=field, font=("Arial", 9, "italic"), bg=bg_color, fg="#333333")  # (20)
                    lbl.pack(anchor="w")  # (20)
                    # текст в полях вертикальных карточек
                    entry = tk.Entry(sub_frame_full, font=("Arial", 10, "bold"), bd=1, relief=tk.GROOVE, fg="#000055")  # (20)

                    val = data_dict.get(field, "")  # (20)
                    entry.insert(0, str(val) if pd.notna(val) else "")  # (20)
                    entry.config(state="readonly")  # (20)
                    entry.pack(fill=tk.X, pady=(2, 0))  # (20)
                    if is_center:
                        self.center_entries[field] = entry
                    grid_row += 1  # (20)


    def move_focus(self, direction: str):  # (4)
        if direction == "prev":  # (8)
            self.current_index -= 1  # (12)
        elif direction == "next":  # (8)
            self.current_index += 1  # (12)
        elif direction == "first":  # (8)
            self.current_index = 0  # (12)
        elif direction == "last":  # (8)
            self.current_index = len(self.app.df) - 1  # (12)

        self.refresh_cards()  # (8)

    def activate_editing(self):  # (4)
        for entry in self.center_entries.values():  # (8)
            entry.config(state="normal")  # (12)
        self.btn_save.config(state=tk.NORMAL)  # (8)
        self.btn_edit.config(state=tk.DISABLED)  # (8)

    #def save_data(self):  # (4)
    def accept_card_changes(self):
        updated_row = {}  # (8)

        # Списки числовых полей из условий вашей задачи для явного приведения типов
        int_fields = ["№ п/п", "Год", "№ рейса", "№ этапа рейса", "№ диска", "Инвентарный № диска"]  # (8)

        for field, entry in self.center_entries.items():  # (8)
            val = entry.get().strip()  # (12)

            # Если поле должно быть числовым и оно заполнено — переводим в int
            if field in int_fields:  # (12)
                if val != "":  # (16)
                    try:  # (20)
                        updated_row[field] = int(val)  # (24)
                    except ValueError:  # (20)
                        messagebox.showerror("Ошибка типа данных",
                                             f"Поле '{field}' должно содержать только целое число!")  # (24)
                        return  # (24)
                else:  # (16)
                    updated_row[field] = np.nan  # (20) Поле не обязательно к заполнению
            else:  # (12)
                updated_row[field] = val  # (16) Для всех текстовых полей оставляем как есть

        # Сохраняем проверенные данные в Pandas DataFrame
        for field, value in updated_row.items():  # (8)
            self.app.df.at[self.current_index, field] = value  # (12)

        # УНИВЕРСАЛЬНОЕ ОБНОВЛЕНИЕ ТАБЛИЦЫ НА ГЛАВНОМ ЭКРАНЕ (без вызова внешних функций)
        try:  # (8)
            # 1. Очищаем старые строки в вашей таблице Treeview
            for item in self.app.tree.get_children():  # (12)
                self.app.tree.delete(item)  # (16)

            # 2. Заново заполняем таблицу актуальными данными из измененного DataFrame
            for idx, row in self.app.df.iterrows():  # (12)
                self.app.tree.insert("", "end", iid=idx, values=list(row))  # (16)
        except Exception as e:  # (8)
            print(f"Предупреждение при отрисовке таблицы: {e}")  # (12)

        # messagebox.showinfo("Успех", "Изменения успешно сохранены в оперативную память!")  # (8)
        save_data(self.app.df, FILE_NAME)
        self.refresh_cards()  # (8)

    def exit_and_refresh(self):  # (4)
        """Принудительно синхронизирует данные и закрывает окно."""  # (8)
        # ИСПРАВЛЕНО: Гарантированно отвязываем колесико мыши от памяти системы
        try:  # (8)
            self.window.unbind_all("<MouseWheel>")  # (12)
        except Exception:  # (8)
            pass  # (12)

        try:  # (8)
            for item in self.app.tree.get_children():  # (12)
                self.app.tree.delete(item)  # (16)
            for idx, row in self.app.df.iterrows():  # (12)
                self.app.tree.insert("", "end", iid=idx, values=list(row))  # (16)
        except Exception:  # (8)
            pass  # (12)

        self.app.root.update_idletasks()  # (8)
        self.window.destroy()  # (8)


if __name__ == "__main__":
    # 1. Запоминаем хэш файла ДО запуска программы
    start_hash = get_file_hash(FILE_NAME)
    # 2. Делаем обязательный бэкап при старте (как у вас и было)
    create_backup(FILE_NAME)

    root = tk.Tk()
    app = VoyageAppTabs(root)

    # Функция безопасного закрытия
    def on_closing():
        # Отключаем перехватчик, чтобы избежать зацикливания при закрытии
        root.protocol("WM_DELETE_WINDOW", lambda: None)

        # 3. Проверяем хэш файла в момент закрытия
        end_hash = get_file_hash(FILE_NAME)

        # Если хэши разные — значит, пользователь вносил изменения или сохранял данные
        if start_hash != end_hash:
            print("Обнаружены изменения в базе данных. Создается финальный бэкап...")
            create_backup(FILE_NAME)
        else:
            print("Изменений не было. Финальный бэкап пропущен.")

        clean_duplicate_backups()  # Функция сама проверит маркер и решит, запускаться или нет

        # 4. Жестко и последовательно останавливаем интерфейс Tkinter
        root.quit()
        root.destroy()

    # Привязываем закрытие окна на крестик к нашей чистой функции
    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()
