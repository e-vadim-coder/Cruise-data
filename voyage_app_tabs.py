# -*- coding: utf-8 -*-
# voyage_app_tabs.py

import tkinter as tk  # (0)
from tkinter import ttk, messagebox  # (0)
import pandas as pd  # (0)
import numpy as np  # (0)
import os  # (0)
from config import FIELDS_CONFIG, TABS_CONFIG, DICTIONARIES  # (0)


class VoyageAppTabs:  # (0)
    def __init__(self, root):  # (4)
        self.root = root  # (8)
        self.file_name = "voyage_data.xlsx"  # (8)
        self.df = self.initial_load_df()  # (8)
        self.selected_pk = None  # (8)
        self.form_fields = {}  # (8)
        self.create_widgets()  # (8)
        self.refresh_table()  # (8)

    def initial_load_df(self) -> pd.DataFrame:  # (4)
        columns = list(FIELDS_CONFIG.keys())  # (8)
        if not os.path.exists(self.file_name):  # (8)
            return pd.DataFrame(columns=columns)  # (12)
        try:  # (8)
            raw_df = pd.read_excel(self.file_name)  # (12)
            rename_map = {info["short"]: key for key, info in FIELDS_CONFIG.items()}  # (12)
            raw_df = raw_df.rename(columns=rename_map)  # (12)

            df = pd.DataFrame(columns=columns)  # (12)
            for col in columns:  # (12)
                if col in raw_df.columns:  # (16)
                    df[col] = raw_df[col]  # (20)

            for col, info in FIELDS_CONFIG.items():  # (12)
                if info["type"] is int:  # (16)
                    df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')  # (20)
                else:  # (16)
                    df[col] = df[col].fillna("").astype(str).str.strip()  # (20)
            return df  # (12)
        except Exception:  # (8)
            return pd.DataFrame(columns=columns)  # (12)

    def save_df_to_disk(self):  # (4)
        export_df = self.df.copy()  # (8)
        rename_map = {key: info["short"] for key, info in FIELDS_CONFIG.items()}  # (8)
        export_df = export_df.rename(columns=rename_map)  # (8)
        export_df.to_excel(self.file_name, index=False)  # (8)

    def create_widgets(self):  # (4)
        # Главный фрейм-контейнер, занимающий верхнюю и нижнюю части окна
        main_frame = tk.Frame(self.root, bg="#f5f5f5")  # (8)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # (8)

        # 1. ВЕРХНЯЯ ПАНЕЛЬ: Вкладки с полями ввода в два столбца
        form_panel = tk.Frame(main_frame, bg="#f5f5f5")  # (8)
        form_panel.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))  # (8)

        self.notebook = ttk.Notebook(form_panel)  # (8)
        self.notebook.pack(fill=tk.X, expand=True)  # (8)

        for tab_name, fields in TABS_CONFIG.items():  # (8)
            tab_frame = tk.Frame(self.notebook, bg="#ffffff", padx=15, pady=10)  # (12)
            self.notebook.add(tab_frame, text=tab_name)  # (12)

            # Настраиваем веса столбцов для правильного распределения по ширине
            # tab_frame.columnconfigure(1, weight=1)  # (12)
            # tab_frame.columnconfigure(3, weight=1)  # (12)

            # Строим сетку строго в два столбца по вашим скриншотам

            # Строим сетку в два столбца на основе латинских ключей
            for idx, key in enumerate(fields):  # (12)
                info = FIELDS_CONFIG[key]  # (16)
                r = idx // 2  # (16)
                c = (idx % 2) * 2  # (16)

                # Берём развёрнутое красивое имя "full" для подписи на вкладке сверху
                lbl = tk.Label(tab_frame, text=info["full"] + ":", bg="#ffffff", anchor="w")  # (16)
                lbl.grid(row=r, column=c, sticky="w", pady=4, padx=(10, 5))  # (16)

                # Ищем выпадающий список в config.py по латинскому ключу переменной (key)
                if key in DICTIONARIES:  # (16)
                    widget = ttk.Combobox(tab_frame, values=DICTIONARIES[key], state="readonly", width=25)  # (20)
                else:  # (16)
                    widget = tk.Entry(tab_frame, bd=1, relief=tk.SOLID, width=27)  # (20)

                widget.grid(row=r, column=c + 1, sticky="ew", pady=4, padx=(0, 20))  # (16)
                self.form_fields[key] = widget  # (16)

        # 2. СРЕДНЯЯ ПАНЕЛЬ: Оригинальный разноцветный ряд кнопок управления
        btn_frame = tk.Frame(main_frame, bg="#f5f5f5")  # (8)
        btn_frame.pack(side=tk.TOP, fill=tk.X, pady=10)  # (8)

        tk.Button(btn_frame, text="🔍 Поиск", command=self.search_data, bg="#0056b3", fg="white").pack(side=tk.LEFT,
                                                                                                      padx=3)  # (8)
        tk.Button(btn_frame, text="➕ Создать новую", command=self.add_data, bg="#28a745", fg="white").pack(side=tk.LEFT,
                                                                                                           padx=3)  # (8)
        tk.Button(btn_frame, text="📋 Из шаблона", command=lambda: None, bg="#6f42c1", fg="white").pack(side=tk.LEFT,
                                                                                                       padx=3)  # (8)
        tk.Button(btn_frame, text="💾 Сохранить изменения", command=self.update_data, bg="#fd7e14", fg="white").pack(
            side=tk.LEFT, padx=3)  # (8)
        tk.Button(btn_frame, text="❌ Удалить запись", command=self.delete_data, bg="#dc3545", fg="white").pack(
            side=tk.LEFT, padx=3)  # (8)
        tk.Button(btn_frame, text="Детальный просмотр", command=self.open_detail_view, bg="#b5835a", fg="white").pack(
            side=tk.LEFT, padx=3)  # (8)
        tk.Button(btn_frame, text="🧹 Сбросить форму", command=self.clear_form, bg="#6c757d", fg="white").pack(
            side=tk.LEFT, padx=3)  # (8)
        tk.Button(btn_frame, text="📊 Экспорт отчета", command=lambda: None, bg="#20c997", fg="white").pack(side=tk.LEFT,
                                                                                                           padx=3)  # (8)

        # 3. НИЖНЯЯ ПАНЕЛЬ: Рамка, оригинальный заголовок и таблица Treeview
        lbl_table_title = tk.Label(main_frame, text="Просмотр базы данных рейсов", bg="#f5f5f5",
                                   font=("Arial", 10, "bold"), anchor="w")  # (8)
        lbl_table_title.pack(side=tk.TOP, fill=tk.X, pady=(5, 2))  # (8)

        table_frame = tk.Frame(main_frame, bg="#ffffff", bd=1, relief=tk.SOLID)  # (8)
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)  # (8)

        self.table_keys = [k for k, _ in sorted(FIELDS_CONFIG.items(), key=lambda x: x[1]["pos_table"])]  # (8)
        headers = [FIELDS_CONFIG[k]["short"] for k in self.table_keys]                   # (8)


        self.tree = ttk.Treeview(table_frame, columns=headers, show="headings", selectmode="browse")  # (8)
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)  # (8)

        scroll_y = tk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)  # (8)
        self.tree.configure(yscrollcommand=scroll_y.set)  # (8)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)  # (8)

        # Настраиваем заголовки и индивидуальную ширину колонок таблицы Treeview
        for h, k in zip(headers, self.table_keys):                                       # (8)
            self.tree.heading(h, text=h)                                                 # (12)
            # Считываем ширину из config.py. Если её там нет, ставим стандарт: 90 пикселей
            col_width = FIELDS_CONFIG[k].get("width", 90)                                # (12)
            self.tree.column(h, width=col_width, anchor="center")                        # (12)


        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)  # (8)

    def refresh_table(self, dataframe: pd.DataFrame = None):  # (4)
        for item in self.tree.get_children():  # (8)
            self.tree.delete(item)  # (12)
        display_df = self.df if dataframe is None else dataframe  # (8)

        for _, row in display_df.iterrows():  # (8)
            values = []  # (12)
            for key in self.table_keys:  # (12)
                val = row[key]  # (16)


                if pd.isna(val):                                                         # (16)
                    values.append("")                                                    # (20)
                elif FIELDS_CONFIG[key]["type"] is int:                                  # (16)
                    values.append(str(int(val)))                                         # (20)
                else:                                                                    # (16)
                    values.append(str(val))                                              # (20)
            pk = str(int(row["id"])) if pd.notna(row["id"]) else str(np.random.randint(10000, 99999)) # (12)
            self.tree.insert("", "end", iid=pk, values=values)                           # (12)

    def on_row_select(self, event):  # (4)
        selected = self.tree.selection()  # (8)
        if not selected:  # (8)
            return  # (12)
        # Берем первый элемент из кортежа выделения, чтобы избежать TypeError
        self.selected_pk = int(selected[0])  # (8)

        rows = self.df[self.df["id"] == self.selected_pk].index  # (8)
        if len(rows) == 0:  # (8)
            return  # (12)
        row_idx = rows[0]  # (8)

        for key, widget in self.form_fields.items():  # (8)
            val = self.df.at[row_idx, key]  # (12)
            text_val = "" if pd.isna(val) else (
                str(int(val)) if FIELDS_CONFIG[key]["type"] is int else str(val))  # (12)

            if isinstance(widget, ttk.Combobox):  # (12)
                widget.set(text_val)  # (16)
            else:  # (12)
                widget.delete(0, tk.END)  # (16)
                widget.insert(0, text_val)  # (16)

    def get_form_values(self) -> dict:                                                   # (4)
        data = {}                                                                        # (8)
        for key, widget in self.form_fields.items():                                     # (8)
            val = widget.get().strip()                                                   # (12)
            if val == "":                                                                # (12)
                data[key] = pd.NA                                                        # (16)
            elif FIELDS_CONFIG[key]["type"] is int:                                      # (12)
                try:                                                                     # (16)
                    data[key] = int(val)                                                 # (20)
                except ValueError:                                                       # (16)
                    data[key] = "ERROR"                                                  # (20)
            else:                                                                        # (12)
                data[key] = val                                                          # (16)
        return data                                                                      # (8)

    def add_data(self):                                                                  # (4)
        data = self.get_form_values()                                                    # (8)
        if "ERROR" in data.values():                                                     # (8)
            messagebox.showerror("Ошибка", "Числовые поля заполнены некорректно!")       # (12)
            return                                                                       # (12)
        max_id = self.df["id"].max() if not self.df.empty else 0                         # (8)
        new_id = 1 if pd.isna(max_id) else int(max_id) + 1                               # (8)
        data["id"] = new_id                                                              # (8)
        self.df = pd.concat([self.df, pd.DataFrame([data])], ignore_index=True)           # (8)
        self.save_df_to_disk()                                                           # (8)
        self.refresh_table()                                                             # (8)
        self.clear_form()                                                                # (8)

    def update_data(self):                                                               # (4)
        if self.selected_pk is None:                                                     # (8)
            return                                                                       # (12)
        data = self.get_form_values()                                                    # (8)
        if "ERROR" in data.values():                                                     # (8)
            messagebox.showerror("Ошибка", "Числовые поля заполнены некорректно!")       # (12)
            return                                                                       # (12)
        idx_list = self.df[self.df["id"] == self.selected_pk].index                      # (8)
        if len(idx_list) == 0:                                                           # (8)
            return                                                                       # (12)
        data["id"] = self.selected_pk                                                    # (8)
        for key, val in data.items():                                                    # (8)
            self.df.at[idx_list, key] = val                                              # (12)
        self.save_df_to_disk()                                                           # (8)
        self.refresh_table()                                                             # (8)

    def delete_data(self):                                                               # (4)
        if self.selected_pk is None:                                                     # (8)
            return                                                                       # (12)
        if not messagebox.askyesno("Подтверждение", "Удалить выбранную запись?"):        # (8)
            return                                                                       # (12)
        self.df = self.df[self.df["id"] != self.selected_pk].reset_index(drop=True)      # (8)
        self.save_df_to_disk()                                                           # (8)
        self.selected_pk = None                                                          # (8)
        self.refresh_table()                                                             # (8)
        self.clear_form()                                                                # (8)

    def clear_form(self):                                                                # (4)
        self.selected_pk = None                                                          # (8)
        for widget in self.form_fields.values():                                         # (8)
            if isinstance(widget, ttk.Combobox):                                         # (12)
                widget.set("")                                                           # (16)
            else:                                                                        # (12)
                widget.delete(0, tk.END)                                                 # (16)

    def search_data(self):                                                               # (4)
        data = self.get_form_values()                                                    # (8)
        filtered_df = self.df.copy()                                                     # (8)
        for key, crit in data.items():                                                   # (8)
            if pd.isna(crit) or crit == "ERROR":                                         # (12)
                continue                                                                 # (16)
            if FIELDS_CONFIG[key]["type"] is int:                                        # (12)
                filtered_df = filtered_df[filtered_df[key] == int(crit)]                 # (16)
            else:                                                                        # (12)
                filtered_df = filtered_df[filtered_df[key].astype(str).str.contains(str(crit), case=False, na=False)] # (16)
        self.refresh_table(filtered_df)                                                  # (8)

    def open_detail_view(self):                                                          # (4)
        from main import DetailViewWindow                                                # (8)
        if self.df.empty:                                                                # (8)
            return                                                                       # (12)
        current_idx = 0                                                                  # (8)
        if self.selected_pk is not None:                                                 # (8)
            idx_list = self.df[self.df["id"] == self.selected_pk].index                      # (12)
            if len(idx_list) > 0:                                                        # (12)
                current_idx = idx_list[0]  # (16)
        DetailViewWindow(self.root, self, current_idx)                                   # (8)