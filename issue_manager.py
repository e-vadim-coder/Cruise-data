# -*- coding: utf-8 -*-
# issue_manager.py
"""Учёт выдачи акустических данных научным сотрудникам по запросу.

Окно журнала выдачи (модальный Toplevel внутри «Учёт рейсовых данных»).
Запускается из main.py кнопкой «Журнал выдачи акустических данных» для
выбранного рейса. Может также запускаться как самостоятельное приложение
(для отладки).

Функции:
  - при открытии с выбранным рейсом спрашивает: «Добавить текущую запись
    в БД выданных акустических данных?»;
  - просмотр/поиск журнала выдач;
  - добавление (с предзаполнением контекста рейса), редактирование, удаление;
  - экспорт отчёта в TXT/XLSX;
  - отдельный бэкап журнала (issue_backups/).
"""

import os
from datetime import datetime
import tkinter as tk
from tkinter import font as tkfont
from tkinter import filedialog, messagebox, ttk

import pandas as pd

try:
    from helpers import FIELDS_CONFIG, SUB_TABLE_FIELDS
except Exception:  # автономный запуск без helpers
    FIELDS_CONFIG = {
        "№ п/п": int, "Год": int, "Рейс": str, "Судно": str, "№ рейса": int,
        "Район исследования": str, "Тип данных": str, "Прибор": str,
        "Степень обработки": str, "Формат файла": str, "Объем данных": str,
    }
    SUB_TABLE_FIELDS = {
        "UID_Родителя": str, "каталоги": str, "каталоги с копией": str,
    }

VOYAGE_FILE = "voyage_data.xlsx"
SUB_FILE = "voyage_sub_data.xlsx"
ISSUE_FILE = "issue_register.xlsx"
BACKUP_DIR = "issue_backups"
MAX_BACKUPS = 50

# Поля рейса, копируемые в запись выдачи: ВСЕ поля voyage_data кроме UID.
# Динамически берём из FIELDS_CONFIG (helpers), чтобы автоматически подхватывать состав.
COPIED_FIELDS = [f for f in FIELDS_CONFIG.keys() if f != "UID"]

# Поля, добавляемые к выдаче (собственно учёт выдачи)
ISSUE_FIELDS = {
    "UID": str,                    # ключ выдачи (для связей/стабильности)
    "UID_рейса": str,              # ссылка на рейс
    "выданные каталоги": str,      # все значения «каталоги» связанных записей sub (заполняется авто)
    "Дата выдачи": str,
    "ФИО сотрудника": str,
    "Подразделение": str,
    "Цель выдачи": str,
    "срок моратория/реализации": str,
    "Научная задача": str,
    "Отметка о реализации": str,   # «Данные обработаны» или выходные данные статьи
    "Дата выполнения": str,        # заменяет «Дата возврата»
    "Комментарий": str,
}

# Объединённый список: контекст рейса + поля выдачи
ISSUE_ROW_FIELDS = COPIED_FIELDS + list(ISSUE_FIELDS.keys())

TABLE_HEADINGS_SHORT = {
    "№ п/п": "№ п/п", "Год": "Год", "Рейс": "Рейс", "Судно": "Судно",
    "№ рейса": "№ рейса", "Район исследования": "Район",
    "Тип данных": "Тип", "Прибор": "Прибор", "Формат файла": "Формат",
    "Объем данных": "Объем", "Дата выдачи": "Дата выдачи",
    "ФИО сотрудника": "Сотрудник", "Подразделение": "Подразделение",
    "Цель выдачи": "Цель", "выданные каталоги": "Выданные каталоги",
    "срок моратория/реализации": "Срок моратория/реализации",
    "Научная задача": "Научная задача", "Отметка о реализации": "Отметка о реализации",
    "Дата выполнения": "Дата выполнения", "Комментарий": "Комментарий",
}

# Возможные цели выдачи (выпадающий список)
ACTIVITY_OPTIONS = [
    "Для обработки данных",
    "Для написания статьи",
    "Для доклада / презентации",
    "Для отчёта",
    "Другое",
]

# Возможные значения «Отметки о реализации»
REALIZATION_OPTIONS = ["Данные обработаны"]

# Поля, отображаемые в таблице выдачи (визуальный порядок)
DISPLAY_COLUMNS = [
    "№ п/п", "Год", "Рейс", "Судно", "Тип данных", "Прибор",
    "Дата выдачи", "ФИО сотрудника", "Подразделение", "Цель выдачи",
    "выданные каталоги", "срок моратория/реализации", "Научная задача",
    "Отметка о реализации", "Дата выполнения", "Комментарий",
]


def load_frame(file_name, columns):
    """Загружает xlsx в DataFrame с нужными колонками (пустые -> "")."""
    if os.path.exists(file_name):
        try:
            df = pd.read_excel(file_name)
        except Exception:
            df = pd.DataFrame()
    else:
        df = pd.DataFrame()
    if df.empty:
        df = pd.DataFrame(columns=columns)
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    for c in df.columns:
        if c not in columns:
            df = df.drop(columns=[c])
    return df


def fmt(v):
    """Аккуратный вывод значения для таблицы/отчёта."""
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def make_backup():
    """Простой бэкап журнала выдач (с учётом лимита)."""
    if not os.path.exists(ISSUE_FILE):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    import glob
    import zipfile
    zips = sorted(glob.glob(os.path.join(BACKUP_DIR, "*.zip")), key=os.path.getmtime)
    while len(zips) >= MAX_BACKUPS:
        try:
            os.remove(zips.pop(0))
        except Exception:
            break
    name = os.path.join(BACKUP_DIR,
                        f"issue_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip")
    with zipfile.ZipFile(name, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(ISSUE_FILE, ISSUE_FILE)
    return name


class IssueManagerWindow:
    """Модальное окно журнала выдачи данных.

    parent  — главное окно приложения (VoyageAppTabs или tk root).
    init_row — dict с данными заранее выбранного рейса (или None).
    """

    def __init__(self, parent, init_row=None, on_exit=None):
        self.root = parent.root if hasattr(parent, "root") else parent
        # on_exit — необязательный колбэк полного выхода из программы
        # (в main.py это on_closing, который делает финальный бэкап основной БД).
        self.on_exit = on_exit

        self.window = tk.Toplevel(self.root)
        self.window.title("Учет выдачи акустических данных научным сотрудникам")
        self.window.geometry("1500x950")
        self.window.state("zoomed")  # развернуть окно на весь экран
        self.window.transient(self.root)
        self.window.grab_set()

        self.voyage_df = load_frame(VOYAGE_FILE, list(FIELDS_CONFIG.keys()) + ["UID"])
        self.sub_df = load_frame(SUB_FILE, ["UID_Родителя", "каталоги", "каталоги с копией"])
        self.issue_df = load_frame(ISSUE_FILE, ISSUE_ROW_FIELDS)

        self.selected_issue_index = None
        self.current_filter = None
        self.selected_voyage_idx = None
        self.ctx_store = {}

        self.create_widgets()

        if init_row is not None:
            self.prepopulate_from_row(init_row)
        self.refresh_issue_table()

    # ---------------------------------------------------------------
    # Построение интерфейса
    # ---------------------------------------------------------------
    def create_widgets(self):
        # Секция «Выбранная запись»: только отображение рейсовых данных (без выбора)
        top = tk.LabelFrame(self.window, text=" Выбранная запись ", font=("Arial", 10, "bold"), padx=8, pady=6)
        top.pack(fill="x", padx=10, pady=5)

        # Строка 1: метка рейса + даты начала/окончания рейса
        self.lbl_selected_voyage = tk.Label(top, text="Рейс не выбран", font=("Arial", 10, "bold"),
                                            fg="#000000", bg="#add8e6", anchor="w", padx=5, pady=2)
        self.lbl_selected_voyage.pack(fill="x", padx=5, pady=2)
        # Строка 2 (объединение контекста и объёма данных)
        self.lbl_context = tk.Label(top, text="", font=("Arial", 9), anchor="w", justify="left",
                                    bg="#add8e6", fg="#000000", padx=5, pady=1)
        self.lbl_context.pack(fill="x", padx=5, pady=2)

        # --- Форма выдачи ---
        form = tk.LabelFrame(self.window, text=" Данные выдачи ", font=("Arial", 10, "bold"), padx=8, pady=6)
        form.pack(fill="x", padx=10, pady=5)

        self.inputs = {}
        # «выданные каталоги» — редактируемое поле: авто-заполняется значениями
        # из voyage_sub_data при выборе рейса, но его можно изменить вручную.
        rows = [
            ("Дата выдачи", "entry", None), ("ФИО сотрудника", "entry", None),
            ("Подразделение", "entry", None), ("Цель выдачи", "combo", ACTIVITY_OPTIONS),
            ("выданные каталоги", "entry", None),
            ("срок моратория/реализации", "entry", None),
            ("Научная задача", "entry", None),
            ("Отметка о реализации", "combo", REALIZATION_OPTIONS),
            ("Дата выполнения", "entry", None),
            ("Комментарий", "entry", None),
        ]
        for i, (field, kind, values) in enumerate(rows):
            r, c = divmod(i, 2)
            tk.Label(form, text=field + ":", font=("Arial", 10)) \
                .grid(row=r, column=c * 2, sticky="w", padx=(6, 2), pady=3)
            if kind == "combo":
                w = ttk.Combobox(form, values=values or [], width=42)
                w.set("")
            elif kind == "combo_yn":
                w = ttk.Combobox(form, values=["Да", "Нет"], width=42)
            else:
                w = tk.Entry(form, width=45)
            w.grid(row=r, column=c * 2 + 1, padx=(0, 10), pady=3, sticky="w")
            self.inputs[field] = w

        btns = tk.Frame(form)
        btns.grid(row=5, column=0, columnspan=4, pady=8)
        tk.Button(btns, text="➕ Добавить запись", bg="#2980b9", fg="white",
                  font=("Arial", 10, "bold"), padx=10, command=self.pick_voyage).pack(side="left", padx=4)
        tk.Button(btns, text="📄 Оформить выдачу", bg="#27ae60", fg="white",
                  font=("Arial", 10, "bold"), padx=10, command=self.add_issue).pack(side="left", padx=4)
        tk.Button(btns, text="💾 Сохранить изменения", bg="#d35400", fg="white",
                  font=("Arial", 10, "bold"), padx=10, command=self.update_issue).pack(side="left", padx=4)
        tk.Button(btns, text="❌ Удалить выдачу", bg="#c0392b", fg="white",
                  font=("Arial", 10, "bold"), padx=10, command=self.delete_issue).pack(side="left", padx=4)
        tk.Button(btns, text="🧹 Очистить форму", bg="#7f8c8d", fg="white",
                  font=("Arial", 10, "bold"), padx=10, command=self.clear_form).pack(side="left", padx=4)
        tk.Button(btns, text="🔍 Поиск", bg="#2980b9", fg="white",
                  font=("Arial", 10, "bold"), padx=10, command=self.set_filter).pack(side="left", padx=4)
        tk.Button(btns, text="🗑 Сбросить фильтр", bg="#7f8c8d", fg="white",
                  font=("Arial", 10, "bold"), padx=10, command=self.reset_filter).pack(side="left", padx=4)

        # --- Таблица выдач ---
        table_frame = tk.LabelFrame(self.window, text=" Журнал выдач ", font=("Arial", 10, "bold"))
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(table_frame, columns=DISPLAY_COLUMNS, show="headings")
        scx = tk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        scy = tk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=scx.set, yscrollcommand=scy.set)
        scx.pack(side="bottom", fill="x")
        scy.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        for col in DISPLAY_COLUMNS:
            self.tree.heading(col, text=TABLE_HEADINGS_SHORT.get(col, col))
            self.tree.column(col, width=120, anchor="center", stretch=False)
        self.col_widths = {}
        self.recompute_issue_widths()
        self.tree.bind("<<TreeviewSelect>>", self.on_issue_select)

        # --- Нижняя панель: отчёты + возврат/выход ---
        bottom = tk.Frame(self.window)
        bottom.pack(fill="x", padx=10, pady=6)
        tk.Button(bottom, text="📄 Экспорт (TXT)", bg="#16a085", fg="white",
                  font=("Arial", 10, "bold"), padx=8,
                  command=lambda: self.export_report("txt")).pack(side="left", padx=4)
        tk.Button(bottom, text="📊 Экспорт (XLSX)", bg="#16a085", fg="white",
                  font=("Arial", 10, "bold"), padx=8,
                  command=lambda: self.export_report("xlsx")).pack(side="left", padx=4)
        self.filter_status = tk.Label(bottom, text="", font=("Arial", 9, "bold"), fg="#2980b9")
        self.filter_status.pack(side="left", padx=15)
        tk.Button(bottom, text="◀ Вернуться в основной режим", bg="#7f8c8d", fg="white",
                  font=("Arial", 10, "bold"), padx=10, command=self.close_window).pack(side="right", padx=4)
        tk.Button(bottom, text="⏻ Выход из программы", bg="#c0392b", fg="white",
                  font=("Arial", 10, "bold"), padx=10, command=self.exit_program).pack(side="right", padx=4)

    def prepopulate_from_row(self, row):
        """Инициализирует контекст из заранее выбранного рейса (передан из main.py)."""
        self.set_context_from_dict(row)
        self._set_selected_label(row)

    def _set_selected_label(self, row):
        """Заполняет 1-ю строку секции: рейс + даты начала/окончания."""
        dates = []
        for k, label in (("Дата начала рейса", "Начало рейса"),
                         ("Дата окончания рейса", "Конец рейса")):
            v = fmt(row.get(k, ""))
            if v:
                dates.append(f"{label}: {v}")
        date_txt = " | ".join(dates) if dates else ""
        self.lbl_selected_voyage.config(
            text=f"{fmt(row.get('№ п/п'))} | {fmt(row.get('Рейс'))} | "
                 f"{fmt(row.get('Судно'))} | {fmt(row.get('Год'))}"
                 + (f"   —   {date_txt}" if date_txt else ""),
            fg="#000000")

    def set_context_from_dict(self, row):
        """Заполняет ctx_store и поле «выданные каталоги» по выбранному рейсу."""
        self.ctx_store = {k: row.get(k, "") for k in COPIED_FIELDS}
        self.ctx_store["UID_рейса"] = fmt(row.get("UID", ""))
        # Авто-заполняем «выданные каталоги» значениями связанных записей sub
        catalogs = self._collect_catalogs(self.ctx_store["UID_рейса"])
        cat_input = self.inputs.get("выданные каталоги")
        if cat_input is not None:
            cat_input.delete(0, tk.END)
            cat_input.insert(0, catalogs)
        self.lbl_context.config(text=self._fmt_record_details(self.ctx_store))

    def _fmt_record_details(self, data):
        """Формирует строку 2 секции «Выбранная запись»: контекст + объём данных."""
        parts = [f"{k}: {fmt(data.get(k, ''))}" for k in
                 ["Тип данных", "Прибор", "Район исследования", "Формат файла",
                  "Степень обработки", "Объем данных"]
                 if fmt(data.get(k, ""))]
        return "   |   ".join(parts) if parts else "Сведения о рейсе не заполнены."

    # ---------------------------------------------------------------
    # Выбор рейса
    # ---------------------------------------------------------------
    def pick_voyage(self):
        pick_cols = [
            "№ п/п", "Год", "Рейс", "Судно", "№ рейса", "Начальник рейса",
            "Район исследования", "Тип данных", "Прибор", "Формат файла",
            "Степень обработки", "Объем данных",
        ]
        win = tk.Toplevel(self.window)
        win.title("Выбор рейса")
        win.geometry("1300x600")
        win.transient(self.window)
        win.grab_set()
        search_var = tk.StringVar()

        def apply_search(_=None):
            q = search_var.get().strip().lower()
            for item in tree.get_children():
                tree.delete(item)
            n = 0
            for idx, row2 in self.voyage_df.iterrows():
                text = " ".join(fmt(row2[c]) for c in self.voyage_df.columns).lower()
                if q and q not in text:
                    continue
                tree.insert("", "end", iid=str(idx), values=[fmt(row2[c]) for c in pick_cols])
                n += 1
            lbl_count.config(text=f"Рейсов найдено: {n} из {len(self.voyage_df)}")

        bar = tk.Frame(win, padx=6, pady=6)
        bar.pack(fill="x")
        tk.Label(bar, text="Поиск по всем полям:", font=("Arial", 10)).pack(side="left")
        ent = tk.Entry(bar, textvariable=search_var, width=50, font=("Arial", 10))
        ent.pack(side="left", padx=6)
        ent.bind("<KeyRelease>", apply_search)
        tk.Button(bar, text="Выбрать", bg="#27ae60", fg="white", font=("Arial", 10, "bold"),
                  command=lambda: self._pick_selected(win, tree)).pack(side="left", padx=6)
        tk.Button(bar, text="Отмена", bg="#7f8c8d", fg="white", font=("Arial", 10),
                  command=win.destroy).pack(side="left", padx=6)
        lbl_count = tk.Label(bar, text="", font=("Arial", 9, "bold"))
        lbl_count.pack(side="left", padx=10)

        tf = tk.LabelFrame(win, text=" Список рейсов ", font=("Arial", 10))
        tf.pack(fill="both", expand=True, padx=6, pady=6)
        tree = ttk.Treeview(tf, columns=pick_cols, show="headings")
        for col in pick_cols:
            tree.heading(col, text=col)
            tree.column(col, width=110, anchor="center")
        ys = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        xs = tk.Scrollbar(tf, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        tree.pack(fill="both", expand=True)
        ys.pack(side="right", fill="y")
        xs.pack(side="bottom", fill="x")
        tree.bind("<Double-Button-1>", lambda _: self._pick_selected(win, tree))
        ent.focus_set()
        apply_search()

    def _pick_selected(self, win, tree):
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Внимание", "Выберите рейс в списке.", parent=self.window)
            return
        idx = int(sel[0])
        row = self.voyage_df.iloc[idx].to_dict()
        self.selected_voyage_idx = idx
        self.set_context_from_dict(row)
        self._set_selected_label(row)
        # После смены рейса очищаем форму выдачи и снимаем выбор: иначе «Сохранить»
        # перезапишет чужую запись контекстом нового рейса.
        self.clear_form()
        self.selected_issue_index = None
        win.destroy()

    # ---------------------------------------------------------------
    # Форма
    # ---------------------------------------------------------------
    def get_form_values(self):
        return {f: w.get().strip() for f, w in self.inputs.items()}

    def clear_form(self):
        for f, w in self.inputs.items():
            if isinstance(w, ttk.Combobox):
                w.set("")
            else:
                w.delete(0, tk.END)
        self.selected_issue_index = None

    # ---------------------------------------------------------------
    # CRUD
    # ---------------------------------------------------------------
    def _build_issue_row(self):
        if not self.ctx_store:
            raise RuntimeError("Сначала выберите рейс: «Выбрать рейс...».")
        vals = self.get_form_values()
        row = {k: self.ctx_store.get(k, "") for k in COPIED_FIELDS}
        row["UID_рейса"] = self.ctx_store.get("UID_рейса", "")
        # «выданные каталоги» берём из формы (авто-заполнены, но пользователь может изменить)
        row["выданные каталоги"] = vals.get("выданные каталоги", "")
        for k in ISSUE_FIELDS:
            if k in ("UID", "UID_рейса", "выданные каталоги"):
                continue
            row[k] = vals.get(k, "")
        if not row["Дата выдачи"]:
            raise RuntimeError("Укажите дату выдачи.")
        return row

    def _collect_catalogs(self, uid):
        """Собирает все значения поля «каталоги» записей voyage_sub_data, связанных с рейсом."""
        if not uid:
            return ""
        subs = self.sub_df[self.sub_df["UID_Родителя"].astype(str).str.strip() == str(uid).strip()]
        cats = [fmt(x) for x in subs["каталоги"] if fmt(x)]
        # подставляем только связанные с выбранной записью; без каталогов с копией
        return "\n".join(dict.fromkeys(cats))

    def save(self):
        while True:
            try:
                self.issue_df.to_excel(ISSUE_FILE, index=False)
                return True
            except PermissionError:
                if not messagebox.askretrycancel(
                        "Ошибка доступа",
                        f"Файл '{ISSUE_FILE}' открыт в Excel.\nЗакройте его и нажмите «Повторить».",
                        parent=self.window):
                    return False

    def add_issue(self):
        try:
            row = self._build_issue_row()
        except RuntimeError as e:
            messagebox.showwarning("Внимание", str(e), parent=self.window)
            return
        row["UID"] = f"{row['UID_рейса']}__iss_{len(self.issue_df)}"
        self.issue_df = pd.concat([self.issue_df, pd.DataFrame([row])], ignore_index=True)
        if not self.save():
            return
        self.refresh_issue_table()
        self.clear_form()
        messagebox.showinfo("Готово", "Выдача оформлена и внесена в журнал.", parent=self.window)

    def update_issue(self):
        """Сохранить: если строка журнала выделена — редактируем её,
        если нет — добавляем новую запись из текущей формы."""
        if self.selected_issue_index is None:
            self.add_issue()  # при предзаполненной форме (рейс выбран) — добавляем новую
            return
        try:
            row = self._build_issue_row()
        except RuntimeError as e:
            messagebox.showwarning("Внимание", str(e), parent=self.window)
            return
        old_uid = self.issue_df.at[self.selected_issue_index, "UID"]
        old_rid = self.issue_df.at[self.selected_issue_index, "UID_рейса"]
        for k, v in row.items():
            self.issue_df.at[self.selected_issue_index, k] = v
        self.issue_df.at[self.selected_issue_index, "UID"] = old_uid
        self.issue_df.at[self.selected_issue_index, "UID_рейса"] = old_rid
        if not self.save():
            return
        self.refresh_issue_table()
        self.selected_issue_index = None
        messagebox.showinfo("Готово", "Изменения сохранены.", parent=self.window)

    def delete_issue(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Внимание", "Выберите выдачу в таблице для удаления.",
                                   parent=self.window)
            return
        idx = int(sel[0])
        if messagebox.askyesno("Подтверждение", "Удалить выбранную запись о выдаче?",
                               parent=self.window):
            self.issue_df = self.issue_df.drop(idx).reset_index(drop=True)
            if not self.save():
                return
            self.refresh_issue_table()
            self.clear_form()
            messagebox.showinfo("Успех", "Запись удалена.", parent=self.window)

    # ---------------------------------------------------------------
    # Таблица, поиск, отчёт
    # ---------------------------------------------------------------
    def _display_df(self):
        if not self.current_filter:
            return self.issue_df
        df = self.issue_df.copy()
        for field, value in self.current_filter.items():
            # regex=False: значение — обычная подстрока, а не регулярное выражение,
            # чтобы спецсимволы (\P, (, [, .) не ломали поиск.
            df = df[df[field].astype(str).str.contains(value, case=False, na=False, regex=False)]
        return df

    def recompute_issue_widths(self):
        """Автоподбор ширины колонок журнала по содержимому и заголовку."""
        tk_font = tkfont.Font(font="TkDefaultFont")
        widths = {}
        for col in DISPLAY_COLUMNS:
            header = TABLE_HEADINGS_SHORT.get(col, col)
            max_len = tk_font.measure(header) + 25
            if col in self.issue_df.columns:
                mx = max((tk_font.measure(fmt(v)) for v in self.issue_df[col]), default=0)
                max_len = max(max_len, mx + 15)
            widths[col] = min(max(max_len, 55), 260)
        self.col_widths = widths

    def refresh_issue_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.recompute_issue_widths()
        df = self._display_df()
        for idx, row in df.iterrows():
            self.tree.insert("", "end", iid=str(idx), values=[fmt(row[c]) for c in DISPLAY_COLUMNS])
        for col, w in self.col_widths.items():
            self.tree.column(col, width=w, anchor="center", stretch=False)
        total = len(self.issue_df)
        shown = len(df)
        if self.current_filter:
            self.filter_status.config(text=f"Фильтр активен: {shown} из {total}", fg="#c0392b")
        else:
            self.filter_status.config(text=f"Всего выдач: {total}", fg="#2980b9")

    def on_issue_select(self, _):
        sel = self.tree.selection()
        if not sel:
            self.selected_issue_index = None
            return
        self.selected_issue_index = int(sel[0])
        row = self.issue_df.loc[self.selected_issue_index].to_dict()
        self.ctx_store = {k: row.get(k, "") for k in COPIED_FIELDS}
        self.ctx_store["UID_рейса"] = fmt(row.get("UID_рейса", ""))
        self.lbl_context.config(text=self._fmt_record_details(self.ctx_store))
        self._set_selected_label(self.ctx_store)
        for f, w in self.inputs.items():
            if isinstance(w, ttk.Combobox):
                w.set(fmt(row.get(f, "")))
            else:
                w.delete(0, tk.END)
                w.insert(0, fmt(row.get(f, "")))

    def set_filter(self):
        """Задаёт фильтр по заполненным полям формы выдачи."""
        crit = {f: v for f, v in self.get_form_values().items() if v}
        if not crit:
            messagebox.showinfo("Фильтр", "Заполните хотя бы одно поле формы для поиска.",
                                parent=self.window)
            return
        self.current_filter = crit
        self.refresh_issue_table()

    def reset_filter(self):
        self.current_filter = None
        # Сбрасываем и форму, и контекст рейса, чтобы не оставался «фантомный»
        # контекст, из-за которого «Оформить выдачу» создаст запись старого рейса.
        self.clear_form()
        self.ctx_store = {}
        self.lbl_selected_voyage.config(text="Рейс не выбран")
        self.lbl_context.config(text="")
        self.refresh_issue_table()

    def export_report(self, kind):
        df = self._display_df().copy()
        if df.empty:
            messagebox.showwarning("Экспорт", "Нет данных для экспорта.", parent=self.window)
            return
        if kind == "txt":
            path = filedialog.asksaveasfilename(
                defaultextension=".txt", filetypes=[("Текстовый файл", "*.txt")],
                title="Сохранить отчет (TXT)", parent=self.window)
            if not path:
                return
            with open(path, "w", encoding="utf-8") as f:
                f.write("=" * 70 + "\n")
                f.write("ЖУРНАЛ ВЫДАЧИ ДАННЫХ\n")
                f.write(f"Дата генерации: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
                f.write(f"Всего записей: {len(df)}\n")
                f.write("=" * 70 + "\n\n")
                for idx, row in df.iterrows():
                    f.write(f"--- ВЫДАЧА № {idx + 1} ---\n")
                    for c in ISSUE_ROW_FIELDS:
                        f.write(f"{c}: {fmt(row.get(c, ''))}\n")
                    f.write("\n" + "." * 40 + "\n\n")
        else:
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel файл", "*.xlsx")],
                title="Сохранить отчет (XLSX)", parent=self.window)
            if not path:
                return
            out = df[[c for c in ISSUE_ROW_FIELDS if c in df.columns]].copy()
            try:
                out.to_excel(path, index=False)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить XLSX:\n{e}", parent=self.window)
                return
        messagebox.showinfo("Успех", f"Отчет сохранён: {os.path.basename(path)}", parent=self.window)

    # ---------------------------------------------------------------
    # Возврат / выход
    # ---------------------------------------------------------------
    def close_window(self):
        """Возврат в основной режим (закрыть окно журнала)."""
        self.save_quiet()
        self.window.destroy()

    def exit_program(self):
        """Полный выход из программы."""
        if messagebox.askyesno("Выход", "Завершить работу программы?", parent=self.window):
            self.save_quiet()
            self.window.destroy()
            # Если передан колбэк выхода (например, on_closing из main.py,
            # делающий финальный бэкап основной БД и очистку дубликатов) — используем его,
            # иначе закрываем root напрямую.
            if self.on_exit is not None:
                try:
                    self.on_exit()
                except Exception:
                    pass
            else:
                try:
                    self.root.destroy()
                except Exception:
                    pass

    def save_quiet(self):
        """Фоновый бэкап при закрытии окна журнала."""
        try:
            make_backup()
        except Exception:
            pass

    @property
    def root_win(self):
        """Для совместимости."""
        return self.window


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    win = IssueManagerWindow(root)
    root.mainloop()