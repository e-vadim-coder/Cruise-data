# -*- coding: utf-8 -*-
# main.py

import tkinter as tk  # (0)
from tkinter import ttk, messagebox, filedialog  # (0)
import os  # (0)
import pandas as pd  # (0)
from config import FIELDS_CONFIG, TABS_CONFIG, DICTIONARIES  # (0)
from voyage_app_tabs import VoyageAppTabs  # (0)
import helpers  # (0)


class MainApplication:  # (0)
    def __init__(self, root):  # (4)
        self.root = root  # (8)
        self.root.title("Система управления данными рейсов")                             # (8)
        scr_width = self.root.winfo_screenwidth()                                        # (8)
        scr_height = self.root.winfo_screenheight()                                      # (8)
        self.root.geometry(f"{scr_width}x{scr_height}+0+0")
        self.root.minsize(1000, 500)

        self.db_file = "voyage_data.xlsx"  # (8)

        self.start_backup_path = helpers.create_zip_backup(self.db_file, "start")  # (8)
        self.start_hash = helpers.get_file_hash(self.db_file)  # (8)

        self.app_tabs = VoyageAppTabs(self.root)  # (8)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close_app)  # (8)

    def on_close_app(self):  # (4)
        final_hash = helpers.get_file_hash(self.db_file)  # (8)
        final_backup = helpers.create_zip_backup(self.db_file, "finish")  # (8)

        if self.start_hash == final_hash and final_backup and os.path.exists(final_backup):  # (8)
            try:  # (12)
                os.remove(final_backup)  # (16)
            except Exception:  # (12)
                pass  # (16)

        helpers.clean_old_backups(10)  # (8)
        self.root.destroy()  # (8)


class DetailViewWindow:  # (0)
    def __init__(self, parent, app, current_index):  # (4)
        self.parent = parent  # (8)
        self.app = app  # (8)
        self.current_index = current_index  # (8)
        self.center_fields = {}  # (8)

        self.window = tk.Toplevel(parent)  # (8)
        self.window.title("Карточки детального просмотра")  # (8)
        self.window.geometry("1200x650")  # (8)
        self.window.grab_set()  # (8)

        self.create_navigation_bar()  # (8)
        self.create_cards_layout()  # (8)
        self.refresh_cards()  # (8)

    def create_navigation_bar(self):  # (4)
        nav = tk.Frame(self.window, bg="#eaeaea", pady=5)  # (8)
        nav.pack(fill=tk.X, side=tk.TOP)  # (8)

        tk.Button(nav, text="⋘ В начало", command=lambda: self.navigate("first")).pack(side=tk.LEFT, padx=5)  # (8)
        tk.Button(nav, text="◁ Назад", command=lambda: self.navigate("prev")).pack(side=tk.LEFT, padx=5)  # (8)
        self.lbl_pos = tk.Label(nav, text="", bg="#eaeaea", font=("Arial", 10, "bold"))  # (8)
        self.lbl_pos.pack(side=tk.LEFT, padx=20)  # (8)
        tk.Button(nav, text="Вперед ▷", command=lambda: self.navigate("next")).pack(side=tk.LEFT, padx=5)  # (8)
        tk.Button(nav, text="В конец ⋙", command=lambda: self.navigate("last")).pack(side=tk.LEFT, padx=5)  # (8)
        tk.Button(nav, text="Экспорт в TXT", command=self.export_current, bg="#009688", fg="white").pack(side=tk.RIGHT, padx=5)  # (8)
        tk.Button(nav, text="Сохранить центр", command=self.save_center, bg="#2196f3", fg="white").pack(side=tk.RIGHT, padx=5)  # (8)

    def create_cards_layout(self):  # (4)
        box = tk.Frame(self.window, bg="#eaeaea")  # (8)
        box.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)  # (8)
        box.columnconfigure(0, weight=1)  # (8)
        box.columnconfigure(1, weight=1)  # (8)
        box.columnconfigure(2, weight=1)  # (8)
        box.rowconfigure(0, weight=1)  # (8)

        self.f_left = tk.LabelFrame(box, text=" Предыдущая запись ", bg="#f0f0f0", fg="#888888", bd=1,
                                    relief=tk.FLAT)  # (8)
        self.f_center = tk.LabelFrame(box, text=" ТЕКУЩАЯ ЗАПИСЬ (Редактирование) ", bg="#ffffff", fg="#1a73e8", bd=2,
                                      relief=tk.SOLID)  # (8)
        self.f_right = tk.LabelFrame(box, text=" Следующая запись ", bg="#f0f0f0", fg="#888888", bd=1,
                                     relief=tk.FLAT)  # (8)

        self.f_left.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)  # (8)
        self.f_center.grid(row=0, column=1, sticky="nsew", padx=4, pady=4)  # (8)
        self.f_right.grid(row=0, column=2, sticky="nsew", padx=4, pady=4)  # (8)

    def build_symmetric_grid(self, parent_frame, data_row, is_editable, text_color, bg_color):  # (4)
        for w in parent_frame.winfo_children():  # (8)
            w.destroy()  # (12)

        canvas = tk.Canvas(parent_frame, highlightthickness=0, bg=bg_color)  # (8)
        sb = tk.Scrollbar(parent_frame, orient="vertical", command=canvas.yview)  # (8)
        sf = tk.Frame(canvas, bg=bg_color)  # (8)

        def _on_wheel(e):  # (8)
            try:
                canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")  # (12)
            except tk.TclError:
                pass  # (12)

        canvas.bind("<MouseWheel>", _on_wheel)  # (8)

        canvas.create_window((0, 0), window=sf, anchor="nw")  # (8)
        canvas.configure(yscrollcommand=sb.set)  # (8)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)  # (8)
        sb.pack(side=tk.RIGHT, fill=tk.Y)  # (8)

        row_idx = 0  # (8)
        for tab_name, keys in TABS_CONFIG.items():  # (8)
            g_lbl = tk.Label(sf, text=f"[{tab_name.upper()}]", font=("Arial", 9, "bold"), fg=text_color, bg=bg_color)  # (12)
            g_lbl.grid(row=row_idx, column=0, columnspan=2, sticky="w", pady=(8, 2), padx=5)  # (12)
            row_idx += 1  # (12)

            for key in keys:  # (12)
                info = FIELDS_CONFIG[key]  # (16)
                lbl = tk.Label(sf, text=info["short"] + ":", font=("Arial", 9), fg=text_color, bg=bg_color,
                               anchor="w")  # (16)
                lbl.grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)  # (16)

                val = "" if data_row is None else data_row.get(key, "")  # (16)
                if pd.isna(val):
                    val = ""  # (16)
                elif info["type"] is int and val != "":
                    val = str(int(val))  # (16)

                if is_editable:  # (16)
                    if info["short"] in DICTIONARIES:  # (20)
                        ent = ttk.Combobox(sf, values=DICTIONARIES[info["short"]])  # (24)
                        ent.set(str(val))  # (24)
                    else:  # (20)
                        ent = tk.Entry(sf, bd=1, relief=tk.SOLID)  # (24)
                        ent.insert(0, str(val))  # (24)
                    ent.grid(row=row_idx, column=1, sticky="ew", padx=5, pady=2)  # (20)
                    self.center_fields[key] = ent  # (20)
                else:  # (16)
                    ent = tk.Entry(sf, bd=1, relief=tk.FLAT, fg=text_color, bg=bg_color, state="disabled")  # (20)
                    ent.config(state="normal")  # (20)
                    ent.insert(0, str(val))  # (20)
                    ent.config(state="disabled")  # (20)
                    ent.grid(row=row_idx, column=1, sticky="ew", padx=5, pady=2)  # (20)

                sf.columnconfigure(1, weight=1)  # (16)
                row_idx += 1  # (16)

        sf.update_idletasks()  # (8)
        canvas.config(scrollregion=canvas.bbox("all"))  # (8)

    def refresh_cards(self):  # (4)
        total = len(self.app.df)  # (8)
        self.lbl_pos.config(text=f"Запись {self.current_index + 1} из {total}")  # (8)

        left_row = self.app.df.iloc[self.current_index - 1].to_dict() if self.current_index > 0 else None  # (8)
        center_row = self.app.df.iloc[self.current_index].to_dict()  # (8)
        right_row = self.app.df.iloc[
            self.current_index + 1].to_dict() if self.current_index < total - 1 else None  # (8)

        self.build_symmetric_grid(self.f_left, left_row, False, "#a0a0a0", "#f0f0f0")  # (8)
        self.build_symmetric_grid(self.f_center, center_row, True, "#000000", "#ffffff")  # (8)
        self.build_symmetric_grid(self.f_right, right_row, False, "#a0a0a0", "#f0f0f0")  # (8)

    def navigate(self, direction):  # (4)
        total = len(self.app.df)  # (8)
        if direction == "first":
            self.current_index = 0  # (8)
        elif direction == "last":
            self.current_index = total - 1  # (8)
        elif direction == "prev" and self.current_index > 0:
            self.current_index -= 1  # (8)
        elif direction == "next" and self.current_index < total - 1:
            self.current_index += 1  # (8)
        self.refresh_cards()  # (8)

    def save_center(self):  # (4)
        idx = self.app.df.index[self.current_index]  # (8)
        for key, widget in self.center_fields.items():  # (8)
            val = widget.get().strip()  # (12)
            if val == "":  # (12)
                self.app.df.at[idx, key] = pd.NA  # (16)
            elif FIELDS_CONFIG[key]["type"] is int:  # (12)
                try:
                    self.app.df.at[idx, key] = int(val)  # (16)
                except ValueError:
                    pass  # (16)
            else:  # (12)
                self.app.df.at[idx, key] = val  # (16)
        self.app.save_df_to_disk()  # (8)
        self.app.refresh_table()  # (8)
        self.refresh_cards()  # (8)

    def export_current(self):  # (4)
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt")])  # (8)
        if not path: return  # (8)
        row_dict = self.app.df.iloc[self.current_index].to_dict()  # (8)
        if helpers.export_to_txt_report(path, row_dict):  # (8)
            messagebox.showinfo("Успех", "Отчет успешно сохранен!")  # (12)


if __name__ == "__main__":  # (0)
    root = tk.Tk()  # (4)
    app = MainApplication(root)  # (4)
    root.mainloop()