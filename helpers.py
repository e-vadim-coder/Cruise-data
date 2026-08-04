# -*- coding: utf-8 -*-
# helpers.py
import os
import tkinter as tk
from tkinter import ttk
import pandas as pd
import zipfile
import glob
from datetime import datetime
import hashlib

FIELDS_CONFIG = {
    "№ п/п": int,
    "Год": int,
    "Рейс": str,
    "Судно": str,
    "№ рейса": int,
    "Дата начала рейса": str,
    "Дата окончания рейса": str,
    "Начальник рейса": str,
    "№ этапа рейса": str,
    "№ диска": int,
    "Инвентарный № диска": int,
    "Начальник отряда геофизики": str,
    "Начальник отряда сейсмических исследований": str,
    "Инициатор научной задачи": str,
    "Район исследования": str,
    "Тип данных": str,
    "Прибор": str,
    "Степень обработки": str,
    "Формат файла": str,
    "Объем данных": str,
}


def create_backup(file_path: str, backup_dir: str = "backups", max_backups: int = 40) -> bool:
    """
    Создает сжатую zip-копию файла данных, сохраняет в указанную директорию
    и удаляет старые копии, оставляя только N последних.
    """
    # Особое условие проекта: Обязательная проверка типов передаваемых параметров
    if not isinstance(file_path, str):
        raise TypeError("Параметр file_path должен быть строкой.")
    if not isinstance(backup_dir, str):
        raise TypeError("Параметр backup_dir должен быть строкой.")
    if not isinstance(max_backups, int):
        raise TypeError("Параметр max_backups должен быть целым числом.")

    # Если исходный файл данных еще не создан (например, самый первый запуск), выходим
    if not os.path.exists(file_path):
        print(f"Исходный файл {file_path} не найден. Резервная копия не создана.")
        return False

    try:
        # Создаем папку для бэкапов, если её нет
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)

        # Формируем имя архива: voyage_data_20261025_143005.zip
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        zip_filename = f"{base_name}_{timestamp}.zip"
        zip_filepath = os.path.join(backup_dir, zip_filename)

        # Сжимаем файл в ZIP-архив
        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # os.path.basename(file_path) нужен, чтобы внутри zip файл лежал без путей
            zipf.write(file_path, arcname=os.path.basename(file_path))

        print(f"Резервная копия успешно создана: {zip_filepath}")

        # Очистка старых копий: ищем все .zip файлы шаблона 'имя_*.zip' в папке бэкапов
        search_pattern = os.path.join(backup_dir, f"{base_name}_*.zip")
        backup_files = glob.glob(search_pattern)

        # Сортируем файлы по времени изменения (mtime) от самых старых к самым новым
        backup_files.sort(key=os.path.getmtime)

        # Если копий больше, чем разрешено, удаляем самые старые
        if len(backup_files) > max_backups:
            files_to_delete = backup_files[:-max_backups]  # Все, кроме последних N
            for old_file in files_to_delete:
                try:
                    os.remove(old_file)
                    print(f"Старая резервная копия удалена: {old_file}")
                except Exception as e:
                    print(f"Не удалось удалить старый файл {old_file}: {e}")

        return True

    except Exception as e:
        print(f"Ошибка при создании резервной копии: {e}")
        return False


def validate_type(value: str, expected_type: type) -> bool:
    """Проверяет соответствие строкового значения ожидаемому типу данных."""
    if not isinstance(expected_type, type) or not isinstance(value, str):
        raise TypeError("Неверные типы параметров в процедуре validate_type.")
    if value.strip() == "":
        return True
    if expected_type is int:
        return value.strip().isdigit()
    return True


def load_data(file_name: str) -> pd.DataFrame:
    """Загружает данные из Excel файла."""
    if not isinstance(file_name, str):
        raise TypeError("Параметр file_name должен быть строкой.")
    if os.path.exists(file_name):
        try:
            df = pd.read_excel(file_name)
            for col, t in FIELDS_CONFIG.items():
                if col in df.columns:
                    if t is int:
                        df[col] = df[col].fillna(0).astype(float).astype(int)
                    else:
                        df[col] = df[col].fillna("").astype(str)
            return df
        except Exception:
            return pd.DataFrame(columns=list(FIELDS_CONFIG.keys()))
    return pd.DataFrame(columns=list(FIELDS_CONFIG.keys()))


def save_data(df: pd.DataFrame, file_name: str) -> None:
    """Сохраняет DataFrame в Excel с защитой от блокировки файла."""
    if not isinstance(df, pd.DataFrame) or not isinstance(file_name, str):
        raise TypeError("Неверные типы параметров в процедуре save_data.")
    from tkinter import messagebox

    while True:
        try:
            df.to_excel(file_name, index=False)
            break
        except PermissionError:
            response = messagebox.askretrycancel(
                "Ошибка доступа",
                f"Файл '{file_name}' открыт в Excel.\n\nПожалуйста, закройте его и нажмите 'Повторить'.",
            )
            if not response:
                break


def reset_form_fields(inputs_dict: dict) -> None:
    """Очищает все виджеты ввода в переданном словаре."""
    if not isinstance(inputs_dict, dict):
        raise TypeError("Параметр inputs_dict должен быть словарем.")
    for widget in inputs_dict.values():
        if isinstance(widget, ttk.Combobox):
            widget.set("")
        elif isinstance(widget, tk.Entry):
            widget.delete(0, tk.END)


def delete_row_from_df(df: pd.DataFrame, index_to_delete: int) -> pd.DataFrame:
    """Удаляет строку из DataFrame по индексу."""
    if not isinstance(df, pd.DataFrame) or not isinstance(index_to_delete, int):
        raise TypeError("Неверные типы параметров.")
    if index_to_delete in df.index:
        df = df.drop(index_to_delete).reset_index(drop=True)
    return df


def is_exact_match(new_data: dict, template_row: pd.Series) -> bool:
    """Проверяет, совпадает ли измененная форма полностью с шаблоном."""
    if not isinstance(new_data, dict) or not isinstance(
        template_row, pd.Series
    ):
        raise TypeError("Неверные типы параметров.")
    for field, value in new_data.items():
        template_val = (
            str(int(template_row[field]))
            if FIELDS_CONFIG[field] is int and pd.notna(template_row[field])
            else str(template_row[field])
        )
        if template_val == "nan" or template_val == "0":
            template_val = ""
        if value.strip() != template_val.strip():
            return False
    return True


def export_to_txt(df: pd.DataFrame, target_file_path: str) -> None:
    """Генерирует текстовый отчет на основе переданного DataFrame."""
    if not isinstance(df, pd.DataFrame) or not isinstance(
        target_file_path, str
    ):
        raise TypeError("Неверные типы параметров в процедуре export_to_txt.")
    try:
        with open(target_file_path, "w", encoding="utf-8") as f:
            f.write("=" * 60 + "\n")
            f.write("ОТЧЕТ ПО РЕЙСОВЫМ ДАННЫМ ГЕОФИЗИКИ\n")
            f.write(
                f"Дата генерации: {pd.Timestamp.now().strftime('%d.%m.%Y %H:%M:%S')}\n"
            )
            f.write(f"Всего записей в отчете: {len(df)}\n")
            f.write("=" * 60 + "\n\n")
            for idx, row in df.iterrows():
                f.write(f"--- ЗАПИСЬ № {idx + 1} ---\n")
                for col in df.columns:
                    val = (
                        ""
                        if pd.isna(row[col])
                        else str(row[col]).replace(".0", "")
                    )
                    f.write(f"{col}: {val}\n")
                f.write("\n" + "." * 40 + "\n\n")
    except Exception as e:
        from tkinter import messagebox

        messagebox.showerror("Ошибка", f"Не удалось записать отчет:\n{str(e)}")

def get_card_context(df, current_index: int) -> dict:  # (0)
    """Определяет данные для 3-х карточек режима «Детальный просмотр»."""  # (4)
    import pandas as pd  # (4)
    if not isinstance(df, pd.DataFrame):  # (4)
        raise TypeError("Параметр 'df' должен быть объектом класса pandas.DataFrame")  # (8)
    if not isinstance(current_index, int):  # (4)
        raise TypeError("Параметр 'current_index' должен быть целым числом (int)")  # (8)

    total_rows = len(df)  # (4)

    if total_rows == 0:  # (4)
        return {  # (8)
            "left_data": "START", "center_data": {}, "right_data": "END",  # (12)
            "has_prev": False, "has_next": False,  # (12)
            "current_pos_text": "База данных пуста"  # (12)
        }  # (8)

    if current_index < 0 or current_index >= total_rows:  # (4)
        raise IndexError(f"Индекс {current_index} выходит за пределы таблицы (всего строк: {total_rows})")  # (8)

    if current_index == 0:  # (4)
        left_data = "START"  # (8)
    else:  # (4)
        left_data = df.iloc[current_index - 1].to_dict()  # (8)

    center_data = df.iloc[current_index].to_dict()  # (4)

    if current_index == total_rows - 1:  # (4)
        right_data = "END"  # (8)
    else:  # (4)
        right_data = df.iloc[current_index + 1].to_dict()  # (8)

    context = {  # (4)
        "left_data": left_data,  # (8)
        "center_data": center_data,  # (8)
        "right_data": right_data,  # (8)
        "has_prev": current_index > 0,  # (8)
        "has_next": current_index < total_rows - 1,  # (8)
        "current_pos_text": f"Запись {current_index + 1} из {total_rows}"  # (8)
    }  # (4)

    return context  # (4)

import hashlib

def get_file_hash(file_path: str) -> str:
    """Вычисляет уникальную хэш-сумму самого файла Excel для сравнения."""
    if not os.path.exists(file_path):
        return ""
    hasher = hashlib.md5()
    try:
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hasher.update(chunk)
        return hasher.hexdigest()
    except Exception:
        return ""
