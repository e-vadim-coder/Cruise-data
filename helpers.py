# -*- coding: utf-8 -*-
# helpers.py
import os
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import zipfile
import glob
from datetime import datetime
import hashlib
import time

MAX_BACKUPS = 50
CLEAN_INTERVAL_DAYS = 10
LAST_CLEAN_FILE = os.path.join("backups", ".last_clean")

DICTIONARIES = {  # (0)
    "Судно": [
        "Академик Николай Страхов",
        "Академик Борис Петров",
        "Академик Мстислав Келдыш",
        "Академик Иоффе",
        "Академик Сергей Вавилов",
        "Профессор Штокман"
    ],  # (4)
    "Начальник рейса": [
        "Сивков В.В.",
        "Щука С.А.",
        "Дорохов Д.В.",
        "Кречик В.А.",
        "Крек А.В.",
        "Пономаренко Е.П.",
        "Бубнова Е.С.",
        "Данченков А.Р.",
        "Дорохова Е.В.",
        "Фрей Д.И.",
        "Баширова Л.Д.",
        "Ульянова М.О."
    ],  # (4)
    "Начальник отряда геофизики": [  # (4)
        "Дорохов Д.В.",  # (8)
        "Ежов В.Е.",  # (8)
        "Данченков А.Р.", # (8)
        "Дорохова Е.В.",
        "Дудков И.Ю.",
        "Кречик В.А.",
        "Пономаренко Е.П.",
        "Сергеев А.Ю.",
        "Луговой Н.Н.",
        "Матуль А.Г.",
        "Сухих Е.А."
    ],  # (4)
    "Начальник отряда сейсмических исследований": [  # (4)
        "Ежов В.Е.",  # (8)
        "Ананьев Р.А."  # (8)
    ],  # (4)
    "Инициатор научной задачи": ["Институт океанологии", "МГУ", "РАН"],  # (4)
    "Район исследования": ["Баренцево море", "SEB", "FZ", "Балтийское море", "Атлантика"],  # (4)
    "Тип данных": ["ОЛЭ", "МЛЭ", "Профилограф", "ГЛБО", "ADCP", "SADCP", "Сейсмокоса"],  # (4)
    "Прибор": [
        "Kongsberg Simrad EA400",
        "Kongsberg Simrad EA600",
        "Reson SeaBat T50",
        "Reson SeaBat 8111",
        "Reson SeaBat 7150",
        "EdgeTech 3300-HM",
        "Parasound P70",
        "Benthos C3D",
        "SES2000",
        "SyQuest Bathy-2010 "
    ],  # (4)
    "Степень обработки": ["Необработанные", "Обработанные", "Сырые и обработанные"],  # (4)
    "Формат файла": [
        "ASD, SEG-Y",
        "JSF, SEG-Y",
        "PDS2000 project",
        "GIS project",
        "Hypack project",
        "Qinsy project",
        "ASCII",
        "DTM",
        "GeoTiff",
        "SEG-Y",
        "RAW",
        "SES2000"
    ],  # (4)
}  # (0)

TABS_CONFIG = {  # (0)
    "Общая информация": [  # (4)
        "№ п/п",  # (8)
        "Год",  # (8)
        "Рейс",  # (8)
        "Судно",  # (8)
        "№ рейса",  # (8)
        "Начальник рейса",  # (8)
        "Дата начала рейса",  # (8)
        "Дата окончания рейса",  # (8)
    ],  # (4)
    "Геофизика и приборы": [  # (4)
        "№ этапа рейса",  # (8)
        "Начальник отряда геофизики",  # (8)
        "Начальник отряда сейсмических исследований",
        "Инициатор научной задачи",  # (8)
        "Район исследования",  # (8)
        "Тип данных",  # (8)
        "Прибор",  # (8)
    ],  # (4)
    "Носители и файлы": [  # (4)
        "№ диска",  # (8)
        "Инвентарный № диска",  # (8)
        "Степень обработки",  # (8)
        "Формат файла",  # (8)
        "Объем данных",  # (8)
    ],  # (4)
}  # (0)

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
    "UID": str
}

# Словарь сокращений только для шапки Treeview
TABLE_HEADINGS_SHORT = {
    "№ этапа рейса": "этап",
    "№ диска": "диск",
    "Дата начала рейса": "Дата нач.рейса",
    "Дата окончания рейса": "Дата оконч.рейса",
    "Начальник рейса": "Нач.рейса",
    "Инвентарный № диска": "Инв.№ диска",
    "Начальник отряда геофизики": "Нач.отр.геофиз.",
    "Начальник отряда сейсмических исследований": "Нач.отр.сейсм.иссл-ий",
    "Инициатор научной задачи": "Инициатор научн.зад.",
    "Район исследования": "Район иссл-ния",
    # Для остальных полей, которых нет в словаре, программа автоматически оставит полное название
}

def validate_type(value: str, expected_type: type) -> bool:
    """Проверяет соответствие строкового значения ожидаемому типу данных."""
    if not isinstance(expected_type, type) or not isinstance(value, str):
        raise TypeError("Неверные типы параметров в процедуре validate_type.")
    if value.strip() == "":
        return True
    if expected_type is int:
        return value.strip().isdigit()
    return True


def apply_field_type(value, expected_type):
    """Приводит строковое значение формы к типу поля (единое правило).
    int: пустая или нечисловая строка -> 0, иначе int(value).
    str: возвращается без изменений.
    """
    value = str(value).strip()
    if expected_type is int:
        s = value[1:] if value.startswith("-") else value
        return int(value) if s.isdigit() else 0
    return value


def calculate_safe_row_uid(row: dict) -> str:
    """
    Отдельная независимая функция.
    Безопасно вычисляет UID: если научные поля пустые, возвращает маркер 'empty_'.
    Если поля заполнены — вызывает вашу родную генерацию хэш-суммы.
    """
    pp = str(row.get("№ п/п", "")).strip()
    vessel = str(row.get("Судно", "")).strip()
    voyage = str(row.get("№ рейса", "")).strip()
    year = str(row.get("Год", "")).strip()

    # Если ключевые поля пустые, не считаем кривой хэш-сироту
    if not vessel or not voyage or not year or vessel == "nan" or year == "nan":
        return f"empty_{pp}"

    # Если поля заполнены — вызываем оригинальный метод хэш-суммы
    return generate_row_uid(row)

def load_data(file_name: str) -> pd.DataFrame:
    """Загружает данные из Excel файла."""
    if not isinstance(file_name, str):
        raise TypeError("Параметр file_name должен быть строкой.")
    if os.path.exists(file_name):
        try:
            df = pd.read_excel(file_name)
            # --- УНИВЕРСАЛЬНОЕ ПРИВЕДЕНИЕ ТИПОВ ДЛЯ ОБОИХ ФАЙЛОВ EXCEL ---
            # Автоматически выбираем нужную конфигурацию полей
            current_config = SUB_TABLE_FIELDS if "UID_Родителя" in df.columns else FIELDS_CONFIG
            for col, t in current_config.items():
                if col in df.columns:
                    if t is int:
                        # Принудительно очищаем от NaN и переводим в целые числа
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
                    else:
                        # Намертво глушим float64 в текстовых колонках, превращая всё в чистые строки
                        df[col] = df[col].fillna("").astype(str).str.strip()
            # -------------------------------------------------------------
            # --- БЕЗОПАСНЫЙ И УМНЫЙ ПЕРЕСЧЕТ КЛЮЧЕЙ UID ---
            # Проверяем, какая именно таблица сейчас загружается программой
            if "UID_Родителя" in df.columns:
                # Для подчиненной таблицы: ВСЕГДА жестко и принудительно пересчитываем UID.
                # Формула: берем UID_Родителя этой строки и прибавляем к нему её физический номер (idx)
                df["UID"] = [f"{str(row['UID_Родителя'])}__sub_{idx}" for idx, row in df.iterrows()]
            else:
                # Для основной таблицы рейсов: пересчитываем UID только для пустых или новых строк
                if "UID" not in df.columns:
                    df["UID"] = [calculate_safe_row_uid(row) for _, row in df.iterrows()]
                else:
                    df["UID"] = df["UID"].fillna("")
                    for idx, row in df.iterrows():
                        current_uid = str(row["UID"]).strip()
                        if not current_uid or current_uid in ("", "nan", "empty_"):
                            df.at[idx, "UID"] = calculate_safe_row_uid(row)
            # --- КОНЕЦ ИСПРАВЛЕНИЯ В HELPERS.PY ---

            return df
        except Exception:
            return pd.DataFrame(columns=list(FIELDS_CONFIG.keys()))
    return pd.DataFrame(columns=list(FIELDS_CONFIG.keys()))


def save_data(df: pd.DataFrame, file_name: str) -> None:
    """Сохраняет DataFrame в Excel. При ошибке доступа бросает PermissionError."""
    if not isinstance(df, pd.DataFrame) or not isinstance(file_name, str):
        raise TypeError("Неверные типы параметров в процедуре save_data.")
    df.to_excel(file_name, index=False)


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


def _to_int(v):
    """Приводит значение к int для сравнения (пусто/NaN -> 0)."""
    if pd.isna(v):
        return 0
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


def _to_str(v):
    """Приводит значение к строке для сравнения (пусто/NaN -> '')."""
    if pd.isna(v):
        return ""
    return str(v).strip()


def find_full_duplicate(df: pd.DataFrame, row_dict: dict, exclude_index: int = None):
    """Ищет в df строку, полностью совпадающую с row_dict (без учёта UID).
    Возвращает индекс найденной строки или None."""
    compare_cols = [c for c in df.columns if c != "UID"]
    for idx, row in df.iterrows():
        if exclude_index is not None and idx == exclude_index:
            continue
        same = True
        for c in compare_cols:
            if FIELDS_CONFIG.get(c) is int:
                a, b = _to_int(row[c]), _to_int(row_dict.get(c))
            else:
                a, b = _to_str(row[c]), _to_str(row_dict.get(c))
            if a != b:
                same = False
                break
        if same:
            return idx
    return None


def _fmt_report_value(v):
    """Форматирует значение ячейки для текстового отчета без искажений."""
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


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
                    val = _fmt_report_value(row[col])
                    f.write(f"{col}: {val}\n")
                f.write("\n" + "." * 40 + "\n\n")
    except Exception as e:
        # from tkinter import messagebox
        messagebox.showerror("Ошибка", f"Не удалось записать отчет:\n{str(e)}")

def get_card_context(df, current_index: int) -> dict:  # (0)
    """Определяет данные для 3-х карточек режима «Детальный просмотр»."""  # (4)

    # import pandas as pd  # (4)
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


def create_backup(file_paths, backup_dir="backups"):
    """
    Создает архив с простым именем 'backup_дата_время.zip' для всех файлов из списка.
    Если файлы не менялись, дубликат по хэшу не создается.
    """
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)

    existing_files = [f for f in file_paths if os.path.exists(f)]
    if not existing_files:
        return ""

    # Имя архива теперь простое и понятное, без имени таблиц
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    zip_name = os.path.join(backup_dir, f"backup_{timestamp}.zip")

    try:
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in existing_files:
                zipf.write(file_path, os.path.basename(file_path))
        return zip_name
    except Exception as e:
        print(f"Ошибка создания бэкапа: {e}")
        return ""


def clean_duplicate_backups(backup_dir="backups"):
    """
    Очистка папки бэкапов. Раз в 10 дней наводит порядок, удаляет дубликаты
    по хэшу и избыток по количеству, выводя отчет в консоль.
    """
    if not os.path.exists(backup_dir):
        return

    current_date = datetime.now()
    deleted_count = 0  # Счётчик удалённых файлов

    # ПРОВЕРКА МЕТКИ ПО ПОНЯТНОЙ ДАТЕ
    if os.path.exists(LAST_CLEAN_FILE):
        try:
            with open(LAST_CLEAN_FILE, 'r') as f:
                last_clean_str = f.read().strip()
            last_clean_date = datetime.strptime(last_clean_str, "%Y-%m-%d")

            days_passed = (current_date - last_clean_date).days
            if days_passed < CLEAN_INTERVAL_DAYS:
                return  # 10 дней еще не прошло, уходим
        except Exception:
            pass

    # ПОИСК И УДАЛЕНИЕ ДУБЛИКАТОВ ПО ХЭШУ
    zips = glob.glob(os.path.join(backup_dir, "*.zip"))
    zips.sort(key=os.path.getmtime)

    seen_hashes = set()
    unique_zips = []
    for old_zip in zips:
        h = get_file_hash(old_zip)
        if h in seen_hashes:
            try:
                os.remove(old_zip)
                deleted_count += 1  # Увеличиваем счётчик при удалении хэш-дубликата
            except Exception:
                pass
        else:
            seen_hashes.add(h)
            unique_zips.append(old_zip)

    # ОГРАНИЧЕНИЕ ПО КОЛИЧЕСТВУ (оставляем последние MAX_BACKUPS)
    if len(unique_zips) > MAX_BACKUPS:
        for old_zip in unique_zips[:-MAX_BACKUPS]:
            try:
                os.remove(old_zip)
                deleted_count += 1  # Увеличиваем счётчик при удалении лишнего по количеству
            except Exception:
                pass

    # ВЫВОД СООБЩЕНИЯ ОБ ОЧИСТКЕ
    if deleted_count > 0:
        print(f"[ОЧИСТКА] Обнаружен день плановой проверки. Удалено устаревших бэкапов: {deleted_count} шт.")
    else:
        print("[ОЧИСТКА] День плановой проверки: папка бэкапов уже в идеальном состоянии.")

    # ЗАПИСЬ СВЕЖЕЙ ПОНЯТНОЙ ДАТЫ В МЕТКУ
    try:
        with open(LAST_CLEAN_FILE, 'w') as f:
            f.write(current_date.strftime("%Y-%m-%d"))
    except Exception:
        pass


# =====================================================================
# КОНФИГУРАЦИЯ ПОДЧИНЁННОЙ ТАБЛИЦЫ (Учёт каталогов и метаданных)
# =====================================================================

SUB_FILE_NAME = "voyage_sub_data.xlsx"

SUB_TABLE_FIELDS = {
    "UID_Родителя"                  : str,  # Внешний ключ для связи
    "каталоги"                      : str,  # Путь к каталогу с данными
    "наличие метаданных"            : int,  # 0 или 1
    "метаданные добавлены в БМД ЛГА": int,  # 0 или 1
    "Наличие 2-ой копии данных"     : int,  # 0 или 1
    "№ диска с копией"              : int,  # Номер диска (число)
    "каталоги с копией"             : str  # Путь к каталогу с копией
}

SUB_TABLE_HEADINGS_SHORT = {
    "UID_Родителя"                  : "Ключ связи",
    "каталоги"                      : "Каталоги (исходные)",
    "наличие метаданных"            : "Наличие МД",
    "метаданные добавлены в БМД ЛГА": "В БМД ЛГА",
    "Наличие 2-ой копии данных"     : "2-я копия",
    "№ диска с копией"              : "№ диска копии",
    "каталоги с копией"             : "Каталоги копии"
}


def generate_row_uid(row_dict: dict) -> str:
    """Генерирует иммунный к сортировкам UID на основе ключевых полей рейса."""
    pp = str(row_dict.get("№ п/п", "0")).split('.')[0]  # убираем .0 если есть
    vessel = str(row_dict.get("Судно", "")).strip()
    voyage_num = str(row_dict.get("№ рейса", "")).split('.')[0]
    year = str(row_dict.get("Год", "")).split('.')[0]

    basis_string = f"{pp}_{vessel}_{voyage_num}_{year}"
    hash_sha = hashlib.md5(basis_string.encode("utf-8")).hexdigest()[:8]

    return f"{pp}_{hash_sha}"


def check_and_clean_relations(main_df, sub_df):
    """
    Проверяет целостность связей между основной и подчиненной таблицами.
    Удаляет записи-сироты из подчиненной таблицы, у которых нет родителя.
    """
    if main_df.empty or sub_df.empty:
        return sub_df

    # 1. Извлекаем список всех реально существующих UID из основной таблицы рейсов
    # (Если вы еще не перешли на физическую колонку "UID", используйте "id" / "№ п/п")
    existing_parent_keys = set(main_df["UID"].dropna().astype(str)) if "UID" in main_df.columns else set(main_df["id"].dropna().astype(str))

    # Определяем имя колонки связи в подчиненной таблице
    sub_key_col = "UID_Родителя" if "UID_Родителя" in sub_df.columns else "№ п/п"

    # 2. Фильтруем подчиненную таблицу: оставляем ТОЛЬКО те строки,
    # чей ключ связи действительно присутствует в списке существующих рейсов
    initial_count = len(sub_df)

    # Дополнительно очищаем строки с пустыми или некорректными индексами связи
    sub_df_cleaned = sub_df[
        sub_df[sub_key_col].dropna().astype(str).isin(existing_parent_keys)
    ].reset_index(drop=True)

    removed_count = initial_count - len(sub_df_cleaned)
    if removed_count > 0:
        print(f"[Контроль целостности]: Обнаружено и удалено {removed_count} записей-сирот из подчиненной таблицы.")

    return sub_df_cleaned
