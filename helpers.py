# -*- coding: utf-8 -*-
# helpers.py
import os
import io
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
BACKUP_DIR = "backups"

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


def get_file_hash_from_bytes(data: bytes) -> str:
    """Вычисляет MD5-хэш содержимого байтов (в памяти)."""
    return hashlib.md5(data).hexdigest()


def get_xlsx_inner_hashes(data) -> dict:
    """
    Возвращает {внутренний_член_xlsx: md5} для xlsx-документа.
    Принимает путь к файлу или байты содержимого (io.BytesIO).
    Служебные метаданные (docProps/* — дата последнего сохранения, автор и т.п.)
    исключаются: они меняются при пересохранении, хотя данные остаются прежними.
    """
    try:
        if isinstance(data, bytes):
            zf = zipfile.ZipFile(io.BytesIO(data), 'r')
        else:
            zf = zipfile.ZipFile(data, 'r')
        with zf:
            result = {}
            for info in zf.infolist():
                if info.filename.startswith("docProps/"):
                    continue
                result[info.filename] = get_file_hash_from_bytes(zf.read(info))
            return result
    except Exception:
        return {}


def get_zip_content_hashes(zip_path: str) -> dict:
    """
    Возвращает {имя_вложенного_файла: кортеж_хэшей_внутренних_членов} для zip-архива бэкапа.
    Сравнение идёт по содержимому ВЛОЖЕННЫХ xlsx (без служебных docProps/*),
    а не по байтам самого архива и не по байтам внешних оболочек.
    """
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            result = {}
            for info in zf.infolist():
                data = zf.read(info)
                inner = get_xlsx_inner_hashes(data)
                if inner:
                    result[info.filename] = tuple(sorted(inner.items()))
            return result
    except Exception:
        return {}


def get_latest_backup(backup_dir=BACKUP_DIR):
    """Возвращает путь к самому свежему архиву в каталоге (или None, если архивов нет)."""
    zips = glob.glob(os.path.join(backup_dir, "*.zip"))
    if not zips:
        return None
    return max(zips, key=os.path.getmtime)


def files_match_archive(file_paths, zip_path) -> bool:
    """
    Сравнивает текущие файлы с содержимым zip-архива бэкапа.
    Возвращает True, если для каждого файла из списка в архиве есть вложенный файл
    с тем же именем и его ДАННЫЕ (внутренние члены xlsx, без служебных docProps/*)
    совпадают. Сравнение по содержимому, а не по байтам оболочек.
    Файлов, которых нет на диске, в сравнении не участвуют.
    """
    if not zip_path or not os.path.exists(zip_path):
        return False

    existing_files = [f for f in file_paths if os.path.exists(f)]
    if not existing_files:
        return False

    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            archived_names = set(zf.namelist())
            for file_path in existing_files:
                name = os.path.basename(file_path)
                if name not in archived_names:
                    return False
                if get_xlsx_inner_hashes(zf.read(name)) != get_xlsx_inner_hashes(file_path):
                    return False
        return True
    except Exception:
        return False


def count_backups(backup_dir=BACKUP_DIR):
    """Возвращает количество zip-архивов в каталоге бэкапов."""
    if not os.path.exists(backup_dir):
        return 0
    return len(glob.glob(os.path.join(backup_dir, "*.zip")))


def _delete_oldest_backups(backup_dir, count):
    """Удаляет самые старые архивы. Возвращает количество фактически удалённых."""
    zips = glob.glob(os.path.join(backup_dir, "*.zip"))
    zips.sort(key=os.path.getmtime)
    deleted = 0
    for old_zip in zips[:count]:
        try:
            os.remove(old_zip)
            deleted += 1
        except Exception:
            pass
    return deleted


def create_backup(file_paths, backup_dir=BACKUP_DIR, enforce_limit=True):
    """
    Создает архив 'backup_дата_время.zip' с текущим содержимым файлов.
    Лимит MAX_BACKUPS при создании не отменяет архивацию:
      - если в папке УЖЕ был лимит MAX_BACKUPS (было ровно MAX) — создаём архив
        и удаляем 1 самый старый (лимит снова соблюдён, предупреждения нет);
      - если в папке УЖЕ было БОЛЬШЕ MAX_BACKUPS — создаём архив, удаляем 1 самый
        старый и выводим предупреждение (переполнение остаётся).
    """
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)

    existing_files = [f for f in file_paths if os.path.exists(f)]
    if not existing_files:
        return ""

    # Считаем количество ДО создания — от него зависит, было ли превышение
    count_before = count_backups(backup_dir)

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    zip_name = os.path.join(backup_dir, f"backup_{timestamp}.zip")

    # Защита от совпадения имени при нескольких вызовах в одну секунду
    counter = 1
    while os.path.exists(zip_name):
        zip_name = os.path.join(backup_dir, f"backup_{timestamp}_{counter}.zip")
        counter += 1

    try:
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in existing_files:
                zipf.write(file_path, os.path.basename(file_path))
    except Exception as e:
        print(f"Ошибка создания бэкапа: {e}")
        return ""

    if enforce_limit and count_before >= MAX_BACKUPS:
        _delete_oldest_backups(backup_dir, 1)
        if count_before > MAX_BACKUPS:
            _show_limit_warning(count_before + 1)

    return zip_name


def _show_limit_warning(n):
    """Выводит предупреждение о превышении лимита архивов (в консоль и GUI)."""
    msg = (f"Превышен лимит архивов: в папке backups {n} шт. при максимуме "
           f"{MAX_BACKUPS} (MAX_BACKUPS). Бэкап создан, но каталог переполнен. "
           f"Очистите каталог вручную, увеличьте лимит или удалите .last_clean "
           f"для автоочистки.")
    print(f"[БЭКАП] {msg}")
    try:
        messagebox.showwarning("Превышен лимит архивов", msg)
    except Exception:
        pass


def clean_duplicate_backups(backup_dir=BACKUP_DIR):
    """
    Плановая очистка папки бэкапов (раз в CLEAN_INTERVAL_DAYS по маркеру .last_clean).
    Этапы:
      1) удаление ДУБЛИКАТОВ (копий) — архивов с тем же содержимым данных, что и у
         уже сохранённого (сравнение по хэшам файлов В АРХИВЕ, а не самого архива).
         Копии удаляются автоматически, без запроса;
      2) если после этого архивов БОЛЬШЕ лимита MAX_BACKUPS — запрос пользователю:
         удалить ли устаревшие архивы сверх лимита (Да/Нет);
      3) если после всего лимит всё ещё превышен (пользователь отказался) — вывод
         напоминания о вариантах: смена лимита, ручная очистка, принудительная
         очистка через удаление маркера .last_clean.
    В консоль выводится раздельная статистика по копиям и устаревшим.
    """
    if not os.path.exists(backup_dir):
        return

    current_date = datetime.now()
    deleted_duplicates = 0  # Удалено копий (дубликатов по содержимому)
    deleted_old = 0         # Удалено устаревших (сверх лимита по количеству)

    # ПРОВЕРКА МЕТКИ ПО ПОНЯТНОЙ ДАТЕ (маркер лежит внутри каталога бэкапов)
    last_clean_file = os.path.join(backup_dir, ".last_clean")
    if os.path.exists(last_clean_file):
        try:
            with open(last_clean_file, 'r') as f:
                last_clean_str = f.read().strip()
            last_clean_date = datetime.strptime(last_clean_str, "%Y-%m-%d")

            days_passed = (current_date - last_clean_date).days
            if days_passed < CLEAN_INTERVAL_DAYS:
                return  # 10 дней еще не прошло, уходим
        except Exception:
            pass

    # ЭТАП 1: УДАЛЕНИЕ КОПИЙ (дубликатов) ПО ХЭШУ СОДЕРЖИМОГО В АРХИВАХ
    zips = glob.glob(os.path.join(backup_dir, "*.zip"))
    zips.sort(key=os.path.getmtime)

    seen_hashes = set()
    unique_zips = []
    for old_zip in zips:
        h = tuple(sorted(get_zip_content_hashes(old_zip).items()))
        if h in seen_hashes:
            try:
                os.remove(old_zip)
                deleted_duplicates += 1  # Удаляем копию (хэш-дубликат содержимого)
            except Exception:
                pass
        else:
            seen_hashes.add(h)
            unique_zips.append(old_zip)

    # ЭТАП 2: ПРИ ПРЕВЫШЕНИИ ЛИМИТА — ЗАПРОС ПЕРЕД УДАЛЕНИЕМ УСТАРЕВШИХ АРХИВОВ
    excess = len(unique_zips) - MAX_BACKUPS
    if excess > 0:
        msg = (f"После удаления копий в папке {os.path.basename(backup_dir)} осталось "
               f"{len(unique_zips)} архивов при лимите {MAX_BACKUPS} (MAX_BACKUPS) — "
               f"превышение на {excess} шт.\n\n"
               f"Удалить {excess} самых старых архивов (устаревших) сейчас?\n"
               f"«Да» — удалить устаревшие архивы до лимита.\n"
               f"«Нет» — оставить все как есть (лимит будет превышен).")
        print(f"[ОЧИСТКА] {msg}")
        try:
            if messagebox.askyesno("Превышен лимит архивов", msg):
                deleted_old = _delete_oldest_backups(backup_dir, excess)
                print(f"[ОЧИСТКА] Удалено устаревших архивов: {deleted_old} шт.")
        except Exception:
            pass

    # ВЫВОД СООБЩЕНИЯ ОБ ОЧИСТКЕ (раздельно: копии и устаревшие)
    if deleted_duplicates > 0 or deleted_old > 0:
        parts = []
        if deleted_duplicates > 0:
            parts.append(f"удалено копий (дубликатов): {deleted_duplicates} шт.")
        if deleted_old > 0:
            parts.append(f"удалено устаревших (сверх лимита): {deleted_old} шт.")
        print(f"[ОЧИСТКА] Обнаружен день плановой проверки. {', '.join(parts)}")
    else:
        print("[ОЧИСТКА] День плановой проверки: папка бэкапов уже в идеальном состоянии.")

    # ЭТАП 3: ЕСЛИ ПОСЛЕ ВСЕГО ВСЁ ЕЩЁ ПРЕВЫШЕН ЛИМИТ — НАПОМНИТЬ О ВАРИАНТАХ
    remaining = count_backups(backup_dir)
    if remaining > MAX_BACKUPS:
        print(f"[ОЧИСТКА] Внимание: лимит {MAX_BACKUPS} (MAX_BACKUPS) всё ещё превышен "
               f"({remaining} архивов). Увеличьте MAX_BACKUPS в helpers.py, удалите лишние "
               f"архивы вручную либо проведите очистку (удалите файл .last_clean "
               f"в каталоге бэкапов и перезапустите программу).")

    # ЗАПИСЬ СВЕЖЕЙ ПОНЯТНОЙ ДАТЫ В МЕТКУ
    try:
        with open(last_clean_file, 'w') as f:
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
