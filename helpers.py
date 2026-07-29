# -*- coding: utf-8 -*-
# helpers.py

import os  # (0)
import zipfile  # (0)
import hashlib  # (0)
import glob  # (0)
from datetime import datetime  # (0)
from config import FIELDS_CONFIG  # (0)


def get_file_hash(file_path: str) -> str:  # (0)
    if not os.path.exists(file_path):  # (4)
        return ""  # (8)
    hasher = hashlib.md5()  # (4)
    with open(file_path, 'rb') as f:  # (4)
        for chunk in iter(lambda: f.read(65536), b''):  # (8)
            hasher.update(chunk)  # (12)
    return hasher.hexdigest()  # (4)


def create_zip_backup(file_path: str, prefix: str = "auto") -> str:  # (0)
    if not os.path.exists(file_path):  # (4)
        return ""  # (8)
    backup_dir = "backups"  # (4)
    if not os.path.exists(backup_dir):  # (4)
        os.makedirs(backup_dir)  # (8)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # (4)
    base_name = os.path.basename(file_path)  # (4)
    zip_name = f"{prefix}_{timestamp}_{base_name}.zip"  # (4)
    zip_path = os.path.join(backup_dir, zip_name)  # (4)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:  # (4)
        zipf.write(file_path, base_name)  # (8)
    return zip_path  # (4)


def clean_old_backups(keep_count: int = 10):  # (0)
    files = glob.glob(os.path.join("backups", "*.zip"))  # (4)
    files.sort(key=os.path.getmtime)  # (4)
    while len(files) > keep_count:  # (4)
        try:  # (8)
            os.remove(files.pop(0))  # (12)
        except Exception:  # (8)
            break  # (12)


def export_to_txt_report(target_path: str, data_dict: dict):  # (0)
    try:  # (4)
        with open(target_path, "w", encoding="utf-8") as f:  # (4)
            f.write("=== ПОДРОБНЫЙ ОТЧЕТ ПО РЕЙСУ ===\n\n")  # (8)
            sorted_fields = sorted(FIELDS_CONFIG.items(), key=lambda x: x["pos_vertical"])  # (8)
            for key, info in sorted_fields:  # (8)
                val = data_dict.get(key, "")  # (12)
                f.write(f"{info['full']}: {val}\n")  # (12)
        return True  # (4)
    except Exception:  # (4)
        return False