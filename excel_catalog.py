import os
import csv
import re
from datetime import datetime
from typing import Dict, Optional

import openpyxl
from openpyxl.styles import Font

from models import ProcessorConfig
from renaming_rules import FileNameFormatter
from utils import scan_all_files_recursive, looks_like_office_temp

class ExcelCatalog:
    """
    Генерация/обновление Excel-каталога.
    - "Abbrev"     — распознанная аббревиатура (папка).
    - "File_Name"  — display name (нормализованное название стандарта без расширения, UPPERCASE).
    - "Format"     — фактическое расширение OS-файла (pdf/doc/...).
    - "Revision"   — попытка извлечь ревизию (REV/буква/год/выпуск).
    - "Category Tags" — импортируемые теги по нормализованному ключу.
    - "Size_MB" / "Modified" / "SHA256" (пока не считаем — оставляем пусто).
    - "File_Link"  — гиперссылка на реальный путь OS-файла.

    Режим "append" не удаляет пользовательские изменения в существующей книге:
    сравнение по (link_target, modified_text).
    """

    HEADERS = ["Abbrev", "File_Name", "Format", "Revision", "Category Tags",
               "Size_MB", "Modified", "SHA256", "File_Link"]

    def __init__(self, config: ProcessorConfig, log_fn, parent_logger, abbrev_fn):
        self.config = config
        self.log = log_fn
        self.get_abbrev_fn = abbrev_fn
        self.logger = parent_logger.getChild(self.__class__.__name__)
        self.xlsx_path = os.path.join(self.config.destination_folder, self.config.catalog_filename)
        self.formatter = FileNameFormatter(force_uppercase=self.config.force_uppercase_names)

    def _normalize_key(self, s: str) -> str:
        if not s:
            return ""
        base = os.path.splitext(os.path.basename(s))[0]
        base = base.replace('/', '_')
        base = base.replace('-', '').replace('_', '').replace(' ', '')
        return base.lower()

    def _load_tags_map(self) -> Dict[str, str]:
        tags: Dict[str, str] = {}
        if not self.config.tag_file:
            return tags
        try:
            if self.config.tag_file.lower().endswith(".xlsx"):
                wb = openpyxl.load_workbook(self.config.tag_file, data_only=True)
                ws = wb.active
                for row in ws.iter_rows(min_row=2):
                    name_val = (row[0].value or "").strip() if len(row) >= 1 and row[0].value else ""
                    tags_val = (row[6].value or "").strip() if len(row) >= 7 and row[6].value else ""
                    if name_val:
                        key = self._normalize_key(name_val)
                        if key:
                            tags[key] = tags_val
            elif self.config.tag_file.lower().endswith(".csv"):
                with open(self.config.tag_file, "r", encoding="utf-8", newline="") as f:
                    reader = csv.reader(f)
                    rows = list(reader)
                    for r in rows[1:]:
                        if not r:
                            continue
                        name_val = (r[0] or "").strip() if len(r) >= 1 else ""
                        tags_val = (r[6] or "").strip() if len(r) >= 7 else ""
                        if name_val:
                            key = self._normalize_key(name_val)
                            if key:
                                tags[key] = tags_val
        except Exception as e:
            self.log(f"Не удалось импортировать теги: {e}", "WARNING")
        return tags

    def _open_or_create_wb(self):
        wb = None
        ws = None
        existing_index = {}
        if self.config.append_mode and os.path.exists(self.xlsx_path):
            try:
                wb = openpyxl.load_workbook(self.xlsx_path)
                ws = wb.active
                for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
                    link_cell = row[8]
                    modified_cell = row[6]
                    link_target = link_cell.hyperlink.target if (link_cell and link_cell.hyperlink) else ""
                    modified_text = modified_cell.value if modified_cell else ""
                    if link_target and modified_text:
                        existing_index[(link_target, str(modified_text))] = i
                self.log(f"Режим 'Дополнять': найдено {len(existing_index)} записей.", "INFO")
            except Exception as e:
                self.log(f"Не удалось открыть существующий каталог: {e}", "WARNING")
                wb = None
                ws = None

        if wb is None:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Каталог Спецификаций"
            ws.append(self.HEADERS)
            for cell in ws[1]:
                cell.font = Font(bold=True)

        return wb, ws, existing_index

    # --- Извлечение ревизии ---
    def _extract_revision(self, display_name: str, filename_no_ext: str) -> str:
        """
        Грубые, но практичные правила:
        - REV / REVISION + токен (буква/цифры)
        - Окончание '/N' или '_N' (для AMS: N <= 2 символа) -> N
        - Окончание на одиночную букву для AMS/Многих SAE (например, "...-416F") -> буква
        - Год вида 19xx/20xx -> принимаем как ревизию при отсутствии других признаков
        """
        s = display_name.upper()
        fn = filename_no_ext.upper()

        # 1) Явные REV/REVISION
        m = re.search(r'\bREV(?:ISION)?\s*([A-Z0-9\-]+)\b', fn)
        if m:
            return m.group(1)

        # 2) AMS: '/N' или '_N' в конце display
        m = re.search(r'/([A-Z0-9]{1,2})$', s)
        if m:
            return m.group(1)

        # 3) AMS/SAE буква на конце кода "...-416F"
        m = re.search(r'-[0-9A-Z]+([A-Z])$', s)
        if m:
            return m.group(1)

        # 4) Год (как fallback)
        m = re.search(r'\b(19\d{2}|20\d{2})\b', fn)
        if m:
            return m.group(1)

        return ""

    def generate(self):
        self.log("Начало генерации каталога.", "INFO")
        os.makedirs(self.config.destination_folder, exist_ok=True)

        wb, ws, existing_index = self._open_or_create_wb()
        tags_map = self._load_tags_map()
        all_files = scan_all_files_recursive(self.config.destination_folder)
        link_font = Font(color="0000FF", underline="single")
        row_count_added = 0

        for filepath, folder_name, filename in all_files:
            if looks_like_office_temp(filename, self.config.office_temp_prefix):
                continue
            if filename == self.config.catalog_filename:
                continue

            base, ext = os.path.splitext(filename)
            abbrev_tuple = self.get_abbrev_fn(filename)
            if abbrev_tuple and len(abbrev_tuple) >= 2:
                abbrev, orig_abbrev = abbrev_tuple[0], abbrev_tuple[1]
            else:
                abbrev, orig_abbrev = folder_name, folder_name

            display = self.formatter.format_display_name(filename, abbrev, orig_abbrev)

            try:
                stats = os.stat(filepath)
                size_mb = round(stats.st_size / (1024 * 1024), 3)
                modified = datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M')
            except FileNotFoundError:
                size_mb, modified = 0, ""

            key = (filepath, modified)
            if existing_index and key in existing_index:
                continue

            norm_disp = self._normalize_key(display)
            norm_raw = self._normalize_key(filename)
            tags_value = tags_map.get(norm_disp) or tags_map.get(norm_raw) or ""

            revision = self._extract_revision(display, base)

            row = [
                abbrev,
                display,
                ext.lstrip('.').lower(),
                revision,
                tags_value,
                size_mb,
                modified,
                "",
                "Открыть файл"
            ]
            ws.append(row)
            link_cell = ws.cell(row=ws.max_row, column=9)
            link_cell.hyperlink = filepath
            link_cell.font = link_font
            row_count_added += 1

        try:
            wb.save(self.xlsx_path)
            self.log(f"Каталог сохранен: {self.xlsx_path}", "SUCCESS")
            self.log(f"Добавлено записей: {row_count_added}", "INFO")
        except PermissionError:
            self.log("Не удалось сохранить каталог. Файл открыт в другой программе.", "ERROR")
