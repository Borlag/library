import os
import re
import time
import shutil
import logging
from dataclasses import asdict
from datetime import datetime
from typing import Dict, List, Optional, Pattern, Tuple, Any
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed  # ThreadPool only
from queue import Queue

from models import ProcessorConfig, FileAction
from renaming_rules import RulesEngine, FileNameFormatter
from utils import (
    get_logger, scan_all_files_recursive, looks_like_office_temp,
    hash_file_sha256, fast_hash_head_tail, create_shortcut_inplace, copy2_atomic
)

from excel_catalog import ExcelCatalog

SUCCESS_LEVEL = 25
logging.addLevelName(SUCCESS_LEVEL, "SUCCESS")

class FileProcessor:
    """
    Основной обработчик:
    - Сканирование исходных каталогов, распознавание семейства
    - Форматирование имён, копирование/перемещение/ярлыки (атомарно)
    - Дубликаты (hash), аудит библиотеки с отчётом и действиями
    """

    ACRONYMS_TREATED_AS_BASE = {'AMS', 'ASTM', 'ASME', 'FED', 'MIL', 'QQ', 'TN', 'ASD', 'AWS', 'PDS'}

    def __init__(self, config: ProcessorConfig, progress_queue: Queue, parent_logger: logging.Logger):
        self.config = config
        self.progress_queue = progress_queue
        self.logger = parent_logger.getChild(self.__class__.__name__)
        self.abbreviations: List[str] = []
        self.abbrev_search_pattern: Optional[Pattern] = None
        self.base_abbrev_pattern: Optional[Pattern] = None
        self._patterns_initialized = False
        self.rules_engine = RulesEngine(self.config.rules_file, self.logger)
        self.formatter = FileNameFormatter(force_uppercase=self.config.force_uppercase_names)
        self.hash_cache: Dict[str, Optional[str]] = {}

    def _log(self, message: str, level: str = "INFO"):
        lvl_map = {"DEBUG":10, "INFO":20, "WARNING":30, "ERROR":40, "CRITICAL":50, "SUCCESS":SUCCESS_LEVEL}
        ts = datetime.now().strftime("%H:%M:%S")
        msg = f"[{ts}] [{level}] {message}"
        self.progress_queue.put({"type":"log","value":msg})
        self.logger.log(lvl_map.get(level, 20), message)

    def _initialize_patterns(self):
        if self._patterns_initialized:
            return
        try:
            with open(self.config.abbreviations_file, 'r', encoding='utf-8') as f:
                self.abbreviations = sorted([line.strip() for line in f if line.strip()], key=len, reverse=True)

            if self.abbreviations:
                # Основной паттерн
                abbrev_pattern = r'^(?:SAE[\s_-]*)?(' + '|'.join(re.escape(a) for a in self.abbreviations) + r')(?![a-zA-Z])'
                self.abbrev_search_pattern = re.compile(abbrev_pattern, re.IGNORECASE)

                # Базовые аббревиатуры с суффиксом
                base_list = [a for a in self.abbreviations if a.upper() in self.ACRONYMS_TREATED_AS_BASE]
                if base_list:
                    base_pattern = r'^(?:SAE[\s_-]*)?(' + '|'.join(re.escape(a) for a in base_list) + r')([A-Z]+)'
                    self.base_abbrev_pattern = re.compile(base_pattern, re.IGNORECASE)
            else:
                self.abbrev_search_pattern = None
                self.base_abbrev_pattern = None
                self._log("Файл аббревиатур пуст. Автоопределение будет выполняться только по правилам.", "WARNING")

            self._patterns_initialized = True
            self._log(f"Паттерны инициализированы. Загружено {len(self.abbreviations)} аббревиатур.")
        except FileNotFoundError:
            self._log(f"Файл с аббревиатурами не найден: {self.config.abbreviations_file}", "CRITICAL")
            raise

    def _get_abbrev_from_filename(self, filename: str) -> Optional[Tuple[str, str]]:
        """
        Возвращает (folder, original_abbrev) или None.
        """
        self._initialize_patterns()

        # 1) YAML rules
        yaml_res = self.rules_engine.apply_rules(filename)
        if yaml_res:
            folder, orig = yaml_res
            if self.rules_engine.is_valid_standard(filename, folder, orig):
                return folder, orig

        # 2) точное совпадение из списка
        if self.abbrev_search_pattern and (m := self.abbrev_search_pattern.search(filename)):
            detected = (m.group(1) or "").strip()
            if detected and self.rules_engine.is_valid_standard(filename, detected, detected):
                folder_token = detected.split()[0].upper()
                if folder_token:
                    return folder_token, detected

        # 3) базовые с суффиксом
        if self.base_abbrev_pattern and (bm := self.base_abbrev_pattern.search(filename)):
            base_abbrev = bm.group(1)
            original_abbrev = base_abbrev + bm.group(2)
            if self.rules_engine.is_valid_standard(filename, base_abbrev, original_abbrev):
                return base_abbrev.upper(), original_abbrev

        return None

    def get_sha256(self, path: str) -> Optional[str]:
        if path in self.hash_cache:
            return self.hash_cache[path]
        _, h = hash_file_sha256(path)
        self.hash_cache[path] = h
        return h

    def _scan_source_directories(self) -> List[Tuple[str, Tuple[str,str]]]:
        self._initialize_patterns()
        formats = [f".{ext.strip().lower()}" for ext in self.config.file_formats.split(',')]
        found_files: List[Tuple[str, Tuple[str,str]]] = []
        for source_dir in self.config.source_folders:
            if not os.path.exists(source_dir):
                continue
            for filepath, _, filename in scan_all_files_recursive(source_dir):
                if looks_like_office_temp(filename, self.config.office_temp_prefix):
                    continue
                if any(filename.lower().endswith(fmt) for fmt in formats):
                    ab = self._get_abbrev_from_filename(filename)
                    if ab:
                        found_files.append((filepath, ab))
        return found_files

    def run_sorting_only(self):
        try:
            self._log("="*60)
            self._log("НАЧАЛО ПРОЦЕССА СОРТИРОВКИ")
            self.progress_queue.put({"type":"scan_start"})
            source_files = self._scan_source_directories()
            self.progress_queue.put({"type":"scan_complete","value":len(source_files)})

            if not source_files:
                self._log("Новых файлов для обработки не найдено.")
                return

            self._log(f"Найдено {len(source_files)} файлов. Обработка...")

            completed = 0
            ok, fail = 0, 0
            self.progress_queue.put({"type":"progress","value":0})

            with ThreadPoolExecutor(max_workers=max(1, self.config.max_workers)) as ex:
                futures = [ex.submit(self._process_single_with_retry, fp, data) for fp, data in source_files]
                for fut in as_completed(futures):
                    res = False
                    try:
                        res = fut.result()
                    except Exception as e:
                        self._log(f"Ошибка обработки: {e}", "WARNING")
                    ok += int(bool(res))
                    fail += int(not res)
                    completed += 1
                    if completed % max(1, self.config.gui_update_batch_size) == 0 or completed == len(source_files):
                        self.progress_queue.put({"type":"progress","value":completed})

            self._log(f"СОРТИРОВКА ЗАВЕРШЕНА: OK={ok}, FAIL={fail}", "SUCCESS")
        except Exception as e:
            self._log(f"Критическая ошибка во время сортировки: {e}", "CRITICAL")
            self.logger.exception("Критическая ошибка в run_sorting_only")
        finally:
            self.progress_queue.put({"type":"finish"})

    def _process_single_with_retry(self, source_path: str, abbrev_data: Tuple[str,str]) -> bool:
        attempts = getattr(self.config, "max_retry_attempts", 3)
        delay = getattr(self.config, "retry_delay_sec", 0.5)
        for attempt in range(attempts):
            try:
                return self._process_single(source_path, abbrev_data)
            except Exception as e:
                if attempt < attempts - 1:
                    self._log(f"Попытка {attempt+1} не удалась: {e}", "WARNING")
                    time.sleep(delay)
                else:
                    self._log(f"Не удалось обработать файл после {attempts} попыток: {e}", "ERROR")
                    return False
        return False

    def _qpl_target_folder(self, fs_name: str, folder: str) -> str:
        """
        Для QPL: определить конечную папку по префиксу в имени (AMS/BMS/MIL/ASTM/BAC*/NAS/MS/LN/ISO/DIN...).
        """
        m = re.match(r'^(AMS|BMS|MIL|ASTM|BAC[A-Z]?|NAS|MS|LN|ISO|DIN|ASME|FED|QQ|TN|ASD|AWS|PDS)', fs_name, flags=re.IGNORECASE)
        if m:
            return m.group(1).upper()
        return folder

    def _process_single(self, source_path: str, abbrev_data: Tuple[str,str]) -> bool:
        if not os.path.isfile(source_path):
            self._log(f"Файл не найден: {source_path}", "WARNING")
            return False

        filename = os.path.basename(source_path)
        folder, original_abbrev = abbrev_data

        # Имя по правилам
        fs_name = self.formatter.format_fs_name(filename, folder, original_abbrev)

        # QPL: выбирать папку из префикса
        target_folder = folder
        if original_abbrev.upper() == "QPL":
            target_folder = self._qpl_target_folder(fs_name, folder)

        dest_dir = os.path.join(self.config.destination_folder, target_folder)
        os.makedirs(dest_dir, exist_ok=True)
        final_dest = os.path.join(dest_dir, fs_name)

        # если источник уже там
        if os.path.abspath(final_dest) == os.path.abspath(source_path):
            self._log(f"Файл уже на месте: {fs_name}", "DEBUG")
            return True

        # Дубликаты по содержимому (при совпадении имени)
        is_dup = False
        if os.path.exists(final_dest):
            s_h = self.get_sha256(source_path)
            d_h = self.get_sha256(final_dest)
            if s_h and d_h and s_h == d_h:
                is_dup = True
                self._log(f"Пропуск дубликата: {fs_name}", "INFO")
            else:
                # уникализировать имя
                idx = 2
                stem, ext = os.path.splitext(fs_name)
                while True:
                    candidate = os.path.join(dest_dir, f"{stem}_{idx}{ext}")
                    if not os.path.exists(candidate):
                        final_dest = candidate
                        break
                    idx += 1
                self._log(f"Файл существует, новое имя: {os.path.basename(final_dest)}", "INFO")

        if not is_dup:
            # Атомарная копия
            copy2_atomic(source_path, final_dest)

        # Режимы
        if self.config.mode == "move":
            try:
                os.remove(source_path)
            except Exception as e:
                self._log(f"Не удалось удалить исходный файл: {e}", "WARNING")
        elif self.config.mode == "shortcut":
            try:
                create_shortcut_inplace(source_path, final_dest, self.logger)
            except Exception as e:
                self._log(f"Не удалось создать ярлык: {e}", "ERROR")
                raise

        return True

    def run_catalog_only(self):
        try:
            self._log("="*60)
            self._log("НАЧАЛО ОБНОВЛЕНИЯ КАТАЛОГА")
            catalog = ExcelCatalog(self.config, self._log, self.logger, self._get_abbrev_from_filename)
            catalog.generate()
        except Exception as e:
            self._log(f"Критическая ошибка при обновлении каталога: {e}", "CRITICAL")
            self.logger.exception("Критическая ошибка в run_catalog_only")
        finally:
            self.progress_queue.put({"type":"finish"})

    def run_library_audit(self):
        try:
            self._log("="*60)
            self._log("НАЧАЛО АУДИТА БИБЛИОТЕКИ")
            mode = "Только отчет" if self.config.dry_run else "Выполнение действий"
            if self.config.rename_on_audit and not self.config.dry_run:
                mode += " с переименованием"
            self._log(f"Режим: {mode}")

            self.progress_queue.put({"type":"scan_start"})
            actions = self._collect_audit_actions()
            if not actions:
                self._log("Проверка завершена. Проблем не найдено.", "SUCCESS")
                self.progress_queue.put({"type":"finish_no_actions"})
                return

            self._report_actions(actions)

            if self.config.audit_report:
                rp = self._export_audit_report(actions)
                if rp:
                    self._log(f"Отчет сохранен: {rp}", "SUCCESS")

            if not self.config.dry_run:
                counts = {}
                for a in actions:
                    counts[a.action_type] = counts.get(a.action_type, 0) + 1
                parts = [f"{v} {k.lower()}" for k, v in counts.items()]
                confirm_msg = f"Найдено {len(actions)} действий: {', '.join(parts)}.\n\n"
                if 'DUPLICATE' in counts:
                    confirm_msg += "ВНИМАНИЕ: Дубликаты будут удалены безвозвратно!\n\n"
                confirm_msg += "Выполнить изменения?"
                self.progress_queue.put({"type": "confirm_action", "value": confirm_msg, "actions": actions})
            else:
                self._log("Аудит завершен (режим отчета).", "INFO")
                self.progress_queue.put({"type":"finish"})
        except Exception as e:
            self._log(f"Критическая ошибка во время аудита: {e}", "CRITICAL")
            self.logger.exception("Критическая ошибка в run_library_audit")
            self.progress_queue.put({"type":"finish"})

    def _collect_audit_actions(self) -> List[FileAction]:
        dest = self.config.destination_folder
        all_tuples = scan_all_files_recursive(dest)
        corrupt_norm = os.path.normcase(self.config.corrupt_folder)
        filepaths = []
        for fp, _parent_folder, filename in all_tuples:
            if looks_like_office_temp(filename, self.config.office_temp_prefix):
                continue
            if filename.lower().endswith('.xlsx'):
                continue
            # Игнорируем файлы внутри папки CORRUPT
            parts_norm = [os.path.normcase(part) for part in Path(fp).parts]
            if corrupt_norm in parts_norm:
                continue
            filepaths.append(fp)
        self._log(f"Найдено {len(filepaths)} файлов для проверки...")
        self.progress_queue.put({"type":"scan_complete","value":len(filepaths)})

        infos: List[Dict[str, Any]] = []
        if self.config.hash_mode == "none":
            batch = max(1, self.config.gui_update_batch_size)
            for i, p in enumerate(filepaths):
                infos.append({
                    "path": p,
                    "hash": None,
                    "folder": os.path.basename(os.path.dirname(p)),
                    "name": os.path.basename(p)
                })
                if ((i + 1) % batch == 0) or (i + 1 == len(filepaths)):
                    self.progress_queue.put({"type": "progress", "value": i + 1})
        else:
            worker = fast_hash_head_tail if self.config.hash_mode == "sampled" else hash_file_sha256
            with ThreadPoolExecutor(max_workers=max(1, self.config.max_workers)) as ex:
                futures = {ex.submit(worker, p): p for p in filepaths}
                for i, fut in enumerate(as_completed(futures)):
                    filepath, sha256 = fut.result()
                    infos.append({
                        "path": filepath, "hash": sha256,
                        "folder": os.path.basename(os.path.dirname(filepath)),
                        "name": os.path.basename(filepath)
                    })
                    if ((i + 1) % max(1, self.config.gui_update_batch_size) == 0) or (i + 1 == len(filepaths)):
                        self.progress_queue.put({"type":"progress","value":i+1})

        # Группировка по хэшу (или по пути, если без хэшей)
        if self.config.hash_mode == "none":
            groups = {info["path"]:[info] for info in infos}
        else:
            groups: Dict[str, List[Dict[str, Any]]] = {}
            for info in infos:
                key = info["hash"] or info["path"]
                groups.setdefault(key, []).append(info)

        actions: List[FileAction] = []

        for key, group in groups.items():
            golden = group[0]
            # Дубликаты
            if self.config.hash_mode != "none" and len(group) > 1:
                for fi in group[1:]:
                    actions.append(FileAction(action_type="DUPLICATE", source_path=fi["path"], original_path=golden["path"]))

            # Проверка положения/именования
            current = golden
            ab = self._get_abbrev_from_filename(current["name"])
            action = FileAction(action_type="", source_path=current["path"])

            is_misplaced = False
            is_malnamed = False
            target_folder = None
            new_name = None

            if ab:
                correct_folder, orig = ab
                # QPL -> целевой префикс
                if orig.upper() == "QPL":
                    fs_name = current["name"]
                    m = re.match(r'^(AMS|BMS|MIL|ASTM|BAC[A-Z]?|NAS|MS|LN|ISO|DIN|ASME|FED|QQ|TN|ASD|AWS|PDS)', fs_name, flags=re.IGNORECASE)
                    if m:
                        correct_folder = m.group(1).upper()

                if current["folder"].upper() != correct_folder.upper():
                    is_misplaced = True
                    target_folder = correct_folder

                if self.config.rename_on_audit:
                    expected = self.formatter.format_fs_name(current["name"], correct_folder, orig)
                    if expected != current["name"]:
                        is_malnamed = True
                        new_name = expected

                if is_misplaced and is_malnamed:
                    action.action_type = "MOVE_RENAME"
                elif is_misplaced:
                    action.action_type = "MOVE"
                elif is_malnamed:
                    action.action_type = "RENAME"
                if action.action_type:
                    action.destination_folder = target_folder
                    action.new_filename = new_name
                    actions.append(action)
            else:
                if current["folder"].upper() != self.config.unknown_folder.upper():
                    actions.append(FileAction(action_type="UNKNOWN", source_path=current["path"], destination_folder=self.config.unknown_folder))

            # Проверка "повреждённости" PDF (заголовок)
            if self.config.rename_on_audit:
                if current["name"].lower().endswith(".pdf"):
                    try:
                        with open(current["path"], "rb") as f:
                            head = f.read(5)
                        if head != b"%PDF-":
                            actions.append(FileAction(action_type="CORRUPT", source_path=current["path"], destination_folder=self.config.corrupt_folder))
                    except Exception:
                        actions.append(FileAction(action_type="CORRUPT", source_path=current["path"], destination_folder=self.config.corrupt_folder))

        return actions

    def _report_actions(self, actions: List[FileAction]):
        self._log(f"\nНайдено {len(actions)} действий:")
        groups: Dict[str, List[FileAction]] = {}
        for a in actions:
            groups.setdefault(a.action_type, []).append(a)
        for t, g in groups.items():
            self._log(f"\n{t}: {len(g)} файлов")
            for a in g[:5]:
                if t == "DUPLICATE":
                    self._log(f"  • '{os.path.basename(a.source_path)}' дубликат '{os.path.basename(a.original_path)}'")
                elif t in ("MOVE", "MOVE_RENAME", "UNKNOWN", "CORRUPT"):
                    self._log(f"  • '{os.path.basename(a.source_path)}' -> папка '{a.destination_folder or ''}'"
                              + (f", новое имя: '{a.new_filename}'" if a.new_filename else ""))
                elif t == "RENAME":
                    self._log(f"  • '{os.path.basename(a.source_path)}' -> '{a.new_filename}'")
            if len(g) > 5:
                self._log(f"  ... и еще {len(g)-5}")

    def execute_audit_actions(self, actions: List[FileAction]):
        self._log("="*60)
        self._log("ВЫПОЛНЕНИЕ ДЕЙСТВИЙ АУДИТА")
        order = {'MOVE':0, 'RENAME':1, 'MOVE_RENAME':2, 'UNKNOWN':3, 'CORRUPT':4, 'DUPLICATE':5}
        actions_sorted = sorted(actions, key=lambda a: order.get(a.action_type, 99))
        self.progress_queue.put({"type":"scan_complete","value":len(actions_sorted)})

        ok = err = 0
        for i, a in enumerate(actions_sorted):
            if self._execute_one(a):
                ok += 1
            else:
                err += 1
            if ((i+1) % max(1, self.config.gui_update_batch_size) == 0) or (i+1 == len(actions_sorted)):
                self.progress_queue.put({"type":"progress","value":i+1})

        self.progress_queue.put({"type":"progress","value":len(actions_sorted)})
        self._log("-"*40)
        self._log(f"Выполнено успешно: {ok}", "SUCCESS")
        if err:
            self._log(f"Ошибок: {err}", "WARNING")
        self.progress_queue.put({"type":"finish"})

    def _execute_one(self, a: FileAction) -> bool:
        if not os.path.exists(a.source_path):
            return True
        try:
            if a.action_type == "DUPLICATE":
                os.remove(a.source_path)
                self._log(f"✓ Удален дубликат: {os.path.basename(a.source_path)}", "SUCCESS")
                return True

            dest_folder = os.path.join(self.config.destination_folder,
                                       a.destination_folder or os.path.basename(os.path.dirname(a.source_path)))
            new_name = a.new_filename or os.path.basename(a.source_path)
            dest_path = os.path.join(dest_folder, new_name)

            if os.path.abspath(a.source_path) == os.path.abspath(dest_path):
                return True

            os.makedirs(dest_folder, exist_ok=True)

            if a.action_type == "CORRUPT" and not self.config.rename_on_audit:
                self._log(f"Обнаружен поврежденный файл (no-op): {os.path.basename(a.source_path)}", "WARNING")
                return True

            if os.path.exists(dest_path):
                self._log(f"Конфликт: '{new_name}' уже существует", "ERROR")
                return False

            shutil.move(a.source_path, dest_path)
            self._log(f"✓ {a.action_type}: {os.path.basename(a.source_path)} -> {os.path.relpath(dest_path, self.config.destination_folder)}", "SUCCESS")
            return True
        except Exception as e:
            self._log(f"Ошибка обработки '{os.path.basename(a.source_path)}': {e}", "ERROR")
            return False

    def _export_audit_report(self, actions: List[FileAction]) -> Optional[str]:
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            rp = os.path.join(self.config.destination_folder, f"audit_report_{ts}.txt")
            with open(rp, "w", encoding="utf-8") as f:
                f.write(f"SpecSorter Audit Report - {datetime.now().isoformat(sep=' ', timespec='seconds')}\n")
                f.write(f"Destination: {self.config.destination_folder}\n")
                f.write("="*80 + "\n\n")
                types = {}
                for a in actions:
                    types.setdefault(a.action_type, []).append(a)
                for t in sorted(types.keys()):
                    f.write(f"{t} ({len(types[t])})\n")
                    for a in types[t]:
                        if t == 'DUPLICATE':
                            f.write(f"  DUP: {a.source_path} <- копия -> {a.original_path}\n")
                        elif t in ('MOVE', 'UNKNOWN'):
                            f.write(f"  MOVE: {a.source_path} -> {a.destination_folder}\n")
                        elif t == 'RENAME':
                            f.write(f"  RENAME: {a.source_path} -> {a.new_filename}\n")
                        elif t == 'MOVE_RENAME':
                            f.write(f"  MOVE_RENAME: {a.source_path} -> {a.destination_folder}/{a.new_filename}\n")
                        elif t == 'CORRUPT':
                            f.write(f"  CORRUPT: {a.source_path} -> {a.destination_folder}\n")
                    f.write("\n")
            return rp
        except Exception as e:
            self._log(f"Не удалось сохранить отчет: {e}", "ERROR")
            return None
