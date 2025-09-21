"""
SpecSorter v8.2 - точка входа.
Запуск приложения, конфигурация логирования (консоль + файл с ротацией),
обработка критических исключений.
"""

import sys
import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from tkinter import messagebox

from gui import Application
import tkinter as tk

DEFAULT_CONFIG = {
    "version": "8.2",
    "default_file_formats": "pdf, doc, docx, tif, tiff, dwg, xls, xlsx, gif, jpg, JPG",
    "catalog_filename": "specifications_catalog.xlsx",
    "unknown_folder": "_UNKNOWN",
    "corrupt_folder": "_CORRUPT",
    "office_temp_prefix": "~$",
    "max_retry_attempts": 3,
    "retry_delay_sec": 0.5,
    "gui_update_batch_size": 50,
    "default_max_workers": 4,
    "hash_mode": "full",  # "full" | "sampled" | "none"
    "log_file": "specsorter.log",
    "log_level": "INFO",
    "force_uppercase_names": True
}

def configure_root_logger(cfg: dict):
    level = getattr(logging, str(cfg.get("log_level", "INFO")).upper(), logging.INFO)
    logger = logging.getLogger("SpecSorter")
    logger.setLevel(level)

    # консоль
    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(level)
    sh.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))

    # файл с ротацией
    log_path = Path(cfg.get("log_file", "specsorter.log"))
    fh = RotatingFileHandler(log_path, maxBytes=5_000_000, backupCount=3, encoding="utf-8")
    fh.setLevel(level)
    fh.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))

    logger.handlers.clear()
    logger.addHandler(sh)
    logger.addHandler(fh)
    return logger

def load_config(path: Path) -> dict:
    cfg = dict(DEFAULT_CONFIG)
    if path.exists():
        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            cfg.update(data)
        except Exception:
            pass
    return cfg

def main():
    config = load_config(Path("config.json"))
    root_logger = configure_root_logger(config)

    try:
        root = tk.Tk()
        app = Application(master=root, app_logger=root_logger, app_config=config)
        def on_closing():
            if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
                root.destroy()
        root.protocol("WM_DELETE_WINDOW", on_closing)
        app.mainloop()
    except Exception:
        root_logger.critical("Не удалось запустить приложение.", exc_info=True)
        try:
            messagebox.showerror("Критическая ошибка", "Приложение не запустилось, см. логи в консоли/файле.")
        except Exception:
            pass
        sys.exit(1)

if __name__ == "__main__":
    main()
