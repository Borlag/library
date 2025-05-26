#!/usr/bin/env python
"""
Spec Sync – autonomous tool for ordering aircraft specification files
and drawings on Windows file servers.

This is v1.1
    • Robust abbreviation search (underscores, dots, spaces, hyphens).
    • Empty INDEX.xlsx bug fixed.
    • Journal‑file cross‑check (existing INDEX / registry).
    • Skip copy if file already present in target.
    • GUI polish + Journal picker + extension filter field.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import contextlib
import hashlib
import json
import logging
import os
import re
import shutil
import sqlite3
import sys
import threading
import warnings
from dataclasses import dataclass, field
from datetime import datetime as dt
from pathlib import Path
from typing import Iterable, Optional, Set

# ───────────── third‑party ───────────────────────────────────────────
import openpyxl
import pytesseract
import yaml
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter
from pdfminer.high_level import extract_text
from PIL import Image
from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QRadioButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

try:
    import pythoncom
    from win32com.client import Dispatch
except ImportError:  # pragma: no cover
    pythoncom = None

APP_NAME = "Spec Sync"
DB_NAME = "inventory.db"
LOG_FILE = "sync.log"
EXCEL_NAME = "INDEX.xlsx"
DUMP_DIR = "dumps"
DEFAULT_RULES = "rules.yaml"
DEFAULT_EXTS = {
    ".pdf",
    ".tif",
    ".tiff",
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".gif",
}
REV_PAT = re.compile(r"(?:[_\-\s]REV[_\-\s]?|[_\-\s]ISSUE[_\-\s]?)([A-Z0-9]+)", re.I)

# ────────────────────────────── logging ──────────────────────────────
log = logging.getLogger(APP_NAME)
log.setLevel(logging.DEBUG)
_fh = logging.FileHandler(LOG_FILE, "a", "utf-8")
_fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
log.addHandler(_fh)

warnings.filterwarnings("ignore", category=UserWarning, module="pdfminer")


@dataclass(slots=True)
class FileRecord:
    abbrev: str
    src_path: Path
    dest_path: Path
    fmt: str
    size_mb: float
    mtime: str
    sha256: str
    revision: str = ""
    category: str = ""
    tags: list[str] = field(default_factory=list)


# ────────────────────── utilities ────────────────────────────────────

def sha256_file(path: Path, chunk: int = 1 << 16) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while buf := f.read(chunk):
            h.update(buf)
    return h.hexdigest()


def create_shortcut(link_path: Path, target: Path) -> None:
    if pythoncom is None:
        return
    try:
        pythoncom.CoInitialize()
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(str(link_path))
        shortcut.Targetpath = str(target)
        shortcut.WorkingDirectory = str(target.parent)
        shortcut.IconLocation = str(target)
        shortcut.save()
    except Exception as exc:  # pragma: no cover
        log.debug("Shortcut failure: %s", exc)


def guess_revision(fname: str) -> str:
    m = REV_PAT.search(fname)
    return m.group(1).upper() if m else ""


def safe_extract_pdf(path: Path) -> str:
    """Return first‑page text or empty string on any failure."""
    try:
        return extract_text(str(path), maxpages=1, caching=True)
    except (PDFSyntaxError, ValueError, RuntimeError) as exc:
        log.debug("PDF extract failed %s: %s", path.name, exc)
        return ""


def safe_ocr_image(path: Path) -> str:
    try:
        with Image.open(path) as img:
            return pytesseract.image_to_string(img)
    except (UnidentifiedImageError, OSError) as exc:
        log.debug("OCR failed %s: %s", path.name, exc)
        return ""


def ocr_or_pdf_text(path: Path) -> str:
    if path.suffix.lower() == ".pdf":
        return safe_extract_pdf(path)
    return safe_ocr_image(path)



# ────────────────────── DB layer ─────────────────────────────────────

class InventoryDB:
    def __init__(self, db_file: Path):
        self.db = db_file
        self._init()

    def _init(self) -> None:
        with sqlite3.connect(self.db) as con:
            con.execute(
                """CREATE TABLE IF NOT EXISTS files(
                       sha256 TEXT PRIMARY KEY,
                       dest_path TEXT,
                       size REAL,
                       mtime TEXT)"""
            )
            con.commit()

    def has(self, sha: str) -> bool:
        with sqlite3.connect(self.db) as con:
            cur = con.cursor()
            cur.execute("SELECT 1 FROM files WHERE sha256=?", (sha,))
            return cur.fetchone() is not None

    def add(self, rec: FileRecord) -> None:
        with sqlite3.connect(self.db) as con:
            con.execute(
                "INSERT OR IGNORE INTO files VALUES (?,?,?,?)",
                (rec.sha256, str(rec.dest_path), rec.size_mb, rec.mtime),
            )
            con.commit()


# ─────────────────── categoriser ─────────────────────────────────────

class Categoriser:
    def __init__(self, yaml_file: Path):
        try:
            self.rules = yaml.safe_load(yaml_file.read_text("utf-8")) or {}
        except FileNotFoundError:
            self.rules = {}
        self.comp = {k: re.compile("|".join(v), re.I) for k, v in self.rules.items()}

    def classify(self, rec: FileRecord, text: str) -> str:
        for cat, pat in self.comp.items():
            if pat.search(rec.src_path.name) or pat.search(text):
                return cat
        return "UNCLASSIFIED"


# ─────────────────── SpecSync core ───────────────────────────────────

class SpecSync(QObject):
    progress = Signal(int, int, int)  # scanned, uniq, dup
    message = Signal(str, str)       # level, text
    finished = Signal()

    def __init__(
        self,
        src_roots: list[Path],
        target_root: Path,
        abbr_file: Path,
        case_sensitive: bool,
        mode_move: bool,
        rules_file: Path,
        journal_file: Optional[Path] = None,
        extensions: Optional[Set[str]] = None,
        dry_run: bool = False,
    ):
        super().__init__()
        self.src_roots = src_roots
        self.target_root = target_root
        self.case_sensitive = case_sensitive
        self.mode_move = mode_move
        self.dry = dry_run
        self.extensions = {e.lower().strip() if e.startswith(".") else f".{e.lower().strip()}" for e in (extensions or DEFAULT_EXTS)}

        self.abbrevs = self._load_abbrevs(abbr_file)
        self.db = InventoryDB(target_root / DB_NAME)
        self.cat = Categoriser(rules_file)
        self.journal_names = self._load_journal(journal_file) if journal_file else set()

        self.stop_event = threading.Event()
        self.records: list[FileRecord] = []
        self.scanned = self.uniq = self.dup = 0

    # ───────── helpers ──────────────────────────────────────────────
    def _load_abbrevs(self, file: Path) -> Set[str]:
        with file.open("r", encoding="utf-8") as f:
            raw = {ln.strip() for ln in f if ln.strip()}
        if self.case_sensitive:
            return raw
        return {s.upper() for s in raw}

    def _load_journal(self, path: Path) -> Set[str]:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        names = {row[1].value for row in ws.iter_rows(min_row=2, values_only=True) if row[1].value}
        wb.close()
        return names

    @staticmethod
    def _tokenise(name: str) -> list[str]:
        # split on delimiters and remove empties
        return [tok for tok in re.split(r"[\s_\-.]+", name) if tok]

    def _match_abbrev(self, name: str) -> Optional[str]:
        for tok in self._tokenise(name):
            key = tok if self.case_sensitive else tok.upper()
            if key in self.abbrevs:
                return tok
        return None

    # ───────── main API ─────────────────────────────────────────────
    def run(self) -> None:  # slot for QThread
        try:
            to_process = list(self._iter_files())
            total = len(to_process)
            log.info("%d candidate files", total)

            with concurrent.futures.ThreadPoolExecutor() as ex:
                for rec in ex.map(self._handle_file, to_process):
                    if self.stop_event.is_set():
                        break
                    self.scanned += 1
                    if rec:
                        self.records.append(rec)
                    self.progress.emit(self.scanned, self.uniq, self.dup)

            if not self.dry:
                self._export_index()
            self.message.emit("INFO", "Finished successfully.")
        except Exception as exc:  # pragma: no cover
            log.exception("Fatal")
            self.message.emit("ERROR", str(exc))
        finally:
            self.finished.emit()

    def _iter_files(self) -> Iterable[Path]:
        for root in self.src_roots:
            for p in root.rglob("*"):
                if p.is_file() and p.suffix.lower() in self.extensions:
                    yield p

    def _handle_file(self, path: Path) -> Optional[FileRecord]:
        abbrev = self._match_abbrev(path.stem)
        if not abbrev:
            return None

        sha = sha256_file(path)
        if self.db.has(sha):
            self.dup += 1
            return None

        if path.name in self.journal_names:
            self.message.emit("INFO", f"Skip (journal present): {path.name}")
            self.dup += 1
            return None

        dest_dir = self.target_root / abbrev
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest_path = dest_dir / path.name
        if dest_path.exists():
            self.message.emit("INFO", f"Skip (already in target): {dest_path.name}")
            self.dup += 1
            return None

        if not self.dry:
            if self.mode_move:
                shutil.move(path, dest_path)
                create_shortcut(path.with_suffix(".lnk"), dest_path)
            else:
                shutil.copy2(path, dest_path)

        text_sample = "" if self.dry else ocr_or_pdf_text(dest_path)[:800]

        rec = FileRecord(
            abbrev=abbrev,
            src_path=path,
            dest_path=dest_path,
            fmt=path.suffix.lower()[1:],
            size_mb=round(dest_path.stat().st_size / 2**20, 2) if dest_path.exists() else round(path.stat().st_size / 2**20, 2),
            mtime=dt.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d"),
            sha256=sha,
            revision=guess_revision(path.name),
        )
        rec.category = self.cat.classify(rec, text_sample)
        self.db.add(rec)
        self.uniq += 1
        return rec

    # ───────── index export ─────────────────────────────────────────
    def _export_index(self) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "INDEX"
        header = [
            "Abbrev",
            "File_Name",
            "Format",
            "Revision",
            "Category",
            "Tags",
            "Size_MB",
            "Modified",
            "SHA256",
            "Link",
        ]
        ws.append(header)

        for r in self.records:
            ws.append([
                r.abbrev,
                r.dest_path.name,
                r.fmt.upper(),
                r.revision,
                r.category,
                ", ".join(r.tags),
                r.size_mb,
                r.mtime,
                r.sha256,
                f'=HYPERLINK("{r.dest_path.as_posix()}", "OPEN")',
            ])

        for i, _ in enumerate(header, 1):
            ws.column_dimensions[get_column_letter(i)].bestFit = True
        ws.auto_filter = AutoFilter(ref=f"A1:{get_column_letter(len(header))}{len(self.records)+1}")
        wb.save(self.target_root / EXCEL_NAME)

    # ───────── dumps ────────────────────────────────────────────────
    def create_dump(self) -> Path:
        dump_dir = self.target_root / DUMP_DIR
        dump_dir.mkdir(exist_ok=True)
        ts = dt.now().strftime("%Y%m%d_%H%M%S")
        dump = dump_dir / f"dump_{ts}.json"
        payload = {
            "records": [r.__dict__ for r in self.records],
            "db": (self.target_root / DB_NAME).read_bytes().hex(),
        }
        dump.write_text(json.dumps(payload, indent=2), "utf-8")
        return dump

    def restore_dump(self, dump: Path) -> None:
        obj = json.loads(dump.read_text("utf-8"))
        for r in obj["records"]:
            rec = FileRecord(**r)
            if not rec.dest_path.exists() and rec.src_path.exists():
                shutil.copy2(rec.src_path, rec.dest_path)
        (self.target_root / DB_NAME).write_bytes(bytes.fromhex(obj["db"]))


# ─────────────────── threading wrapper ───────────────────────────────

class Worker(QThread):
    def __init__(self, sync: SpecSync):
        super().__init__()
        self.sync = sync

    def run(self):  # noqa: D401
        self.sync.run()


# ─────────────────── GUI ─────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1000, 640)
        self._build()
        self.worker: Optional[Worker] = None
        self.sync: Optional[SpecSync] = None

    def _build(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        main = QVBoxLayout(central)

        # Source group
        src_grp = QGroupBox("Source roots")
        src_layout = QVBoxLayout(src_grp)
        self.src_list = QListWidget()
        btns = QHBoxLayout()
        btn_add = QPushButton("Add")
        btn_del = QPushButton("Remove")
        btn_add.clicked.connect(self._add_src)
        btn_del.clicked.connect(self._del_src)
        btns.addWidget(btn_add)
        btns.addWidget(btn_del)
        src_layout.addWidget(self.src_list)
        src_layout.addLayout(btns)
        main.addWidget(src_grp)

        # Target & files group
        path_grp = QGroupBox("Paths & files")
        grid = QGridLayout(path_grp)

        self.target_edit, btn_target = self._mk_browse(grid, "Target root", 0, True)
        self.abbrev_edit, btn_abbr = self._mk_browse(grid, "abbreviations.txt", 1, False)
        self.journal_edit, btn_jour = self._mk_browse(grid, "Journal .xlsx (optional)", 2, False)

        # extension filter
        grid.addWidget(QLabel("Extensions (csv, starts with dot):"), 3, 0)
        self.ext_edit = QLineEdit(",".join(sorted(DEFAULT_EXTS)))
        grid.addWidget(self.ext_edit, 3, 1)

        main.addWidget(path_grp)

        # options group
        opt_grp = QGroupBox("Options")
        opt = QGridLayout(opt_grp)
        self.rb_copy = QRadioButton("Copy")
        self.rb_move = QRadioButton("Move + Link")
        self.rb_copy.setChecked(True)
        self.chk_case = QRadioButton("Case sensitive abbrev match")
        self.chk_dry = QRadioButton("Dry‑Run")
        opt.addWidget(self.rb_copy, 0, 0)
        opt.addWidget(self.rb_move, 0, 1)
        opt.addWidget(self.chk_case, 1, 0)
        opt.addWidget(self.chk_dry, 1, 1)
        main.addWidget(opt_grp)

        # progress & log
        self.progress = QProgressBar()
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        main.addWidget(self.progress)
        main.addWidget(self.console, 2)

        # start / cancel
        h = QHBoxLayout()
        self.btn_start = QPushButton("Start")
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setEnabled(False)
        self.btn_start.clicked.connect(self._start)
        self.btn_cancel.clicked.connect(self._cancel)
        h.addWidget(self.btn_start)
        h.addWidget(self.btn_cancel)
        main.addLayout(h)

        # menu safety
        m_tools = self.menuBar().addMenu("Tools")
        act_dump = QAction("Create dump", self)
        act_rest = QAction("Restore dump", self)
        act_dump.triggered.connect(self._mk_dump)
        act_rest.triggered.connect(self._restore_dump)
        m_tools.addAction(act_dump)
        m_tools.addAction(act_rest)

    def _mk_browse(self, grid: QGridLayout, label: str, row: int, dir_: bool) -> tuple[QLineEdit, QPushButton]:
        grid.addWidget(QLabel(label + ":"), row, 0)
        edit = QLineEdit()
        btn = QPushButton("…")
        def cb():
            if dir_:
                p = QFileDialog.getExistingDirectory(self, label)
            else:
                p, _ = QFileDialog.getOpenFileName(self, label)
            if p:
                edit.setText(p)
        btn.clicked.connect(cb)
        grid.addWidget(edit, row, 1)
        grid.addWidget(btn, row, 2)
        return edit, btn

    # ───────── interaction slots ────────────────────────────────────
    def _add_src(self):
        p = QFileDialog.getExistingDirectory(self, "Select folder")
        if p:
            self.src_list.addItem(QListWidgetItem(p))

    def _del_src(self):
        for i in self.src_list.selectedItems():
            self.src_list.takeItem(self.src_list.row(i))

    def _log(self, lvl: str, txt: str):
        colour = {"INFO": "black", "ERROR": "red", "DEBUG": "gray"}.get(lvl, "black")
        self.console.append(f'<span style="color:{colour}">[{lvl}] {txt}</span>')

    def _start(self):
        if self.worker and self.worker.isRunning():
            return

        # validation
        if not self.src_list.count():
            QMessageBox.warning(self, APP_NAME, "Add at least one source root")
            return
        t_root = Path(self.target_edit.text())
        abbr = Path(self.abbrev_edit.text())
        if not t_root.exists():
            QMessageBox.warning(self, APP_NAME, "Target path invalid")
            return
        if not abbr.is_file():
            QMessageBox.warning(self, APP_NAME, "abbreviations.txt missing")
            return

        exts = {e.strip() for e in self.ext_edit.text().split(',') if e.strip()}

        self.sync = SpecSync(
            src_roots=[Path(self.src_list.item(i).text()) for i in range(self.src_list.count())],
            target_root=t_root,
            abbr_file=abbr,
            case_sensitive=self.chk_case.isChecked(),
            mode_move=self.rb_move.isChecked(),
            rules_file=Path(DEFAULT_RULES),
            journal_file=Path(self.journal_edit.text()) if self.journal_edit.text() else None,
            extensions=exts,
            dry_run=self.chk_dry.isChecked(),
        )
        self.sync.progress.connect(self._update)
        self.sync.message.connect(self._log)
        self.worker = Worker(self.sync)
        self.worker.finished.connect(lambda: self._toggle(False))
        self._toggle(True)
        self.worker.start()

    def _toggle(self, running: bool):
        self.btn_start.setEnabled(not running)
        self.btn_cancel.setEnabled(running)

    def _cancel(self):
        if self.sync:
            self.sync.stop_event.set()
        if self.worker:
            self.worker.quit()
        self._toggle(False)

    def _update(self, scanned: int, uniq: int, dup: int):
        self.progress.setValue(scanned)
        self.progress.setMaximum(scanned + dup + uniq)
        self.statusBar().showMessage(f"Scanned: {scanned}  Unique: {uniq}  Dup: {dup}")

    def _mk_dump(self):
        if not self.sync:
            return
        p = self.sync.create_dump()
        QMessageBox.information(self, APP_NAME, f"Dump saved: {p}")

    def _restore_dump(self):
        p, _ = QFileDialog.getOpenFileName(self, "Select dump", str(Path.cwd()))
        if p and self.sync:
            self.sync.restore_dump(Path(p))
            QMessageBox.information(self, APP_NAME, "Restored")

    def closeEvent(self, ev):  # noqa: N802
        if self.worker and self.worker.isRunning():
            self._cancel()
            self.worker.wait(2000)
        ev.accept()


# ─────────────────── CLI entry ───────────────────────────────────────

def parse_cli() -> argparse.Namespace:
    p = argparse.ArgumentParser(APP_NAME)
    p.add_argument("src", nargs="*", type=Path)
    p.add_argument("--target", type=Path, required=True)
    p.add_argument("--abbr", type=Path, required=True)
    p.add_argument("--journal", type=Path)
    p.add_argument("--move", action="store_true")
    p.add_argument("--case", action="store_true")
    p.add_argument("--dry", action="store_true")
    p.add_argument("--ext", type=str, help="Comma list of extensions, e.g. .pdf,.tiff")
    return p.parse_args()


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        ns = parse_cli()
        sync = SpecSync(
            src_roots=ns.src,
            target_root=ns.target,
            abbr_file=ns.abbr,
            case_sensitive=ns.case,
            mode_move=ns.move,
            rules_file=Path(DEFAULT_RULES),
            journal_file=ns.journal,
            extensions={e.strip() for e in (ns.ext.split(',') if ns.ext else DEFAULT_EXTS)},
            dry_run=ns.dry,
        )
        sync.run()
        return

    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
