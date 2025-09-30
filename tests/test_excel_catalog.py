import logging
import os
import sys
from pathlib import Path

import openpyxl

sys.path.append(str(Path(__file__).resolve().parents[1]))

from excel_catalog import ExcelCatalog
from models import ProcessorConfig


def _make_config(tmp_path: Path, catalog_name: str = "catalog.xlsx") -> ProcessorConfig:
    return ProcessorConfig(
        abbreviations_file="",
        destination_folder=str(tmp_path),
        catalog_filename=catalog_name,
        append_mode=True,
    )


def _make_logger():
    return logging.getLogger("excel_catalog_test")


def _make_abbrev_fn(abbrev: str = "LIB"):
    def _fn(_filename: str):
        return (abbrev, abbrev)

    return _fn


def _collect_row_values(xlsx_path: Path):
    wb = openpyxl.load_workbook(xlsx_path)
    try:
        ws = wb.active
        values = []
        for row in ws.iter_rows(min_row=2, max_col=9, values_only=False):
            if not any(cell.value for cell in row):
                continue
            link_cell = row[8]
            values.append(
                {
                    "Abbrev": row[0].value,
                    "File_Name": row[1].value,
                    "Format": row[2].value,
                    "File_Link": link_cell.hyperlink.target if link_cell.hyperlink else None,
                }
            )
        return values
    finally:
        wb.close()


def _run_catalog(tmp_path: Path, catalog_name: str = "catalog.xlsx"):
    config = _make_config(tmp_path, catalog_name)
    log_messages = []

    def _log(msg, level="INFO"):
        log_messages.append((level, msg))

    catalog = ExcelCatalog(config, _log, _make_logger(), _make_abbrev_fn())
    catalog.generate()
    return log_messages


def test_catalog_updates_links_on_rename(tmp_path):
    dest = tmp_path / "library"
    dest.mkdir()

    file_path = dest / "example.pdf"
    file_path.write_text("content")

    _run_catalog(dest)

    xlsx_path = dest / "catalog.xlsx"
    rows = _collect_row_values(xlsx_path)
    assert len(rows) == 1
    original_link = rows[0]["File_Link"]
    assert original_link == str(file_path)

    renamed_path = dest / "example_renamed.pdf"
    os.rename(file_path, renamed_path)

    _run_catalog(dest)

    rows = _collect_row_values(xlsx_path)
    assert len(rows) == 1
    assert rows[0]["File_Link"] == str(renamed_path)
    assert rows[0]["File_Name"] is not None


def test_catalog_removes_missing_files(tmp_path):
    dest = tmp_path / "library"
    dest.mkdir()

    file_one = dest / "keep.pdf"
    file_two = dest / "remove.pdf"
    file_one.write_text("one")
    file_two.write_text("two")

    _run_catalog(dest)

    file_two.unlink()

    _run_catalog(dest)

    xlsx_path = dest / "catalog.xlsx"
    rows = _collect_row_values(xlsx_path)
    assert len(rows) == 1
    assert rows[0]["File_Link"] == str(file_one)
