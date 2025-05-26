# ==========  test_extract_abbrev.py  ======================================
"""
Unit-тесты проверяют:
1. Нижний регистр не принимается («and.pdf» → None).
2. Корректная фильтрация и выбор *самого длинного* кода.
"""

import tempfile
from pathlib import Path

from sync_specs import build_abbrev_regex, extract_abbrev

WL = {"AMS-QQ-P-416", "BAC5602", "MIL-STD-500", "AND"}


def test_case_sensitivity():
    pat = build_abbrev_regex(WL)
    assert extract_abbrev("and_connector.pdf", pat) is None
    assert extract_abbrev("AND_connector.pdf", pat) == "AND"


def test_longest_match():
    pat = build_abbrev_regex(WL)
    name = "Spec_AMS-QQ-P-416_RevA.pdf"
    assert extract_abbrev(name, pat) == "AMS-QQ-P-416"


def test_scans_all_valid_files(tmp_path: Path):
    """
    Создаём 4 файла в temp-dir: три валидных + один мусор.
    Проверяем, что SpecSync найдёт ровно три.
    """
    from sync_specs import SpecSync
    #  --- подготовка временного окружения
    src = tmp_path / "src"
    tgt = tmp_path / "tgt"
    src.mkdir()
    tgt.mkdir()
    (tmp_path / "abbreviations.txt").write_text("\n".join(WL), encoding="utf-8")
    (tmp_path / "rules.yaml").write_text("{}", encoding="utf-8")

    good = ["AMS-QQ-P-416.pdf", "BAC5602_revB.pdf", "MIL-STD-500.PDF"]
    bad = ["random_note.pdf"]
    for fn in good + bad:
        (src / fn).write_text("dummy", encoding="utf-8")

    sync = SpecSync(
        source_root=src,
        target_root=tgt,
        abbreviations=tmp_path / "abbreviations.txt",
        rules_yaml=tmp_path / "rules.yaml",
        mode="copy",
        dry_run=True,    # важно: ничего не трогаем
    )
    files = list(sync._iter_files())
    assert len(files) == len(good)
    assert sorted(f.name for f in files) == sorted(good)
