import os
import sys

import pytest

sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from renaming_rules import FileNameFormatter


@pytest.mark.parametrize(
    "source, expected",
    [
        ("TAN13715", "TAN 137 15"),
        ("TAN 137 15", "TAN 137 15"),
        ("TAN164-40", "TAN 164 40"),
        ("TAN 164 40_1", "TAN 164 40"),
        ("Random TAN-164-40 extra", "TAN 164 40"),
    ],
)
def test_tan_spacing_preserves_suffix(source: str, expected: str):
    formatter = FileNameFormatter(force_uppercase=False)

    assert formatter._ensure_tan_spacing(source) == expected


@pytest.mark.parametrize(
    "basename, folder, original_abbrev, expected",
    [
        ("AIPI 12 34 567.pdf", "AIPI", "AIPI", "AIPI12-34-567.pdf"),
        ("AN 123.pdf", "AN", "AN", "AN123.pdf"),
        ("ARP 567.doc", "ARP", "ARP", "ARP567.doc"),
        ("BAC 123.txt", "BAC", "BAC", "BAC123.txt"),
        ("Dan-99.docx", "DAN", "DAN", "DAN99.docx"),
        ("nas_1100.pdf", "NAS", "NAS", "NAS1100.pdf"),
        ("hst 100.pdf", "HST", "HST", "HST100.pdf"),
        ("mep123.pdf", "MEP", "MEP", "MEP 123.pdf"),
        ("ne-45.pdf", "NE", "NE", "NE 45.pdf"),
        ("spm_789.pdf", "SPM", "SPM", "SPM 789.pdf"),
        ("tan164-40_1.pdf", "TAN", "TAN", "TAN 164 40.pdf"),
        ("DIN en iso 123.pdf", "DIN", "DIN", "DIN EN ISO 123.pdf"),
        ("DIN iso 123.pdf", "DIN", "DIN", "DIN ISO 123.pdf"),
        ("DIN sae spec 456.pdf", "DIN", "DIN", "DIN SAE SPEC 456.pdf"),
        ("DIN   789.pdf", "DIN", "DIN", "DIN 789.pdf"),
        ("ISO IEC 60027.pdf", "ISO", "ISO", "ISO IEC 60027.pdf"),
        ("ISO SAE PAS 123.pdf", "ISO", "ISO", "ISO 123.pdf"),
        ("ISO TR 987.pdf", "ISO", "ISO", "ISO TR 987.pdf"),
        ("ISO TS 555.pdf", "ISO", "ISO", "ISO TS 555.pdf"),
        ("ISO123.pdf", "ISO", "ISO", "ISO 123.pdf"),
        ("ASTMA123.pdf", "ASTM", "ASTM", "ASTM-A123.pdf"),
        ("MILSTD-2000.pdf", "MIL", "MIL", "MIL-STD-2000.pdf"),
        ("ppph-321.pdf", "PPP", "PPP", "PPP-H-321.pdf"),
        ("AMSQQP416.pdf", "AMS", "AMS", "AMS-QQP416.pdf"),
        ("SAE AMS 5643.pdf", "AMS", "AMS", "AMS5643.pdf"),
        ("BMS8-79_QPL_D.pdf", "BMS", "BMS", "BMS8-79 REV D.QPL.pdf"),
        ("BS EN ISO 9001.pdf", "BS", "BS", "BS EN ISO 9001.pdf"),
        ("bs 123.pdf", "BS", "BS", "BS 123.pdf"),
    ],
)
def test_format_fs_name_normalizes_all_rules(
    basename: str, folder: str, original_abbrev: str, expected: str
) -> None:
    formatter = FileNameFormatter()

    normalized = formatter.format_fs_name(basename, folder, original_abbrev)

    assert normalized == expected
