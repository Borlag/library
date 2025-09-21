import os
import sys
import hashlib
import logging
import shutil
import tempfile
from typing import List, Tuple, Optional

# Windows Shortcut Support
SHORTCUTS_AVAILABLE = False
if sys.platform == "win32":
    try:
        import win32com.client  # type: ignore
        SHORTCUTS_AVAILABLE = True
    except Exception:
        SHORTCUTS_AVAILABLE = False

def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(name)

def hash_file_sha256(filepath: str) -> Tuple[str, Optional[str]]:
    sha256_hash = hashlib.sha256()
    try:
        with open(filepath, "rb") as f:
            for byte_block in iter(lambda: f.read(8192), b""):
                sha256_hash.update(byte_block)
        return filepath, sha256_hash.hexdigest()
    except (IOError, PermissionError, FileNotFoundError):
        return filepath, None

def fast_hash_head_tail(filepath: str, sample_size: int = 65536) -> Tuple[str, Optional[str]]:
    try:
        size = os.path.getsize(filepath)
        h = hashlib.sha256()
        with open(filepath, "rb") as f:
            head = f.read(sample_size)
            h.update(head)
            if size > sample_size * 2:
                f.seek(max(0, size - sample_size))
                tail = f.read(sample_size)
                h.update(tail)
        return filepath, h.hexdigest()
    except (IOError, PermissionError, FileNotFoundError):
        return filepath, None

def copy2_atomic(src: str, dst: str) -> None:
    """
    Атомарная копия:
    - копирует во временный файл на том же разделе;
    - затем os.replace -> атомарная замена.
    """
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    dir_name = os.path.dirname(dst) or "."
    base_name = os.path.basename(dst)
    fd, tmp_path = tempfile.mkstemp(prefix=f".{base_name}.part-", dir=dir_name)
    os.close(fd)
    try:
        shutil.copy2(src, tmp_path)
        os.replace(tmp_path, dst)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        finally:
            raise

def create_shortcut_inplace(source_path: str, target_path: str, logger: logging.Logger):
    """
    Создаёт ярлык (.lnk) на месте исходного файла (только Windows):
    - сохраняет имя .lnк на месте source_path (с тем же stem).
    - исходный файл удаляется только если операция создания ярлыка прошла успешно.
    """
    if not SHORTCUTS_AVAILABLE:
        raise RuntimeError("Создание ярлыков недоступно на данной платформе.")

    link_path = os.path.splitext(source_path)[0] + ".lnk"
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(link_path)
        shortcut.TargetPath = os.path.abspath(target_path)
        shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(target_path))
        shortcut.save()
        try:
            os.remove(source_path)
        except Exception as e:
            logger.warning(f"Ярлык создан, но не удалось удалить исходный файл: {e}")
    except Exception as e:
        raise RuntimeError(f"Не удалось создать ярлык: {e}")

def scan_all_files_recursive(path: str) -> List[Tuple[str, str, str]]:
    """Возвращает список (fullpath, parent_folder_name, filename)."""
    result: List[Tuple[str, str, str]] = []
    try:
        with os.scandir(path) as it:
            for entry in it:
                if entry.is_dir(follow_symlinks=False):
                    result.extend(scan_all_files_recursive(entry.path))
                else:
                    result.append((entry.path, os.path.basename(path), entry.name))
    except (PermissionError, FileNotFoundError):
        pass
    return result

def looks_like_office_temp(name: str, prefix: str = "~$") -> bool:
    return name.startswith(prefix)
