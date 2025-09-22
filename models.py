from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class ProcessorConfig:
    """
    Конфигурация процессора.
    - force_uppercase_names: принудительно приводить имена файлов и display name к UPPERCASE.
    - hash_mode: режим для аудита дубликатов: "full" | "sampled" | "none".
    """
    abbreviations_file: str
    destination_folder: str
    source_folders: List[str] = field(default_factory=list)
    file_formats: str = "pdf, doc, docx, tif, tiff, dwg, xls, xlsx, gif, jpg, JPG"
    mode: str = "copy"  # copy|move|shortcut
    tag_file: Optional[str] = None
    append_mode: bool = True
    dry_run: bool = False
    rename_on_audit: bool = False
    rules_file: Optional[str] = None
    max_workers: int = 4
    audit_report: bool = True
    hash_mode: str = "full"  # "full" | "sampled" | "none"
    unknown_folder: str = "_UNKNOWN"
    corrupt_folder: str = "_CORRUPT"
    catalog_filename: str = "specifications_catalog.xlsx"
    office_temp_prefix: str = "~$"
    move_corrupt: bool = False

    # Дополнительные поля
    max_retry_attempts: int = 3
    retry_delay_sec: float = 0.5
    gui_update_batch_size: int = 50
    force_uppercase_names: bool = True  # новое поле

@dataclass
class FileAction:
    action_type: str
    source_path: str
    destination_folder: Optional[str] = None
    new_filename: Optional[str] = None
    original_path: Optional[str] = None  # для DUPLICATE
