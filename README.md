# SpecSorter v8.2

SpecSorter is a desktop assistant for maintaining a clean library of engineering and aerospace specifications. It helps teams bring order to sprawling shared folders by recognising abbreviations, enforcing consistent file naming, updating an Excel catalogue and auditing an existing archive for mistakes.

> 🇷🇺 **Описание на русском** расположено в конце файла.

## Table of contents

1. [Key capabilities](#key-capabilities)
2. [How the tool works](#how-the-tool-works)
3. [Requirements](#requirements)
4. [Installation](#installation)
5. [Running the GUI](#running-the-gui)
6. [Preparing input data](#preparing-input-data)
7. [Operating modes](#operating-modes)
8. [Configuration & logging](#configuration--logging)
9. [Testing](#testing)
10. [Project structure](#project-structure)
11. [Русская версия](#русская-версия)

## Key capabilities

- **Abbreviation driven sorting** – recognises specification families by their codes (AMS, BAC*, QPL, ISO …) and places files into the corresponding folders.【F:file_processor.py†L37-L116】
- **Robust renaming** – normalises the filesystem name and display name using more than a hundred handcrafted rules for tricky families (AMS/QPL, BAC, DIN EN ISO, TAN…).【F:renaming_rules.py†L1-L116】【F:renaming_rules.py†L118-L221】
- **Excel catalogue generation** – creates or appends to `specifications_catalog.xlsx` with hyperlinks, detected revisions and optional category tags imported from CSV/XLSX.【F:excel_catalog.py†L1-L99】【F:excel_catalog.py†L101-L160】
- **Duplicate and corruption audit** – scans the destination library, detects misplaced files, conflicts and duplicates using SHA-256 or sampled hashing, and optionally fixes issues automatically.【F:file_processor.py†L118-L204】【F:file_processor.py†L206-L277】
- **Windows shortcut mode** – replaces sorted files with `.lnk` shortcuts when working on Windows while keeping the source data safe.【F:utils.py†L9-L63】
- **Fast, observable processing** – parallel workers with retry logic, incremental GUI progress updates and a dedicated log window.【F:file_processor.py†L206-L277】【F:gui.py†L1-L120】

## How the tool works

1. **Scan sources** – recursively walks every selected source folder, skipping Office temporary files, and filters files by extension.【F:file_processor.py†L129-L176】【F:utils.py†L65-L99】
2. **Detect family** – loads abbreviations from a plain text file and augments them with optional YAML rules for tricky formats.【F:file_processor.py†L44-L116】【F:rules.yaml†L1-L32】
3. **Format names** – applies deterministic rules to produce filesystem names (always uppercase, no duplicated extensions) and display names for the Excel catalogue.【F:renaming_rules.py†L1-L116】【F:file_processor.py†L206-L277】
4. **Execute action** – copies, moves or replaces with a shortcut, resolving conflicts atomically and keeping partial results from leaking into the library.【F:file_processor.py†L206-L277】【F:utils.py†L31-L63】
5. **Update catalogue / audit** – depending on the mode, updates the Excel workbook or runs an audit of the existing library, producing a report and optional fixes.【F:excel_catalog.py†L1-L160】【F:file_processor.py†L279-L420】

## Requirements

- Python 3.9 or newer (tested with CPython).
- [tkinter](https://docs.python.org/3/library/tkinter.html) – ships with the standard Python installer on Windows/macOS; on some Linux distributions it must be installed separately.
- Python packages:
  - `openpyxl` for working with Excel files.
  - `pyyaml` for YAML-based detection rules.
  - `pywin32` *(optional, Windows only)* to create `.lnk` shortcuts.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install --upgrade pip
pip install openpyxl pyyaml pywin32
```

If you plan to extend the project or run the tests, also install the development tooling:

```bash
pip install pytest
```

## Running the GUI

1. Ensure that `config.json` is present in the project root (the default file shipped with the repo is a good starting point).【F:main.py†L42-L64】【F:config.json†L1-L16】
2. Launch the application:

   ```bash
   python main.py
   ```

3. In the GUI:
   - Provide a text file with known abbreviations.
   - Select one or more source folders and the destination library root.
   - Optionally supply a YAML rules file (`rules.yaml`) and a CSV/XLSX file with tag metadata.
   - Choose the operating mode (sorting, catalogue update, library audit) and press the corresponding action button.【F:gui.py†L24-L150】

The window contains a live log view and progress bar. All events are also written to `specsorter.log` (rotating log file).

## Preparing input data

### Abbreviations list

Create a plain text file where each line contains a specification family code, e.g.:

```
AMS
BAC
BACB
ISO
MIL
```

The order does not matter – the application will sort the entries by length so that specific prefixes take precedence when matching filenames.【F:file_processor.py†L52-L90】

### YAML rules (optional)

Use `rules.yaml` to express complex detection logic, such as mapping `DIN EN ISO` to the `DIN` folder or recognising the family directly from a filename pattern. Each rule may set a fixed folder name, derive it from a regex group, and provide the original abbreviation that will appear in the catalogue.【F:rules.yaml†L1-L32】

### Category tags (optional)

Provide a CSV or XLSX file that contains user-maintained metadata. SpecSorter reads the first sheet (or the entire CSV), expecting the specification name in the first column and tags in the seventh column – the lookup is case-insensitive and whitespace agnostic.【F:excel_catalog.py†L33-L80】

## Operating modes

- **Sorting** – scans sources, renames the files and copies/moves or replaces them with shortcuts inside the library. Duplicate names are resolved automatically by appending numeric suffixes.【F:file_processor.py†L206-L277】
- **Catalogue update** – updates `specifications_catalog.xlsx` in the destination folder. In append mode existing manual edits are preserved; new rows are added only when the combination of file hyperlink and modification timestamp is new.【F:excel_catalog.py†L81-L160】
- **Library audit** – analyses the current library structure, verifies naming conventions, checks for duplicates (configurable hash modes: `full`, `sampled`, `none`) and can optionally rename or remove duplicates. Results are shown in the GUI and can be saved to a report file.【F:file_processor.py†L279-L420】

Additional options available from the GUI include:

- Dry-run audit that only reports findings without touching files.
- Automatic move of corrupt files to the `_CORRUPT` folder when rename/delete is enabled.
- Control over the number of worker threads and retry/back-off strategy for transient failures.【F:gui.py†L120-L210】【F:file_processor.py†L279-L420】

## Configuration & logging

`config.json` stores application defaults such as file formats, log level, maximum number of worker threads and folder names for unknown or corrupt files.【F:config.json†L1-L16】

Logging is configured through `main.py` using a rotating file handler; log level and destination file can be customised in `config.json`. GUI events propagate through a thread-safe queue so that long running operations do not freeze the interface.【F:main.py†L17-L41】【F:file_processor.py†L206-L277】【F:gui.py†L18-L57】

## Testing

Run the unit tests with:

```bash
pytest
```

The current suite focuses on the renaming and rule engine logic.【F:tests/test_renaming_rules.py†L1-L200】

## Project structure

```
main.py              # Application entry point and logging setup
config.json          # Default configuration values
file_processor.py    # Core pipeline for scanning, renaming, moving and auditing files
renaming_rules.py    # Rule engine for family detection and name normalisation
excel_catalog.py     # Excel catalogue builder with tag import support
gui.py               # Tkinter based desktop UI
utils.py             # Shared helpers (hashing, atomic copy, Windows shortcuts)
rules.yaml           # Sample YAML rules
```

---

## Русская версия

SpecSorter — это настольное приложение для упорядочивания библиотеки технических спецификаций. Оно распознаёт семейства документов по их аббревиатурам, приводит имена к единому виду, обновляет Excel-каталог с гиперссылками и проводит аудит уже существующей библиотеки.

**Основные возможности**

- Сортировка по аббревиатурам (AMS, BAC*, QPL, ISO и др.) и раскладка файлов по нужным папкам.【F:file_processor.py†L37-L116】
- Нормализация имён по расширенным правилам для «сложных» семейств (AMS/QPL, BAC, DIN EN ISO, TAN…).【F:renaming_rules.py†L1-L116】【F:renaming_rules.py†L118-L221】
- Формирование/обновление Excel-каталога `specifications_catalog.xlsx` с гиперссылками, ревизиями и пользовательскими тегами.【F:excel_catalog.py†L1-L160】
- Аудит библиотеки: проверка размещения, конфликтов и дубликатов (SHA-256 либо быстрый хеш). Возможна автоматическая коррекция ошибок.【F:file_processor.py†L118-L204】【F:file_processor.py†L206-L420】
- Режим ярлыков для Windows — исходные файлы заменяются на `.lnk` только после успешной сортировки.【F:utils.py†L9-L63】

**Быстрый старт**

1. Установите зависимости `openpyxl`, `pyyaml`, (опционально) `pywin32` и запустите `python main.py`.
2. Укажите файлы с аббревиатурами, правилами и тегами, а также исходные и целевые папки.
3. Выберите режим работы (сортировка, обновление каталога, аудит) и нажмите соответствующую кнопку.

Все события отображаются в окне лога и записываются в `specsorter.log`. Дополнительные параметры (количество потоков, список расширений, имена служебных папок) задаются в `config.json`.

