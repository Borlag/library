import os
import logging
import threading
from queue import Queue, Empty
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Listbox

from models import ProcessorConfig
from file_processor import FileProcessor

class Application(tk.Frame):
    """
    GUI:
    - Настройки (аббревиатуры, источники, назначение, форматы, rules.yaml, теги)
    - Режимы (copy/move/shortcut), потоки, аудит/каталог
    - Статус-строка, прогресс, окно логов
    """

    def __init__(self, master: tk.Tk, app_logger: logging.Logger, app_config: dict):
        super().__init__(master)
        self.master = master
        self.logger = app_logger.getChild(self.__class__.__name__)
        self.config_dict = app_config

        self.master.title(f"SpecSorter v{app_config.get('version','8.2')}")
        self.master.geometry("1000x1000")
        self._center_window()
        self.pack(fill="both", expand=True)
        self.progress_queue = Queue()
        self.processor = None
        self._build_ui()
        self.master.after(100, self._process_queue)

    def _center_window(self):
        self.master.update_idletasks()
        w = self.master.winfo_width()
        h = self.master.winfo_height()
        x = (self.master.winfo_screenwidth() // 2) - (w // 2)
        y = (self.master.winfo_screenheight() // 2) - (h // 2)
        self.master.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        self._settings_frame(main)
        self._options_frame(main)
        self._mode_frame(main)
        self._actions_frame(main)
        self._progress_frame(main)
        self._statusbar(main)

        s = ttk.Style()
        s.configure("Accent.TButton", font=("Helvetica", 10, "bold"))

    def _settings_frame(self, parent):
        sf = ttk.LabelFrame(parent, text="Настройки", padding=(10, 10))
        sf.pack(side="top", fill="x", pady=(0, 5))
        sf.columnconfigure(1, weight=1)

        # abbreviations
        self.abbreviations_file = tk.StringVar()
        ttk.Label(sf, text="Файл с аббревиатурами (.txt):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(sf, textvariable=self.abbreviations_file, state="readonly").grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(sf, text="Выбрать...", command=self._select_abbrev).grid(row=0, column=2, sticky="e")

        # sources
        ttk.Label(sf, text="Исходные папки:").grid(row=1, column=0, sticky="nw", pady=2)
        sff = ttk.Frame(sf); sff.grid(row=1, column=1, columnspan=2, sticky="ew")
        sff.columnconfigure(0, weight=1)
        self.source_listbox = Listbox(sff, height=4); self.source_listbox.grid(row=0, column=0, sticky="ew")
        sfb = ttk.Frame(sff); sfb.grid(row=0, column=1, sticky="ns", padx=5)
        ttk.Button(sfb, text="Добавить...", command=self._add_source).pack(fill="x")
        ttk.Button(sfb, text="Удалить", command=self._remove_source).pack(fill="x", pady=2)

        # destination
        self.destination_folder = tk.StringVar()
        ttk.Label(sf, text="Целевая папка (библиотека):").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(sf, textvariable=self.destination_folder, state="readonly").grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(sf, text="Выбрать...", command=self._select_destination).grid(row=2, column=2, sticky="e")

        # formats
        self.file_formats = tk.StringVar(value=self.config_dict.get("default_file_formats"))
        ttk.Label(sf, text="Форматы файлов:").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(sf, textvariable=self.file_formats).grid(row=3, column=1, columnspan=2, sticky="ew", padx=5)

        # rules.yaml
        self.rules_file = tk.StringVar()
        ttk.Label(sf, text="YAML правила (опционально):").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(sf, textvariable=self.rules_file, state="readonly").grid(row=4, column=1, sticky="ew", padx=5)
        ttk.Button(sf, text="Выбрать...", command=self._select_rules).grid(row=4, column=2, sticky="e")

    def _options_frame(self, parent):
        oc = ttk.Frame(parent); oc.pack(fill="x", pady=5)
        oc.columnconfigure(0, weight=1); oc.columnconfigure(1, weight=1)

        # Импорт тегов
        tif = ttk.LabelFrame(oc, text="Импорт тегов", padding=(10, 5))
        tif.grid(row=0, column=0, sticky="nswe", padx=(0, 5)); tif.columnconfigure(1, weight=1)
        self.tag_file = tk.StringVar()
        ttk.Label(tif, text="Файл (XLSX/CSV):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(tif, textvariable=self.tag_file, state="readonly").grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(tif, text="...", command=self._select_tag).grid(row=0, column=2, sticky="e")

        # Каталог
        cmf = ttk.LabelFrame(oc, text="Режим каталога", padding=(10, 5))
        cmf.grid(row=0, column=1, sticky="nswe", padx=(5, 0))
        self.append_mode = tk.BooleanVar(value=True)
        ttk.Checkbutton(cmf, text="Дополнять (не перезаписывать)", variable=self.append_mode).pack(side="left", padx=5)

    def _mode_frame(self, parent):
        mf = ttk.LabelFrame(parent, text="Режим работы с новыми файлами", padding=(10, 5))
        mf.pack(side="top", fill="x", pady=5)
        self.mode = tk.StringVar(value="copy")
        for text, val in [("Копировать","copy"),("Переместить","move"),("Заменить на ярлык","shortcut")]:
            ttk.Radiobutton(mf, text=text, variable=self.mode, value=val).pack(side="left", padx=10)
        wf = ttk.Frame(mf); wf.pack(side="right")
        ttk.Label(wf, text="Потоков:").pack(side="left", padx=(10,5))
        self.max_workers_var = tk.IntVar(value=self.config_dict.get("default_max_workers", 4))
        ttk.Spinbox(wf, from_=1, to=16, textvariable=self.max_workers_var, width=4).pack(side="left")

    def _actions_frame(self, parent):
        af = ttk.LabelFrame(parent, text="Действия", padding=(10, 10))
        af.pack(side="top", fill="x", pady=5)
        af.columnconfigure(0, weight=1); af.columnconfigure(1, weight=1)
        self.start_button = ttk.Button(af, text="НАЙТИ И ОТСОРТИРОВАТЬ", command=self._start_processing, style="Accent.TButton")
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0,5), ipady=5)
        self.update_catalog_button = ttk.Button(af, text="ОБНОВИТЬ КАТАЛОГ", command=self._start_catalog)
        self.update_catalog_button.grid(row=0, column=1, sticky="ew", padx=(5,0), ipady=5)

        adf = ttk.Frame(af); adf.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10,0))
        adf.columnconfigure(0, weight=1)
        self.audit_button = ttk.Button(adf, text="АУДИТ БИБЛИОТЕКИ", command=self._start_audit)
        self.audit_button.grid(row=0, column=0, sticky="ew", ipady=5, padx=(0,5))

        aof = ttk.Frame(adf); aof.grid(row=0, column=1, sticky="w", padx=10)
        self.dry_run_mode = tk.BooleanVar(value=True)
        ttk.Checkbutton(aof, text="Только отчет (без действий)", variable=self.dry_run_mode).pack(anchor="w")
        self.rename_on_audit = tk.BooleanVar(value=False)
        ttk.Checkbutton(aof, text="Переименовывать/удалять дубликаты", variable=self.rename_on_audit).pack(anchor="w")
        self.save_audit_report = tk.BooleanVar(value=True)
        ttk.Checkbutton(aof, text="Сохранять отчет в файл", variable=self.save_audit_report).pack(anchor="w")

        # обработка дубликатов
        dupf = ttk.LabelFrame(adf, text="Проверка дубликатов", padding=(10, 5))
        dupf.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10,0))
        self.hash_mode_var = tk.StringVar(value=self.config_dict.get("hash_mode","full"))
        ttk.Radiobutton(dupf, text="Точная (SHA-256)", variable=self.hash_mode_var, value="full").pack(side="left", padx=8)
        ttk.Radiobutton(dupf, text="Быстрая (сэмплирование)", variable=self.hash_mode_var, value="sampled").pack(side="left", padx=8)
        ttk.Radiobutton(dupf, text="Без проверки", variable=self.hash_mode_var, value="none").pack(side="left", padx=8)

        # поврежденные
        crf = ttk.LabelFrame(adf, text="Повреждённые файлы", padding=(10,5))
        crf.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10,0))
        self.move_corrupt = tk.BooleanVar(value=False)
        ttk.Checkbutton(crf, text="Перемещать поврежденные в _CORRUPT (только при 'Переименовывать...')", variable=self.move_corrupt).pack(anchor="w")

    def _progress_frame(self, parent):
        self.progress = ttk.Progressbar(parent, orient="horizontal", length=100, mode="determinate")
        self.progress.pack(side="top", fill="x", pady=(10, 5))
        lf = ttk.LabelFrame(parent, text="Лог выполнения", padding=(10,5))
        lf.pack(side="top", fill="both", expand=True)
        tf = ttk.Frame(lf); tf.pack(fill="both", expand=True)
        self.log_text = tk.Text(tf, wrap="word", state="disabled", height=15, bg="#f5f5f5", fg="#333")
        sb = ttk.Scrollbar(tf, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.log_text.tag_config("ERROR", foreground="#dc3545")
        self.log_text.tag_config("WARNING", foreground="#ffc107")
        self.log_text.tag_config("SUCCESS", foreground="#28a745")
        self.log_text.tag_config("INFO", foreground="#17a2b8")

    def _statusbar(self, parent):
        self.status_var = tk.StringVar(value="Готово")
        sb = ttk.Label(parent, relief="sunken", anchor="w", textvariable=self.status_var)
        sb.pack(side="bottom", fill="x", pady=(5,0))

    # ---------- Файловые диалоги ----------
    def _select_abbrev(self):
        fn = filedialog.askopenfilename(title="Выберите файл", filetypes=(("Text files","*.txt"),))
        if fn:
            self.abbreviations_file.set(fn)

    def _select_rules(self):
        fn = filedialog.askopenfilename(title="Выберите файл", filetypes=(("YAML files","*.yml *.yaml"),))
        if fn:
            self.rules_file.set(fn)

    def _select_tag(self):
        fn = filedialog.askopenfilename(title="Выберите файл", filetypes=(("Excel","*.xlsx"),("CSV","*.csv")))
        if fn:
            self.tag_file.set(fn)

    def _add_source(self):
        dn = filedialog.askdirectory(title="Выберите папку")
        if dn and dn not in self.source_listbox.get(0, tk.END):
            self.source_listbox.insert(tk.END, dn)

    def _remove_source(self):
        sel = self.source_listbox.curselection()
        if sel:
            for i in reversed(sel):
                self.source_listbox.delete(i)

    def _select_destination(self):
        dn = filedialog.askdirectory(title="Выберите папку")
        if dn:
            self.destination_folder.set(dn)

    # ---------- Лог ----------
    def _add_log(self, msg, tp="INFO"):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, msg + "\n", tp)
        self.log_text.config(state="disabled")
        self.log_text.see(tk.END)

    def _set_buttons(self, enabled: bool):
        st = "normal" if enabled else "disabled"
        for btn in [self.start_button, self.update_catalog_button, self.audit_button]:
            btn.config(state=st)

    # ---------- Валидация ----------
    def _validate_paths(self, require_source=False, require_dest=False, require_abbrev=False) -> bool:
        if require_source and not self.source_listbox.get(0, tk.END):
            messagebox.showerror("Ошибка", "Добавьте исходную папку!")
            return False
        if require_dest and not self.destination_folder.get():
            messagebox.showerror("Ошибка", "Выберите целевую папку!")
            return False
        if require_abbrev and not self.abbreviations_file.get():
            messagebox.showerror("Ошибка", "Выберите файл аббревиатур!")
            return False
        return True

    def _prepare_and_run(self, target_method_name: str, cfg: ProcessorConfig):
        self._set_buttons(False)
        self.progress['value'] = 0
        self.log_text.config(state="normal"); self.log_text.delete(1.0, tk.END); self.log_text.config(state="disabled")
        self.status_var.set("Работа выполняется...")
        try:
            self.processor = FileProcessor(cfg, self.progress_queue, self.logger)
            target = getattr(self.processor, target_method_name)
            threading.Thread(target=target, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Ошибка запуска", f"Не удалось запустить задачу: {e}")
            self._set_buttons(True)
            self.status_var.set("Ошибка запуска")

    # ---------- Кнопки ----------
    def _start_processing(self):
        if not self._validate_paths(require_source=True, require_dest=True, require_abbrev=True):
            return
        cfg = ProcessorConfig(
            abbreviations_file=self.abbreviations_file.get(),
            source_folders=list(self.source_listbox.get(0, tk.END)),
            destination_folder=self.destination_folder.get(),
            file_formats=self.file_formats.get(),
            mode=self.mode.get(),
            tag_file=self.tag_file.get() or None,
            append_mode=True,
            rules_file=self.rules_file.get() or None,
            max_workers=int(self.max_workers_var.get() or 4),
            unknown_folder=self.config_dict.get("unknown_folder","_UNKNOWN"),
            corrupt_folder=self.config_dict.get("corrupt_folder","_CORRUPT"),
            catalog_filename=self.config_dict.get("catalog_filename","specifications_catalog.xlsx"),
            office_temp_prefix=self.config_dict.get("office_temp_prefix","~$"),
            force_uppercase_names=bool(self.config_dict.get("force_uppercase_names", True)),
            hash_mode=self.hash_mode_var.get()
        )
        self._prepare_and_run("run_sorting_only", cfg)

    def _start_catalog(self):
        if not self._validate_paths(require_dest=True, require_abbrev=True):
            return
        cfg = ProcessorConfig(
            abbreviations_file=self.abbreviations_file.get(),
            destination_folder=self.destination_folder.get(),
            append_mode=self.append_mode.get(),
            tag_file=self.tag_file.get() or None,
            rules_file=self.rules_file.get() or None,
            max_workers=int(self.max_workers_var.get() or 4),
            unknown_folder=self.config_dict.get("unknown_folder","_UNKNOWN"),
            corrupt_folder=self.config_dict.get("corrupt_folder","_CORRUPT"),
            catalog_filename=self.config_dict.get("catalog_filename","specifications_catalog.xlsx"),
            office_temp_prefix=self.config_dict.get("office_temp_prefix","~$"),
            force_uppercase_names=bool(self.config_dict.get("force_uppercase_names", True)),
            hash_mode=self.hash_mode_var.get()
        )
        self._prepare_and_run("run_catalog_only", cfg)

    def _start_audit(self):
        if not self._validate_paths(require_dest=True, require_abbrev=True):
            return
        cfg = ProcessorConfig(
            destination_folder=self.destination_folder.get(),
            abbreviations_file=self.abbreviations_file.get(),
            dry_run=self.dry_run_mode.get(),
            rename_on_audit=self.rename_on_audit.get(),
            rules_file=self.rules_file.get() or None,
            max_workers=int(self.max_workers_var.get() or 4),
            audit_report=self.save_audit_report.get(),
            hash_mode=self.hash_mode_var.get(),
            unknown_folder=self.config_dict.get("unknown_folder","_UNKNOWN"),
            corrupt_folder=self.config_dict.get("corrupt_folder","_CORRUPT"),
            catalog_filename=self.config_dict.get("catalog_filename","specifications_catalog.xlsx"),
            office_temp_prefix=self.config_dict.get("office_temp_prefix","~$"),
            force_uppercase_names=bool(self.config_dict.get("force_uppercase_names", True)),
            move_corrupt=self.move_corrupt.get() and self.rename_on_audit.get()
        )
        self._prepare_and_run("run_library_audit", cfg)

    # ---------- Очередь прогресса ----------
    def _process_queue(self):
        try:
            while True:
                msg = self.progress_queue.get_nowait()
                tp, value = msg.get("type"), msg.get("value")
                if tp == "log":
                    log_type = "INFO"
                    if "[ERROR]" in value or "[CRITICAL]" in value:
                        log_type = "ERROR"
                    elif "[WARNING]" in value:
                        log_type = "WARNING"
                    elif "[SUCCESS]" in value:
                        log_type = "SUCCESS"
                    self._add_log(value, log_type)
                elif tp in ("scan_start", "catalog_start"):
                    self.progress.config(mode="indeterminate")
                    self.progress.start(10)
                    self.status_var.set("Сканирование...")
                elif tp == "scan_complete":
                    self.progress.stop()
                    self.progress.config(mode="determinate", maximum=max(value,1), value=0)
                    self.status_var.set("Обработка...")
                elif tp == "progress":
                    self.progress['value'] = value
                elif tp == "confirm_action":
                    actions = msg.get("actions")
                    if messagebox.askyesno("Подтверждение действий", value):
                        if self.processor:
                            threading.Thread(target=self.processor.execute_audit_actions, args=(actions,), daemon=True).start()
                    else:
                        self._add_log("Операция отменена пользователем.", "WARNING")
                        self._set_buttons(True)
                        self.progress.stop()
                        self.status_var.set("Отменено пользователем")
                elif tp in ("finish","finish_no_actions"):
                    self._set_buttons(True)
                    self.progress.stop()
                    self.progress['value'] = self.progress.cget('maximum')
                    final_msg = "Операция успешно завершена!" if tp == "finish" else "Операция завершена. Действий не требуется."
                    messagebox.showinfo("Завершено", final_msg)
                    self.status_var.set("Готово")
        except Empty:
            pass
        finally:
            self.master.after(100, self._process_queue)
