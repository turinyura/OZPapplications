import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime
import os
import threading # Для выполнения долгих операций в фоновом режиме

# Импорт наших модулей
from config_manager import ConfigManager
from commission_manager import CommissionManager
from report_processor import ReportProcessor
from utils import Utils # Импортируем класс Utils

class ReportFillerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Заполнение паспортов готовности МКД к ОЗП")
        self.geometry("1000x700")
        self.minsize(800, 600)

        self._create_widgets()  # Сначала создаём виджеты!

        # Теперь инициализируем менеджеры
        self.config_manager = ConfigManager()
        self.utils = Utils(self.config_manager)
        self.commission_manager = CommissionManager(self.config_manager, self.log_message)
        self.report_processor = ReportProcessor(self.config_manager, self.commission_manager, self.utils, self.log_message)

        self._setup_logging()
        self._load_initial_settings()
        self._populate_commission_trees()

        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.log_message("Приложение запущено.", level="info")

    def _create_widgets(self):
        """Создает основные виджеты пользовательского интерфейса."""
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.report_filling_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.report_filling_frame, text="Заполнение отчётов")
        self._setup_report_filling_tab()

        self.commission_management_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.commission_management_frame, text="Управление комиссиями")
        self._setup_commission_management_tab()

        self.log_frame = ttk.LabelFrame(self, text="Журнал событий")
        self.log_frame.pack(side="bottom", fill="x", padx=10, pady=5)
        self.log_text = tk.Text(self.log_frame, wrap="word", height=10, state="disabled", font=("Arial", 9))
        self.log_text.pack(expand=True, fill="both", padx=5, pady=5)

        self.log_scroll = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=self.log_scroll.set)
        self.log_scroll.pack(side="right", fill="y", in_=self.log_text)

        self.save_log_button = ttk.Button(self.log_frame, text="Сохранить журнал", command=self._save_log_to_file)
        self.save_log_button.pack(side="right", padx=5, pady=2)

        self.status_bar = ttk.Label(self, text="Готов", relief=tk.SUNKEN, anchor="w")
        self.status_bar.pack(side="bottom", fill="x", padx=10, pady=(0, 5))

    def _setup_report_filling_tab(self):
        """Настраивает виджеты для вкладки "Заполнение отчётов"."""
        path_selection_frame = ttk.LabelFrame(self.report_filling_frame, text="Выбор папок")
        path_selection_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(path_selection_frame, text="Папка с отчётами:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.reports_folder_entry = ttk.Entry(path_selection_frame, width=60)
        self.reports_folder_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(path_selection_frame, text="Выбрать", command=self._select_reports_folder).grid(row=0, column=2, padx=5, pady=2)

        ttk.Label(path_selection_frame, text="Папка с данными:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.data_folder_entry = ttk.Entry(path_selection_frame, width=60)
        self.data_folder_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(path_selection_frame, text="Выбрать", command=self._select_data_folder).grid(row=1, column=2, padx=5, pady=2)

        ttk.Label(path_selection_frame, text="Папка для сохранения:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.output_folder_entry = ttk.Entry(path_selection_frame, width=60)
        self.output_folder_entry.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(path_selection_frame, text="Выбрать", command=self._select_output_folder).grid(row=2, column=2, padx=5, pady=2)

        path_selection_frame.grid_columnconfigure(1, weight=1)

        report_actions_frame = ttk.LabelFrame(self.report_filling_frame, text="Действия с отчётами")
        report_actions_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(report_actions_frame, text="Сканировать отчёты", command=self._scan_reports).pack(side="left", padx=5, pady=5)
        self.report_selection_combobox = ttk.Combobox(report_actions_frame, state="readonly", width=50)
        self.report_selection_combobox.pack(side="left", padx=5, pady=5, expand=True, fill="x")
        self.report_selection_combobox.set("Выберите отчёт")

        ttk.Button(report_actions_frame, text="Заполнить выбранный", command=self._start_fill_selected_report_thread).pack(side="left", padx=5, pady=5)
        ttk.Button(report_actions_frame, text="Заполнить ВСЕ", command=self._start_fill_all_reports_thread).pack(side="left", padx=5, pady=5)

        self.progress_label = ttk.Label(self.report_filling_frame, text="Прогресс: Ожидание...")
        self.progress_label.pack(pady=5)
        self.progress_bar = ttk.Progressbar(self.report_filling_frame, orient="horizontal", mode="determinate", length=500)
        self.progress_bar.pack(pady=5)

    def _setup_commission_management_tab(self):
        """Настраивает виджеты для вкладки "Управление комиссиями"."""
        commission_file_frame = ttk.LabelFrame(self.commission_management_frame, text="Файлы данных комиссий")
        commission_file_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(commission_file_frame, text="Файл типов комиссий (Район, Газ, Состав):").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.commission_types_file_entry = ttk.Entry(commission_file_frame, width=60)
        self.commission_types_file_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(commission_file_frame, text="Выбрать", command=self._select_commission_types_file).grid(row=0, column=2, padx=5, pady=2)
        ttk.Button(commission_file_frame, text="Загрузить", command=self._load_commission_types).grid(row=0, column=3, padx=5, pady=2)

        ttk.Label(commission_file_frame, text="Файл сопоставления Адрес-Район-Газ:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.address_map_file_entry = ttk.Entry(commission_file_frame, width=60)
        self.address_map_file_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(commission_file_frame, text="Выбрать", command=self._select_address_map_file).grid(row=1, column=2, padx=5, pady=2)
        ttk.Button(commission_file_frame, text="Загрузить", command=self._load_address_map).grid(row=1, column=3, padx=5, pady=2)

        commission_file_frame.grid_columnconfigure(1, weight=1)

        commission_types_view_frame = ttk.LabelFrame(self.commission_management_frame, text="Состав комиссий по Районам и Газу")
        commission_types_view_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Treeview для отображения типов комиссий
        self.commission_types_tree = ttk.Treeview(commission_types_view_frame, columns=("Район", "Газ", "Председатель", "Должность Председателя", "Член 1", "Должность Члена 1", "Ресурсник", "Должность Ресурсника"), show="headings")
        self.commission_types_tree.heading("Район", text="Район")
        self.commission_types_tree.heading("Газ", text="Газ")
        self.commission_types_tree.heading("Председатель", text="Председатель")
        self.commission_types_tree.heading("Должность Председателя", text="Должность Председателя")
        self.commission_types_tree.heading("Член 1", text="Член 1")
        self.commission_types_tree.heading("Должность Члена 1", text="Должность Члена 1")
        self.commission_types_tree.heading("Ресурсник", text="Ресурсник")
        self.commission_types_tree.heading("Должность Ресурсника", text="Должность Ресурсника")
        # Настройка ширины колонок (можете настроить по необходимости)
        self.commission_types_tree.column("Район", width=80, stretch=tk.NO)
        self.commission_types_tree.column("Газ", width=50, stretch=tk.NO)
        self.commission_types_tree.column("Председатель", width=120, stretch=tk.YES)
        self.commission_types_tree.column("Должность Председателя", width=150, stretch=tk.YES)
        self.commission_types_tree.column("Член 1", width=120, stretch=tk.YES)
        self.commission_types_tree.column("Должность Члена 1", width=150, stretch=tk.YES)
        self.commission_types_tree.column("Ресурсник", width=120, stretch=tk.YES)
        self.commission_types_tree.column("Должность Ресурсника", width=150, stretch=tk.YES)


        self.commission_types_tree.pack(fill="both", expand=True)

        types_buttons_frame = ttk.Frame(commission_types_view_frame)
        types_buttons_frame.pack(fill="x", pady=5)
        ttk.Button(types_buttons_frame, text="Добавить тип", command=self._add_commission_type).pack(side="left", padx=5)
        ttk.Button(types_buttons_frame, text="Редактировать тип", command=self._edit_commission_type).pack(side="left", padx=5)
        ttk.Button(types_buttons_frame, text="Удалить тип", command=self._delete_commission_type).pack(side="left", padx=5)
        ttk.Button(types_buttons_frame, text="Экспорт типов", command=self._export_commission_types).pack(side="right", padx=5)


        address_map_view_frame = ttk.LabelFrame(self.commission_management_frame, text="Сопоставление Адрес-Район-Газ")
        address_map_view_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Treeview для отображения сопоставления адресов
        self.address_map_tree = ttk.Treeview(address_map_view_frame, columns=("Адрес", "Район", "Газ"), show="headings")
        self.address_map_tree.heading("Адрес", text="Адрес")
        self.address_map_tree.heading("Район", text="Район")
        self.address_map_tree.heading("Газ", text="Газ")
        self.address_map_tree.column("Адрес", width=250, stretch=tk.YES)
        self.address_map_tree.column("Район", width=100, stretch=tk.NO)
        self.address_map_tree.column("Газ", width=50, stretch=tk.NO)
        self.address_map_tree.pack(fill="both", expand=True)

        map_buttons_frame = ttk.Frame(address_map_view_frame)
        map_buttons_frame.pack(fill="x", pady=5)
        ttk.Button(map_buttons_frame, text="Добавить сопоставление", command=self._add_address_map).pack(side="left", padx=5)
        ttk.Button(map_buttons_frame, text="Редактировать сопоставление", command=self._edit_address_map).pack(side="left", padx=5)
        ttk.Button(map_buttons_frame, text="Удалить сопоставление", command=self._delete_address_map).pack(side="left", padx=5)
        ttk.Button(map_buttons_frame, text="Экспорт сопоставлений", command=self._export_address_map).pack(side="right", padx=5)

        def open_field_mapping_window(self, report_fields, data_fields):
            """
            Открывает окно для ручного сопоставления полей отчёта и данных.
            report_fields — список полей из шаблона отчёта
            data_fields — список полей из файла данных
            """
            win = tk.Toplevel(self)
            win.title("Сопоставление полей")
            win.grab_set()
    
            tk.Label(win, text="Поле отчёта").grid(row=0, column=0, padx=5, pady=5)
            tk.Label(win, text="Поле данных").grid(row=0, column=1, padx=5, pady=5)
    
            mapping_vars = {}
            for i, report_field in enumerate(report_fields):
                tk.Label(win, text=report_field).grid(row=i+1, column=0, sticky="w", padx=5, pady=2)
                var = tk.StringVar()
                combo = ttk.Combobox(win, textvariable=var, values=data_fields, width=40)
                combo.grid(row=i+1, column=1, padx=5, pady=2)
                mapping_vars[report_field] = var
    
            def save_mapping():
                mapping = {rf: v.get() for rf, v in mapping_vars.items() if v.get()}
                # Сохраните mapping в self.manual_field_mappings или в файл/настройки
                self.manual_field_mappings = mapping
                win.destroy()
                self.log_message("Сопоставление полей сохранено.", "success")
    
            tk.Button(win, text="Сохранить", command=save_mapping).grid(row=len(report_fields)+1, column=0, columnspan=2, pady=10)



    def _setup_logging(self):
        """Настраивает стили для логов."""
        self.log_text.tag_configure("info", foreground="black")
        self.log_text.tag_configure("warning", foreground="orange")
        self.log_text.tag_configure("error", foreground="red")
        self.log_text.tag_configure("success", foreground="green")
        self.log_text.tag_configure("debug", foreground="gray") # Добавим для отладочных сообщений

    def log_message(self, message, level="info"):
        """
        Добавляет сообщение в текстовое поле логов с меткой времени и цветом.
        """
        timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"{timestamp} [{level.upper()}] {message}\n", level)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

        self.status_bar.config(text=f"Статус: {message}")

    def _save_log_to_file(self):
        """Сохраняет содержимое журнала в текстовый файл."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("Log files", "*.log"), ("Text files", "*.txt"), ("All files", "*.*")],
            title="Сохранить журнал как"
        )
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(self.log_text.get("1.0", tk.END))
                self.log_message(f"Журнал успешно сохранён в {file_path}", level="success")
            except Exception as e:
                self.log_message(f"Ошибка при сохранении журнала: {e}", level="error")
                messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить журнал: {e}")

    def _load_initial_settings(self):
        """Загружает сохранённые настройки при запуске и заполняет поля."""
        self.log_message("Попытка загрузить сохранённые настройки...", level="info")
        try:
            self.config_manager.load_config()
            self.reports_folder_entry.insert(0, self.config_manager.get('Paths', 'reports_folder'))
            self.data_folder_entry.insert(0, self.config_manager.get('Paths', 'data_folder'))
            self.output_folder_entry.insert(0, self.config_manager.get('Paths', 'output_folder'))
            self.commission_types_file_entry.insert(0, self.config_manager.get('Paths', 'commission_types_file'))
            self.address_map_file_entry.insert(0, self.config_manager.get('Paths', 'address_map_file'))
            self.log_message("Настройки загружены.", level="info")

            # Попытаемся загрузить данные комиссий сразу после загрузки путей
            self._load_commission_types(initial_load=True)
            self._load_address_map(initial_load=True)

        except Exception as e:
            self.log_message(f"Ошибка при загрузке настроек или начальных данных: {e}", level="error")

    def _on_closing(self):
        """Обработчик события закрытия окна. Сохраняет настройки перед выходом."""
        self.log_message("Сохранение настроек перед выходом...", level="info")
        try:
            self.config_manager.set('Paths', 'reports_folder', self.reports_folder_entry.get())
            self.config_manager.set('Paths', 'data_folder', self.data_folder_entry.get())
            self.config_manager.set('Paths', 'output_folder', self.output_folder_entry.get())
            self.config_manager.set('Paths', 'commission_types_file', self.commission_types_file_entry.get())
            self.config_manager.set('Paths', 'address_map_file', self.address_map_file_entry.get())
            self.config_manager.save_config()
            self.log_message("Настройки сохранены.", level="info")
        except Exception as e:
            self.log_message(f"Ошибка при сохранении настроек: {e}", level="error")
        self.destroy()

    # --- Методы для выбора папок ---
    def _select_reports_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.reports_folder_entry.delete(0, tk.END)
            self.reports_folder_entry.insert(0, folder_selected)
            self.log_message(f"Выбрана папка с отчётами: {folder_selected}", level="info")
            self.config_manager.set('Paths', 'reports_folder', folder_selected)
            self.config_manager.save_config()

    def _select_data_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.data_folder_entry.delete(0, tk.END)
            self.data_folder_entry.insert(0, folder_selected)
            self.log_message(f"Выбрана папка с данными: {folder_selected}", level="info")
            self.config_manager.set('Paths', 'data_folder', folder_selected)
            self.config_manager.save_config()

    def _select_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, folder_selected)
            self.log_message(f"Выбрана папка для сохранения: {folder_selected}", level="info")
            self.config_manager.set('Paths', 'output_folder', folder_selected)
            self.config_manager.save_config()

    # --- Методы для выбора и загрузки файлов комиссий ---
    def _select_commission_types_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")])
        if file_selected:
            self.commission_types_file_entry.delete(0, tk.END)
            self.commission_types_file_entry.insert(0, file_selected)
            self.log_message(f"Выбран файл типов комиссий: {file_selected}", level="info")
            self.config_manager.set('Paths', 'commission_types_file', file_selected)
            self.config_manager.save_config()

    def _load_commission_types(self, initial_load=False):
        file_path = self.commission_types_file_entry.get()
        if not file_path or not os.path.exists(file_path):
            if not initial_load: # Не выводим предупреждение при первом запуске, если файл не выбран
                self.log_message("Файл типов комиссий не выбран или не существует.", level="warning")
                messagebox.showwarning("Предупреждение", "Пожалуйста, выберите существующий файл типов комиссий.")
            return

        self.log_message(f"Загрузка типов комиссий из: {file_path}", level="info")
        if self.commission_manager.load_commission_types(file_path):
            self._populate_commission_types_tree()
        else:
            self.log_message("Ошибка при загрузке типов комиссий.", level="error")

    def _select_address_map_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")])
        if file_selected:
            self.address_map_file_entry.delete(0, tk.END)
            self.address_map_file_entry.insert(0, file_selected)
            self.log_message(f"Выбран файл сопоставления адресов: {file_selected}", level="info")
            self.config_manager.set('Paths', 'address_map_file', file_selected)
            self.config_manager.save_config()

    def _load_address_map(self, initial_load=False):
        file_path = self.address_map_file_entry.get()
        if not file_path or not os.path.exists(file_path):
            if not initial_load:
                self.log_message("Файл сопоставления адресов не выбран или не существует.", level="warning")
                messagebox.showwarning("Предупреждение", "Пожалуйста, выберите существующий файл сопоставления адресов.")
            return

        self.log_message(f"Загрузка сопоставления адресов из: {file_path}", level="info")
        if self.commission_manager.load_address_commission_map(file_path):
            self._populate_address_map_tree()
        else:
            self.log_message("Ошибка при загрузке сопоставления адресов.", level="error")

    def _populate_commission_trees(self):
        """Обновляет Treeview с данными о комиссиях и сопоставлениях."""
        self._populate_commission_types_tree()
        self._populate_address_map_tree()

    def _populate_commission_types_tree(self):
        """Заполняет Treeview с типами комиссий."""
        for item in self.commission_types_tree.get_children():
            self.commission_types_tree.delete(item)

        data_to_display = self.commission_manager.get_all_commission_types_for_display()
        
        # Обновление колонок, если появляются новые роли
        current_cols = set(self.commission_types_tree["columns"])
        all_possible_cols = set()
        for row_data in data_to_display:
            all_possible_cols.update(row_data.keys())
        
        # Основные колонки, которые всегда должны быть в начале
        fixed_cols = ["Район", "Газ"]
        # Отфильтровываем уже существующие и добавляем новые
        new_cols_to_add = sorted(list(all_possible_cols - current_cols - set(fixed_cols)))

        # Добавление новых колонок (это сложная операция в Treeview и может потребовать пересоздания)
        # Для простоты, пока просто убедимся, что они есть в данных и будут отображены,
        # но если колонка динамически не создана, она не будет показана.
        # Для полноценной динамики нужно пересоздавать Treeview или использовать более продвинутые методы.
        # Предполагаем, что наши основные колонки (Район, Газ, Председатель, Член 1, Ресурсник) уже заданы.
        
        for row_data in data_to_display:
            values = [row_data.get(col, "") for col in self.commission_types_tree["columns"]]
            self.commission_types_tree.insert("", tk.END, values=values)


    def _populate_address_map_tree(self):
        """Заполняет Treeview с сопоставлением адресов."""
        for item in self.address_map_tree.get_children():
            self.address_map_tree.delete(item)

        data_to_display = self.commission_manager.get_all_address_maps_for_display()
        for row_data in data_to_display:
            self.address_map_tree.insert("", tk.END, values=(row_data["Адрес"], row_data["Район"], row_data["Газ"]))

    # --- Методы для работы с отчетами ---
    def _scan_reports(self):
        reports_folder = self.reports_folder_entry.get()
        if not reports_folder or not os.path.isdir(reports_folder):
            self.log_message("Папка с отчётами не выбрана или не существует.", level="error")
            messagebox.showerror("Ошибка", "Пожалуйста, выберите существующую папку с отчётами.")
            return

        self.log_message(f"Сканирование папки с отчётами: {reports_folder}", level="info")
        report_paths = self.report_processor.scan_reports(reports_folder)
        if report_paths:
            # Извлекаем только имена файлов для комбобокса
            display_names = [os.path.basename(p) for p in report_paths]
            self.report_selection_combobox['values'] = display_names
            self.report_selection_combobox.set("Выберите отчёт")
            self.log_message(f"Найдено {len(report_paths)} отчётов.", level="success")
        else:
            self.report_selection_combobox['values'] = []
            self.report_selection_combobox.set("Отчёты не найдены")
            self.log_message("Отчёты не найдены в указанной папке.", level="warning")

    def _start_fill_selected_report_thread(self):
        selected_report_name = self.report_selection_combobox.get()
        if selected_report_name == "Выберите отчёт" or not selected_report_name:
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите отчёт для заполнения.")
            return

        reports_folder = self.reports_folder_entry.get()
        data_folder = self.data_folder_entry.get()
        output_folder = self.output_folder_entry.get()

        if not all([reports_folder, data_folder, output_folder]) or \
           not all(map(os.path.isdir, [reports_folder, data_folder, output_folder])):
            messagebox.showerror("Ошибка", "Пожалуйста, укажите все необходимые папки.")
            return

        # Находим полный путь к выбранному отчету
        selected_report_path = None
        for path in self.report_processor.report_files:
            if os.path.basename(path) == selected_report_name:
                selected_report_path = path
                break
        
        if not selected_report_path:
            self.log_message(f"Полный путь к отчёту '{selected_report_name}' не найден. Возможно, список устарел.", level="error")
            messagebox.showerror("Ошибка", "Не удалось найти выбранный отчёт.")
            return


        self.log_message(f"Начинаю заполнение выбранного отчёта: {selected_report_name}", level="info")
        self.progress_bar['value'] = 0
        self.progress_label['text'] = "Прогресс: Запуск..."

        # Запускаем операцию в отдельном потоке
        threading.Thread(target=self._fill_selected_report_task, args=(selected_report_path, data_folder, output_folder)).start()

    def _fill_selected_report_task(self, report_path, data_folder, output_folder):
        """Задача для заполнения одного отчёта в фоновом потоке."""
        try:
            result = self.report_processor.process_single_report(report_path, data_folder, output_folder)
            self.after(0, self._update_progress_and_log, 100, f"Заполнение отчёта '{os.path.basename(report_path)}' завершено. Статус: {result['status']}.", result['status'].lower())
        except Exception as e:
            self.after(0, self.log_message, f"Ошибка при заполнении отчёта '{os.path.basename(report_path)}': {e}", "error")
            self.after(0, self._update_progress_and_log, 0, "Ошибка заполнения.", "error")


    def _start_fill_all_reports_thread(self):
        reports_folder = self.reports_folder_entry.get()
        data_folder = self.data_folder_entry.get()
        output_folder = self.output_folder_entry.get()

        if not all([reports_folder, data_folder, output_folder]) or \
           not all(map(os.path.isdir, [reports_folder, data_folder, output_folder])):
            messagebox.showerror("Ошибка", "Пожалуйста, укажите все необходимые папки.")
            return
        
        if not self.report_processor.report_files:
            messagebox.showwarning("Предупреждение", "Список отчётов пуст. Сначала просканируйте папку с отчётами.")
            return

        self.log_message("Начинаю заполнение ВСЕХ отчётов...", level="info")
        self.progress_bar['value'] = 0
        self.progress_label['text'] = "Прогресс: Запуск..."

        # Запускаем операцию в отдельном потоке
        threading.Thread(target=self._fill_all_reports_task, args=(reports_folder, data_folder, output_folder)).start()

    def _fill_all_reports_task(self, reports_folder, data_folder, output_folder):
        """Задача для заполнения всех отчётов в фоновом потоке."""
        def update_progress_ui(current, total, file_name):
            progress_value = int((current / total) * 100)
            self.after(0, self.progress_bar.config, value=progress_value)
            self.after(0, self.progress_label.config, text=f"Прогресс: Обработка {current}/{total} - {file_name}")

        try:
            results = self.report_processor.process_all_reports(reports_folder, data_folder, output_folder, update_progress_ui)
            self.after(0, self._update_progress_and_log, 100, "Заполнение ВСЕХ отчётов завершено.", "success")
        except Exception as e:
            self.after(0, self.log_message, f"Ошибка при заполнении всех отчётов: {e}", "error")
            self.after(0, self._update_progress_and_log, 0, "Ошибка заполнения.", "error")

    def _update_progress_and_log(self, value, message, level="info"):
        """Вспомогательная функция для обновления прогресса и логов (из главного потока)."""
        self.progress_bar['value'] = value
        self.progress_label['text'] = f"Прогресс: {message}"
        self.log_message(message, level)


    # --- Методы для управления комиссиями ---
    def _add_commission_type(self):
        def save():
            region = entry_region.get().strip()
            has_gas = var_gas.get()
            chairman = entry_chairman.get().strip()
            chairman_pos = entry_chairman_pos.get().strip()
            member = entry_member.get().strip()
            member_pos = entry_member_pos.get().strip()
            resource = entry_resource.get().strip()
            resource_pos = entry_resource_pos.get().strip()
            if not region:
                messagebox.showwarning("Ошибка", "Укажите район.")
                return
            composition = {
                "Председатель": chairman,
                "Должность Председателя": chairman_pos,
                "Член 1": member,
                "Должность Члена 1": member_pos,
                "Ресурсник": resource,
                "Должность Ресурсника": resource_pos,
            }
            self.commission_manager.add_commission_type(region, has_gas, composition)
            self._populate_commission_types_tree()
            win.destroy()
            self.log_message(f"Добавлен тип комиссии для района '{region}' (Газ: {'Да' if has_gas else 'Нет'})", "success")
    
        win = tk.Toplevel(self)
        win.title("Добавить тип комиссии")
        win.grab_set()
        tk.Label(win, text="Район:").grid(row=0, column=0, sticky="e")
        entry_region = tk.Entry(win)
        entry_region.grid(row=0, column=1)
        tk.Label(win, text="Газ:").grid(row=1, column=0, sticky="e")
        var_gas = tk.BooleanVar()
        tk.Checkbutton(win, variable=var_gas, text="Есть газ").grid(row=1, column=1, sticky="w")
        tk.Label(win, text="Председатель:").grid(row=2, column=0, sticky="e")
        entry_chairman = tk.Entry(win)
        entry_chairman.grid(row=2, column=1)
        tk.Label(win, text="Должность Председателя:").grid(row=3, column=0, sticky="e")
        entry_chairman_pos = tk.Entry(win)
        entry_chairman_pos.grid(row=3, column=1)
        tk.Label(win, text="Член 1:").grid(row=4, column=0, sticky="e")
        entry_member = tk.Entry(win)
        entry_member.grid(row=4, column=1)
        tk.Label(win, text="Должность Члена 1:").grid(row=5, column=0, sticky="e")
        entry_member_pos = tk.Entry(win)
        entry_member_pos.grid(row=5, column=1)
        tk.Label(win, text="Ресурсник:").grid(row=6, column=0, sticky="e")
        entry_resource = tk.Entry(win)
        entry_resource.grid(row=6, column=1)
        tk.Label(win, text="Должность Ресурсника:").grid(row=7, column=0, sticky="e")
        entry_resource_pos = tk.Entry(win)
        entry_resource_pos.grid(row=7, column=1)
        tk.Button(win, text="Сохранить", command=save).grid(row=8, column=0, columnspan=2, pady=10)

    def _edit_commission_type(self):
        selected_item = self.commission_types_tree.focus()
        if not selected_item:
            messagebox.showwarning("Предупреждение", "Выберите тип комиссии для редактирования.")
            return
    
        values = self.commission_types_tree.item(selected_item, 'values')
        region = values[0]
        has_gas = self.utils.get_boolean_from_text(values[1])
        chairman = values[2]
        chairman_pos = values[3]
        member = values[4]
        member_pos = values[5]
        resource = values[6]
        resource_pos = values[7]
    
        def save():
            new_region = entry_region.get().strip()
            new_has_gas = var_gas.get()
            new_chairman = entry_chairman.get().strip()
            new_chairman_pos = entry_chairman_pos.get().strip()
            new_member = entry_member.get().strip()
            new_member_pos = entry_member_pos.get().strip()
            new_resource = entry_resource.get().strip()
            new_resource_pos = entry_resource_pos.get().strip()
            if not new_region:
                messagebox.showwarning("Ошибка", "Укажите район.")
                return
            composition = {
                "Председатель": new_chairman,
                "Должность Председателя": new_chairman_pos,
                "Член 1": new_member,
                "Должность Члена 1": new_member_pos,
                "Ресурсник": new_resource,
                "Должность Ресурсника": new_resource_pos,
            }
            # Удаляем старую запись и добавляем новую (или используйте update, если реализовано)
            self.commission_manager.delete_commission_type(region, has_gas)
            self.commission_manager.add_commission_type(new_region, new_has_gas, composition)
            self._populate_commission_types_tree()
            win.destroy()
            self.log_message(f"Тип комиссии для района '{new_region}' обновлён.", "success")
    
        win = tk.Toplevel(self)
        win.title("Редактировать тип комиссии")
        win.grab_set()
        tk.Label(win, text="Район:").grid(row=0, column=0, sticky="e")
        entry_region = tk.Entry(win)
        entry_region.insert(0, region)
        entry_region.grid(row=0, column=1)
        tk.Label(win, text="Газ:").grid(row=1, column=0, sticky="e")
        var_gas = tk.BooleanVar(value=has_gas)
        tk.Checkbutton(win, variable=var_gas, text="Есть газ").grid(row=1, column=1, sticky="w")
        tk.Label(win, text="Председатель:").grid(row=2, column=0, sticky="e")
        entry_chairman = tk.Entry(win)
        entry_chairman.insert(0, chairman)
        entry_chairman.grid(row=2, column=1)
        tk.Label(win, text="Должность Председателя:").grid(row=3, column=0, sticky="e")
        entry_chairman_pos = tk.Entry(win)
        entry_chairman_pos.insert(0, chairman_pos)
        entry_chairman_pos.grid(row=3, column=1)
        tk.Label(win, text="Член 1:").grid(row=4, column=0, sticky="e")
        entry_member = tk.Entry(win)
        entry_member.insert(0, member)
        entry_member.grid(row=4, column=1)
        tk.Label(win, text="Должность Члена 1:").grid(row=5, column=0, sticky="e")
        entry_member_pos = tk.Entry(win)
        entry_member_pos.insert(0, member_pos)
        entry_member_pos.grid(row=5, column=1)
        tk.Label(win, text="Ресурсник:").grid(row=6, column=0, sticky="e")
        entry_resource = tk.Entry(win)
        entry_resource.insert(0, resource)
        entry_resource.grid(row=6, column=1)
        tk.Label(win, text="Должность Ресурсника:").grid(row=7, column=0, sticky="e")
        entry_resource_pos = tk.Entry(win)
        entry_resource_pos.insert(0, resource_pos)
        entry_resource_pos.grid(row=7, column=1)
        tk.Button(win, text="Сохранить", command=save).grid(row=8, column=0, columnspan=2, pady=10)


    def _delete_commission_type(self):
        selected_items = self.commission_types_tree.selection()
        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите один или несколько типов комиссий для удаления.")
            return

        if messagebox.askyesno("Подтверждение удаления", f"Вы уверены, что хотите удалить {len(selected_items)} выбранных типов комиссий?"):
            deleted_count = 0
            for item_id in selected_items:
                values = self.commission_types_tree.item(item_id, 'values')
                region = values[0]
                has_gas = self.utils.get_boolean_from_text(values[1])
                if self.commission_manager.delete_commission_type(region, has_gas):
                    deleted_count += 1
            self.log_message(f"Удалено {deleted_count} типов комиссий.", level="success")
            self._populate_commission_types_tree() # Обновляем Treeview


    def _export_commission_types(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            title="Экспортировать типы комиссий как"
        )
        if file_path:
            self.commission_manager.export_commission_types(file_path)

    def _add_address_map(self):
        def save():
            address = entry_address.get().strip()
            region = entry_region.get().strip()
            has_gas = var_gas.get()
            if not address or not region:
                messagebox.showwarning("Ошибка", "Укажите адрес и район.")
                return
            self.commission_manager.add_address_map(address, region, has_gas)
            self._populate_address_map_tree()
            win.destroy()
            self.log_message(f"Добавлено сопоставление: {address} → {region} (Газ: {'Да' if has_gas else 'Нет'})", "success")
    
        win = tk.Toplevel(self)
        win.title("Добавить сопоставление адреса")
        win.grab_set()
        tk.Label(win, text="Адрес:").grid(row=0, column=0, sticky="e")
        entry_address = tk.Entry(win, width=40)
        entry_address.grid(row=0, column=1)
        tk.Label(win, text="Район:").grid(row=1, column=0, sticky="e")
        entry_region = tk.Entry(win, width=30)
        entry_region.grid(row=1, column=1)
        tk.Label(win, text="Газ:").grid(row=2, column=0, sticky="e")
        var_gas = tk.BooleanVar()
        tk.Checkbutton(win, variable=var_gas, text="Есть газ").grid(row=2, column=1, sticky="w")
        tk.Button(win, text="Сохранить", command=save).grid(row=3, column=0, columnspan=2, pady=10)

    def _edit_address_map(self):
        selected_item = self.address_map_tree.focus()
        if not selected_item:
            messagebox.showwarning("Предупреждение", "Выберите сопоставление адреса для редактирования.")
            return

        values = self.address_map_tree.item(selected_item, 'values')
        address = values[0]
        self.log_message("Редактирование выбранного сопоставления адреса (не реализовано в UI).", level="info")
        messagebox.showinfo("Информация", f"Редактирование сопоставления адреса '{address}' будет реализовано позже.")
        # Пример использования:
        # self.commission_manager.add_address_map(address, "Обновленный Район", True) # add_address_map также обновляет
        # self._populate_address_map_tree()

    def _delete_address_map(self):
        selected_items = self.address_map_tree.selection()
        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите одно или несколько сопоставлений адресов для удаления.")
            return
        
        if messagebox.askyesno("Подтверждение удаления", f"Вы уверены, что хотите удалить {len(selected_items)} выбранных сопоставлений адресов?"):
            deleted_count = 0
            for item_id in selected_items:
                values = self.address_map_tree.item(item_id, 'values')
                address = values[0]
                if self.commission_manager.delete_address_map(address):
                    deleted_count += 1
            self.log_message(f"Удалено {deleted_count} сопоставлений адресов.", level="success")
            self._populate_address_map_tree() # Обновляем Treeview

    def _export_address_map(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            title="Экспортировать сопоставления адресов как"
        )
        if file_path:
            self.commission_manager.export_address_map(file_path)


if __name__ == "__main__":
    app = ReportFillerApp()
    app.mainloop()