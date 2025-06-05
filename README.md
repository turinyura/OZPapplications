
Структура модулей приложения



excel_report_filler/
├── __init__.py
├── main_app.py           # Главный модуль приложения (Tkinter GUI, инициализация)
├── config_manager.py     # Модуль для работы с настройками (сохранение/загрузка путей и регексов)
├── commission_manager.py # Модуль для управления всеми данными комиссий (загрузка, хранение, CRUD)
├── report_processor.py   # Модуль для логики обработки отчётов (чтение, поиск полей, запись, вставка строк)
├── utils.py              # Вспомогательные функции (нечёткое сравнение, парсинг адресов, валидация)
└── ui_components/        # Папка для отдельных UI-компонентов (если интерфейс сильно разрастётся)
    ├── __init__.py
    ├── commission_editor_dialog.py # Диалог редактирования комиссии
    └── settings_dialog.py          # Диалог настроек


Описание модулей и их функций
1. main_app.py
Это будет основной файл, который запускает приложение. Он будет отвечать за:
Инициализацию главного окна Tkinter (root).
Создание экземпляров ttk.Notebook для вкладок.
Создание экземпляров классов, определенных в других модулях (ConfigManager, CommissionManager, ReportProcessor).
Передачу ссылок на необходимые объекты между модулями (например, main_app будет иметь доступ к commission_manager и report_processor).
Создание основных виджетов на вкладках, но их функционал (обработка событий кнопок, обновление таблиц) будет делегирован соответствующим модулям.
Логирование в текстовое поле UI.
2. config_manager.py
Этот модуль будет инкапсулировать всю логику работы с конфигурацией приложения:
Класс ConfigManager:
Методы для сохранения настроек (пути к папкам, регулярные выражения, ручные сопоставления полей, возможно, последняя выбранная тема GUI) в файл (например, config.ini с помощью configparser или config.json с помощью json).
Методы для загрузки настроек при запуске приложения.
Геттеры и сеттеры для доступа к параметрам конфигурации.
3. commission_manager.py
Один из самых важных новых модулей. Он будет управлять всеми данными, связанными с комиссиями:
Класс CommissionManager:
Хранит данные:
self.commission_types: Словарь {(район, наличие_газа): {состав_комиссии}}
self.address_to_commission_map: Словарь {адрес: (район, наличие_газа)}
Методы для загрузки данных о типах комиссий из файла (Excel/CSV).
Методы для загрузки данных о сопоставлении адресов и типов комиссий из отдельного файла.
Методы для создания шаблонов этих файлов.
Методы для экспорта текущих загруженных данных.
Методы CRUD (Create, Read, Update, Delete) для управления записями комиссий и сопоставлений (добавление вручную, удаление, редактирование).
Методы для валидации данных при загрузке и добавлении.
Метод для получения состава комиссии по заданному адресу и наличию газа.
Взаимодействует с utils.py для валидации и парсинга данных.
4. report_processor.py
Этот модуль будет содержать всю основную логику по работе с Excel-отчётами:
Класс ReportProcessor:
Зависимости: При инициализации ему будут переданы ссылки на ConfigManager и CommissionManager.
Метод scan_reports(): Сканирование папки с отчётами, извлечение адресов (используя utils.extract_address_from_filename).
Метод process_single_report(report_path, data_folder_path, output_path): Обработка одного файла.
Метод process_all_reports(): Итерация по всем отчётам, вызывая process_single_report для каждого.
Метод read_data_file(file_path): Чтение данных из файла (Excel/CSV).
Метод find_data_folder_by_address(data_folder, address): Поиск папки с данными по адресу.
Метод update_report_with_data(...):
Загрузка отчёта (openpyxl).
Чтение наличия газа из отчёта.
Получение состава комиссии из CommissionManager на основе адреса и наличия газа.
Логика вставки строк для дополнительных членов комиссии (например, "Ресурсника").
Поиск полей в отчёте (используя utils.fuzzy_match).
Запись значений в ячейки.
Сохранение изменённого отчёта.
Генерация отчёта об обработке: Метод для создания и сохранения финального отчёта после массовой обработки.
Взаимодействует с utils.py для нечёткого сравнения и поиска ячеек.
5. utils.py
Этот модуль будет содержать общие вспомогательные функции:
extract_address_from_filename(filename, regex_pattern): Извлечение адреса с использованием заданного регулярного выражения.
fuzzy_match(str1, str2, threshold): Нечёткое сравнение строк с fuzzywuzzy.
find_value_cell(worksheet, field_row, field_col): Поиск ячейки для записи значения.
is_suitable_for_value(worksheet, cell): Проверка, подходит ли ячейка для записи.
Функции для валидации форматов данных (числа, проценты).
Функция для ручного сопоставления полей (возможно, это будет отдельный UI-компонент, но логика сопоставления может быть здесь).
6. ui_components/ (папка и модули внутри)
По мере роста интерфейса, сложные диалоговые окна или группы виджетов могут быть вынесены в отдельные классы и файлы:
commission_editor_dialog.py: Класс для модального окна редактирования одной записи комиссии.
settings_dialog.py: Класс для окна настроек (где будут регексы, правила сопоставления и т.д.).
Преимущества такой модульной структуры:
Разделение ответственности (SRP): Каждый модуль отвечает за конкретный аспект функциональности (UI, конфигурация, комиссии, обработка отчётов, утилиты).
Улучшенная читаемость: Код становится легче понимать, так как он разбит на логические блоки.
Упрощение отладки: Проблемы легче локализовать в конкретном модуле.
Переиспользуемость: Функции и классы из utils или commission_manager могут быть легко использованы в других частях приложения или даже в других проектах.
Масштабируемость: Добавление новых функций (например, поддержка нового типа отчётов или новый источник данных) требует изменений только в соответствующих модулях, а не во всём коде.
Командная работа: Если над проектом будет работать несколько человек, они смогут работать над разными модулями одновременно без конфликтов.
Следующие шаги:
Создайте папки и пустые файлы согласно предложенной структуре.
Начните переносить существующий код в соответствующие модули и классы.
Определите, какие данные и методы должны быть публичными (доступными для других модулей) и какие приватными.
main_app.py
Этот файл будет содержать главный класс приложения Tkinter, который управляет основным окном, вкладками и взаимодействием между другими модулями.

Python


import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime # Для меток времени в логах
import os       # Для проверки существования файлов/папок (позднее будет в config_manager)

# Заглушки для модулей, которые мы создадим позже
# Позже здесь будут реальные импорты:
# from config_manager import ConfigManager
# from commission_manager import CommissionManager
# from report_processor import ReportProcessor
# from utils import extract_address_from_filename, fuzzy_match

class ReportFillerApp(tk.Tk):
    """
    Главный класс приложения для заполнения отчётов и управления комиссиями.
    """
    def __init__(self):
        super().__init__()
        self.title("Заполнение паспортов готовности МКД к ОЗП")
        self.geometry("1000x700") # Начальный размер окна
        self.minsize(800, 600)    # Минимальный размер окна

        # Инициализация заглушек для будущих менеджеров
        # self.config_manager = ConfigManager()
        # self.commission_manager = CommissionManager(self.config_manager)
        # self.report_processor = ReportProcessor(self.config_manager, self.commission_manager)

        self._create_widgets()
        self._setup_logging()
        self._load_initial_settings() # Попытка загрузить настройки при запуске

        # Привязка функции сохранения настроек к закрытию окна
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        self.log_message("Приложение запущено.", level="info")

    def _create_widgets(self):
        """Создает основные виджеты пользовательского интерфейса."""
        # Создаем Notebook (систему вкладок)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # Вкладка "Заполнение отчётов"
        self.report_filling_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.report_filling_frame, text="Заполнение отчётов")
        self._setup_report_filling_tab()

        # Вкладка "Управление комиссиями"
        self.commission_management_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.commission_management_frame, text="Управление комиссиями")
        self._setup_commission_management_tab()

        # Вкладка "Логи" (или область логов внизу)
        self.log_frame = ttk.LabelFrame(self, text="Журнал событий")
        self.log_frame.pack(side="bottom", fill="x", padx=10, pady=5)
        self.log_text = tk.Text(self.log_frame, wrap="word", height=10, state="disabled", font=("Arial", 9))
        self.log_text.pack(expand=True, fill="both", padx=5, pady=5)

        self.log_scroll = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=self.log_scroll.set)
        self.log_scroll.pack(side="right", fill="y", in_=self.log_text)

        # Кнопка сохранения журнала
        self.save_log_button = ttk.Button(self.log_frame, text="Сохранить журнал", command=self._save_log_to_file)
        self.save_log_button.pack(side="right", padx=5, pady=2)


        # Строка состояния внизу
        self.status_bar = ttk.Label(self, text="Готов", relief=tk.SUNKEN, anchor="w")
        self.status_bar.pack(side="bottom", fill="x", padx=10, pady=(0, 5))

    def _setup_report_filling_tab(self):
        """Настраивает виджеты для вкладки "Заполнение отчётов"."""
        # Рамка для выбора папок
        path_selection_frame = ttk.LabelFrame(self.report_filling_frame, text="Выбор папок")
        path_selection_frame.pack(fill="x", padx=10, pady=5)

        # Путь к папке с отчетами
        ttk.Label(path_selection_frame, text="Папка с отчётами:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.reports_folder_entry = ttk.Entry(path_selection_frame, width=60)
        self.reports_folder_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(path_selection_frame, text="Выбрать", command=self._select_reports_folder).grid(row=0, column=2, padx=5, pady=2)

        # Путь к папке с данными
        ttk.Label(path_selection_frame, text="Папка с данными:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        self.data_folder_entry = ttk.Entry(path_selection_frame, width=60)
        self.data_folder_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(path_selection_frame, text="Выбрать", command=self._select_data_folder).grid(row=1, column=2, padx=5, pady=2)

        # Путь для сохранения заполненных отчетов
        ttk.Label(path_selection_frame, text="Папка для сохранения:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.output_folder_entry = ttk.Entry(path_selection_frame, width=60)
        self.output_folder_entry.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(path_selection_frame, text="Выбрать", command=self._select_output_folder).grid(row=2, column=2, padx=5, pady=2)

        path_selection_frame.grid_columnconfigure(1, weight=1) # Растягиваем поле ввода

        # Рамка для управления отчетами
        report_actions_frame = ttk.LabelFrame(self.report_filling_frame, text="Действия с отчётами")
        report_actions_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(report_actions_frame, text="Сканировать отчёты", command=self._scan_reports).pack(side="left", padx=5, pady=5)
        self.report_selection_combobox = ttk.Combobox(report_actions_frame, state="readonly", width=50)
        self.report_selection_combobox.pack(side="left", padx=5, pady=5, expand=True, fill="x")
        self.report_selection_combobox.set("Выберите отчёт")

        ttk.Button(report_actions_frame, text="Заполнить выбранный", command=self._fill_selected_report).pack(side="left", padx=5, pady=5)
        ttk.Button(report_actions_frame, text="Заполнить ВСЕ", command=self._fill_all_reports).pack(side="left", padx=5, pady=5)

        # Добавим индикатор прогресса
        self.progress_label = ttk.Label(self.report_filling_frame, text="Прогресс: Ожидание...")
        self.progress_label.pack(pady=5)
        self.progress_bar = ttk.Progressbar(self.report_filling_frame, orient="horizontal", mode="determinate", length=500)
        self.progress_bar.pack(pady=5)


    def _setup_commission_management_tab(self):
        """Настраивает виджеты для вкладки "Управление комиссиями"."""
        # Рамка для загрузки/сохранения файлов комиссий
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

        # Рамка для просмотра и редактирования типов комиссий (заглушка)
        commission_types_view_frame = ttk.LabelFrame(self.commission_management_frame, text="Состав комиссий по Районам и Газу")
        commission_types_view_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Здесь будет Treeview для отображения типов комиссий
        self.commission_types_tree = ttk.Treeview(commission_types_view_frame, columns=("Район", "Газ", "Председатель", "Член 1", "Ресурсник"), show="headings")
        self.commission_types_tree.heading("Район", text="Район")
        self.commission_types_tree.heading("Газ", text="Газ")
        self.commission_types_tree.heading("Председатель", text="Председатель")
        self.commission_types_tree.heading("Член 1", text="Член 1")
        self.commission_types_tree.heading("Ресурсник", text="Ресурсник") # Новый столбец
        self.commission_types_tree.column("Район", width=100)
        self.commission_types_tree.column("Газ", width=50)
        self.commission_types_tree.column("Председатель", width=150)
        self.commission_types_tree.column("Член 1", width=150)
        self.commission_types_tree.column("Ресурсник", width=150)
        self.commission_types_tree.pack(fill="both", expand=True)

        # Кнопки для управления типами комиссий
        types_buttons_frame = ttk.Frame(commission_types_view_frame)
        types_buttons_frame.pack(fill="x", pady=5)
        ttk.Button(types_buttons_frame, text="Добавить тип", command=self._add_commission_type).pack(side="left", padx=5)
        ttk.Button(types_buttons_frame, text="Редактировать тип", command=self._edit_commission_type).pack(side="left", padx=5)
        ttk.Button(types_buttons_frame, text="Удалить тип", command=self._delete_commission_type).pack(side="left", padx=5)
        ttk.Button(types_buttons_frame, text="Экспорт типов", command=self._export_commission_types).pack(side="right", padx=5)


        # Рамка для просмотра и редактирования сопоставления Адрес-Район-Газ (заглушка)
        address_map_view_frame = ttk.LabelFrame(self.commission_management_frame, text="Сопоставление Адрес-Район-Газ")
        address_map_view_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Здесь будет Treeview для отображения сопоставления адресов
        self.address_map_tree = ttk.Treeview(address_map_view_frame, columns=("Адрес", "Район", "Газ"), show="headings")
        self.address_map_tree.heading("Адрес", text="Адрес")
        self.address_map_tree.heading("Район", text="Район")
        self.address_map_tree.heading("Газ", text="Газ")
        self.address_map_tree.column("Адрес", width=250)
        self.address_map_tree.column("Район", width=100)
        self.address_map_tree.column("Газ", width=50)
        self.address_map_tree.pack(fill="both", expand=True)

        # Кнопки для управления сопоставлениями адресов
        map_buttons_frame = ttk.Frame(address_map_view_frame)
        map_buttons_frame.pack(fill="x", pady=5)
        ttk.Button(map_buttons_frame, text="Добавить сопоставление", command=self._add_address_map).pack(side="left", padx=5)
        ttk.Button(map_buttons_frame, text="Редактировать сопоставление", command=self._edit_address_map).pack(side="left", padx=5)
        ttk.Button(map_buttons_frame, text="Удалить сопоставление", command=self._delete_address_map).pack(side="left", padx=5)
        ttk.Button(map_buttons_frame, text="Экспорт сопоставлений", command=self._export_address_map).pack(side="right", padx=5)


    def _setup_logging(self):
        """Настраивает стили для логов."""
        self.log_text.tag_configure("info", foreground="black")
        self.log_text.tag_configure("warning", foreground="orange")
        self.log_text.tag_configure("error", foreground="red")
        self.log_text.tag_configure("success", foreground="green")

    def log_message(self, message, level="info"):
        """
        Добавляет сообщение в текстовое поле логов с меткой времени и цветом.
        """
        timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"{timestamp} [{level.upper()}] {message}\n", level)
        self.log_text.see(tk.END) # Прокручивает к последнему сообщению
        self.log_text.config(state="disabled")

        # Обновление строки состояния
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
        """
        Попытка загрузить сохранённые настройки при запуске.
        Этот метод будет использовать ConfigManager после его реализации.
        """
        self.log_message("Попытка загрузить сохранённые настройки...", level="info")
        # Здесь будет вызов:
        # self.config_manager.load_config()
        # self.reports_folder_entry.insert(0, self.config_manager.get_path("reports_folder"))
        # self.data_folder_entry.insert(0, self.config_manager.get_path("data_folder"))
        # self.output_folder_entry.insert(0, self.config_manager.get_path("output_folder"))
        # self.commission_types_file_entry.insert(0, self.config_manager.get_path("commission_types_file"))
        # self.address_map_file_entry.insert(0, self.config_manager.get_path("address_map_file"))
        # self.log_message("Настройки загружены (если файл настроек существует).", level="info")

    def _on_closing(self):
        """
        Обработчик события закрытия окна. Сохраняет настройки перед выходом.
        """
        self.log_message("Сохранение настроек перед выходом...", level="info")
        # Здесь будет вызов:
        # self.config_manager.set_path("reports_folder", self.reports_folder_entry.get())
        # self.config_manager.set_path("data_folder", self.data_folder_entry.get())
        # self.config_manager.set_path("output_folder", self.output_folder_entry.get())
        # self.config_manager.set_path("commission_types_file", self.commission_types_file_entry.get())
        # self.config_manager.set_path("address_map_file", self.address_map_file_entry.get())
        # self.config_manager.save_config()
        self.log_message("Настройки сохранены. Закрытие приложения.", level="info")
        self.destroy()

    # --- Методы для выбора папок (заглушки) ---
    def _select_reports_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.reports_folder_entry.delete(0, tk.END)
            self.reports_folder_entry.insert(0, folder_selected)
            self.log_message(f"Выбрана папка с отчётами: {folder_selected}", level="info")

    def _select_data_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.data_folder_entry.delete(0, tk.END)
            self.data_folder_entry.insert(0, folder_selected)
            self.log_message(f"Выбрана папка с данными: {folder_selected}", level="info")

    def _select_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, folder_selected)
            self.log_message(f"Выбрана папка для сохранения: {folder_selected}", level="info")

    # --- Методы для выбора файлов комиссий (заглушки) ---
    def _select_commission_types_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")])
        if file_selected:
            self.commission_types_file_entry.delete(0, tk.END)
            self.commission_types_file_entry.insert(0, file_selected)
            self.log_message(f"Выбран файл типов комиссий: {file_selected}", level="info")

    def _load_commission_types(self):
        file_path = self.commission_types_file_entry.get()
        if not file_path or not os.path.exists(file_path):
            self.log_message("Файл типов комиссий не выбран или не существует.", level="warning")
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите существующий файл типов комиссий.")
            return
        self.log_message(f"Загрузка типов комиссий из: {file_path}", level="info")
        # Здесь будет вызов self.commission_manager.load_commission_types(file_path)
        # Затем обновить self.commission_types_tree
        messagebox.showinfo("Информация", "Загрузка типов комиссий (пока не реализовано).")


    def _select_address_map_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")])
        if file_selected:
            self.address_map_file_entry.delete(0, tk.END)
            self.address_map_file_entry.insert(0, file_selected)
            self.log_message(f"Выбран файл сопоставления адресов: {file_selected}", level="info")

    def _load_address_map(self):
        file_path = self.address_map_file_entry.get()
        if not file_path or not os.path.exists(file_path):
            self.log_message("Файл сопоставления адресов не выбран или не существует.", level="warning")
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите существующий файл сопоставления адресов.")
            return
        self.log_message(f"Загрузка сопоставления адресов из: {file_path}", level="info")
        # Здесь будет вызов self.commission_manager.load_address_commission_map(file_path)
        # Затем обновить self.address_map_tree
        messagebox.showinfo("Информация", "Загрузка сопоставления адресов (пока не реализовано).")


    # --- Методы для работы с отчетами (заглушки) ---
    def _scan_reports(self):
        reports_folder = self.reports_folder_entry.get()
        if not reports_folder or not os.path.isdir(reports_folder):
            self.log_message("Папка с отчётами не выбрана или не существует.", level="error")
            messagebox.showerror("Ошибка", "Пожалуйста, выберите существующую папку с отчётами.")
            return

        self.log_message(f"Сканирование папки с отчётами: {reports_folder}", level="info")
        # Здесь будет вызов self.report_processor.scan_reports()
        # и обновление self.report_selection_combobox
        self.report_selection_combobox['values'] = ["Отчёт 1 (заглушка)", "Отчёт 2 (заглушка)"]
        self.report_selection_combobox.set("Отчёты просканированы (заглушка)")
        self.log_message("Отчёты просканированы (функция пока заглушка).", level="success")
        messagebox.showinfo("Информация", "Сканирование отчётов (пока не реализовано).")


    def _fill_selected_report(self):
        selected_report = self.report_selection_combobox.get()
        if selected_report == "Выберите отчёт" or not selected_report:
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите отчёт для заполнения.")
            return

        reports_folder = self.reports_folder_entry.get()
        data_folder = self.data_folder_entry.get()
        output_folder = self.output_folder_entry.get()

        if not all([reports_folder, data_folder, output_folder]) or \
           not all(map(os.path.isdir, [reports_folder, data_folder, output_folder])):
            messagebox.showerror("Ошибка", "Пожалуйста, укажите все необходимые папки.")
            return

        self.log_message(f"Заполнение выбранного отчёта: {selected_report}", level="info")
        # Здесь будет вызов self.report_processor.process_single_report(...)
        self.progress_bar['value'] = 50
        self.progress_label['text'] = f"Прогресс: Заполнение {selected_report}..."
        # Имитация работы
        self.after(1000, lambda: self._update_progress_and_log(100, "Заполнение выбранного отчёта завершено (заглушка).", "success"))


    def _fill_all_reports(self):
        reports_folder = self.reports_folder_entry.get()
        data_folder = self.data_folder_entry.get()
        output_folder = self.output_folder_entry.get()

        if not all([reports_folder, data_folder, output_folder]) or \
           not all(map(os.path.isdir, [reports_folder, data_folder, output_folder])):
            messagebox.showerror("Ошибка", "Пожалуйста, укажите все необходимые папки.")
            return

        self.log_message("Заполнение ВСЕХ отчётов...", level="info")
        # Здесь будет вызов self.report_processor.process_all_reports()
        self.progress_bar['value'] = 0
        self.progress_label['text'] = "Прогресс: Начинаю заполнение всех отчётов..."
        # Имитация работы
        self.after(200, lambda: self._update_progress_and_log(25, "Обрабатываю Отчёт 1...", "info"))
        self.after(1000, lambda: self._update_progress_and_log(75, "Обрабатываю Отчёт 2...", "info"))
        self.after(2000, lambda: self._update_progress_and_log(100, "Заполнение ВСЕХ отчётов завершено (заглушка).", "success"))


    def _update_progress_and_log(self, value, message, level="info"):
        """Вспомогательная функция для обновления прогресса и логов."""
        self.progress_bar['value'] = value
        self.progress_label['text'] = f"Прогресс: {message}"
        self.log_message(message, level)


    # --- Методы для управления комиссиями (заглушки) ---
    def _add_commission_type(self):
        self.log_message("Добавление нового типа комиссии (пока не реализовано).", level="info")
        messagebox.showinfo("Информация", "Добавление типа комиссии (пока не реализовано).")

    def _edit_commission_type(self):
        selected_item = self.commission_types_tree.focus()
        if not selected_item:
            messagebox.showwarning("Предупреждение", "Выберите тип комиссии для редактирования.")
            return
        self.log_message("Редактирование выбранного типа комиссии (пока не реализовано).", level="info")
        messagebox.showinfo("Информация", "Редактирование типа комиссии (пока не реализовано).")

    def _delete_commission_type(self):
        selected_items = self.commission_types_tree.selection()
        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите один или несколько типов комиссий для удаления.")
            return
        self.log_message(f"Удаление {len(selected_items)} выбранных типов комиссий (пока не реализовано).", level="info")
        messagebox.showinfo("Информация", "Удаление типов комиссий (пока не реализовано).")

    def _export_commission_types(self):
        self.log_message("Экспорт типов комиссий (пока не реализовано).", level="info")
        messagebox.showinfo("Информация", "Экспорт типов комиссий (пока не реализовано).")


    def _add_address_map(self):
        self.log_message("Добавление нового сопоставления адреса (пока не реализовано).", level="info")
        messagebox.showinfo("Информация", "Добавление сопоставления адреса (пока не реализовано).")

    def _edit_address_map(self):
        selected_item = self.address_map_tree.focus()
        if not selected_item:
            messagebox.showwarning("Предупреждение", "Выберите сопоставление адреса для редактирования.")
            return
        self.log_message("Редактирование выбранного сопоставления адреса (пока не реализовано).", level="info")
        messagebox.showinfo("Информация", "Редактирование сопоставления адреса (пока не реализовано).")

    def _delete_address_map(self):
        selected_items = self.address_map_tree.selection()
        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите одно или несколько сопоставлений адресов для удаления.")
            return
        self.log_message(f"Удаление {len(selected_items)} выбранных сопоставлений адресов (пока не реализовано).", level="info")
        messagebox.showinfo("Информация", "Удаление сопоставлений адресов (пока не реализовано).")

    def _export_address_map(self):
        self.log_message("Экспорт сопоставлений адресов (пока не реализовано).", level="info")
        messagebox.showinfo("Информация", "Экспорт сопоставлений адресов (пока не реализовано).")


if __name__ == "__main__":
    app = ReportFillerApp()
    app.mainloop()



Что этот код делает:
Инициализация Tkinter: Создает главное окно приложения с заголовком и начальными размерами.
_create_widgets(): Эта функция отвечает за создание основных элементов интерфейса:
ttk.Notebook: Система вкладок, которая содержит "Заполнение отчётов" и "Управление комиссиями".
Область логов: Внизу окна находится текстовое поле для вывода системных сообщений, ошибок и статуса операций. Оно оснащено скроллбаром и кнопкой "Сохранить журнал".
Строка состояния: Еще ниже располагается строка состояния для кратких уведомлений.
_setup_report_filling_tab(): Настраивает элементы управления для первой вкладки:
Поля ввода для путей к папкам (отчёты, данные, выходная папка) с кнопками "Выбрать".
Кнопки для сканирования отчётов, выпадающий список для выбора отчёта, и кнопки для заполнения выбранного/всех отчётов.
Индикатор прогресса (ttk.Progressbar и ttk.Label).
_setup_commission_management_tab(): Настраивает элементы для второй вкладки:
Поля ввода для путей к файлам с типами комиссий и сопоставлением "Адрес-Район-Газ" с кнопками "Выбрать" и "Загрузить".
Заглушки для ttk.Treeview для отображения данных о комиссиях и их сопоставлениях.
Заглушки для кнопок управления (Добавить, Редактировать, Удалить, Экспорт).
log_message(): Утилитарная функция для вывода сообщений в текстовое поле логов с меткой времени и цветовой индикацией (информация, предупреждение, ошибка, успех). Также обновляет строку состояния.
_save_log_to_file(): Позволяет пользователю сохранить содержимое журнала в текстовый файл.
_load_initial_settings() и _on_closing(): Методы-заглушки для интеграции с config_manager.py. Они будут вызывать методы сохранения/загрузки настроек при старте и закрытии приложения.
Методы-заглушки (_select_..._folder, _select_..._file, _scan_reports, _fill_..._report, _add_..., _edit_..., _delete_..., _export_...): Эти методы привязаны к кнопкам и будут вызывать соответствующие функции в других модулях после их реализации. Сейчас они просто выводят сообщения в лог или показывают messagebox.
if __name__ == "__main__":: Стандартная точка входа для запуска приложения.
Что делать дальше:
Создайте файлы: Убедитесь, что вы создали все пустые файлы и папки, как было описано в предыдущем шаге (например, config_manager.py, utils.py, commission_manager.py, report_processor.py).
Запустите этот main_app.py: Вы сможете увидеть базовое окно приложения с вкладками и логами. Кнопки пока будут выдавать сообщения "пока не реализовано".
Переходите к config_manager.py: Это следующий шаг в нашем плане. Мы реализуем его, а затем интегрируем с этим main_app.py, чтобы настроить сохранение/загрузку путей.
Дайте знать, когда этот код будет у вас и вы будете готовы перейти к реализации config_manager.py!
Источники
1. https://github.com/OmarElsiry/yt-bulk-downloader
2. https://github.com/respunn/re-render


1. config_manager.py
Этот модуль будет отвечать за сохранение и загрузку настроек приложения, таких как пути к папкам, регулярные выражения и т.д.

Python


import configparser
import os

class ConfigManager:
    """
    Управляет конфигурацией приложения: сохраняет и загружает пути к папкам,
    шаблоны регулярных выражений и другие настройки.
    """
    def __init__(self, config_file='config.ini'):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self._load_config()

    def _load_config(self):
        """Загружает настройки из файла конфигурации."""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding='utf-8')
        else:
            # Установка значений по умолчанию, если файл не существует
            self._set_default_config()
            self.save_config() # Сохраняем файл с дефолтными значениями

    def _set_default_config(self):
        """Устанавливает значения по умолчанию для всех настроек."""
        if 'Paths' not in self.config:
            self.config['Paths'] = {
                'reports_folder': '',
                'data_folder': '',
                'output_folder': '',
                'commission_types_file': '',
                'address_map_file': ''
            }
        if 'Regex' not in self.config:
            self.config['Regex'] = {
                'address_extraction_pattern': r'\(([^)]+)\)', # Пример: извлекает текст в скобках
                'gas_detection_keywords': 'газ,газоснабжение,газопровод', # Ключевые слова для поиска газа в отчете
                'gas_detection_cell_offset_x': '0', # Смещение по X от газового ключевого слова для поиска "Да/Нет"
                'gas_detection_cell_offset_y': '1'  # Смещение по Y от газового ключевого слова для поиска "Да/Нет"
            }
        if 'FieldMapping' not in self.config:
            self.config['FieldMapping'] = {} # Для ручных сопоставлений полей

    def save_config(self):
        """Сохраняет текущие настройки в файл."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as configfile:
                self.config.write(configfile)
        except Exception as e:
            print(f"Ошибка при сохранении конфигурации: {e}") # В реальном приложении логировать

    def get(self, section, key, default=''):
        """Получает значение настройки по секции и ключу."""
        return self.config.get(section, key, fallback=default)

    def set(self, section, key, value):
        """Устанавливает значение настройки по секции и ключу."""
        if section not in self.config:
            self.config[section] = {}
        self.config[section][key] = str(value)



2. utils.py
Этот модуль будет содержать вспомогательные функции, которые используются в разных частях приложения, такие как извлечение адреса, нечёткое сравнение строк и поиск ячеек в Excel.

Python


import re
from fuzzywuzzy import fuzz
from openpyxl.utils import get_column_letter

class Utils:
    """
    Содержит вспомогательные функции для различных операций в приложении.
    """
    def __init__(self, config_manager=None):
        self.config_manager = config_manager

    def extract_address_from_filename(self, filename):
        """
        Извлекает адрес из имени файла с помощью регулярного выражения из конфига.
        """
        if not self.config_manager:
            # Fallback для тестирования без ConfigManager
            pattern = r'\(([^)]+)\)'
            # print("Warning: ConfigManager not provided to Utils. Using default regex.")
        else:
            pattern = self.config_manager.get('Regex', 'address_extraction_pattern')

        match = re.search(pattern, filename)
        if match:
            return match.group(1).strip()
        return None

    def fuzzy_match(self, search_text, candidates_list, threshold=80):
        """
        Находит наилучшее нечёткое совпадение `search_text` в `candidates_list`.
        Возвращает (совпадение, score), если score выше порога, иначе (None, 0).
        """
        if not candidates_list:
            return None, 0

        best_match = None
        highest_score = 0

        for candidate in candidates_list:
            score = fuzz.ratio(search_text.lower(), candidate.lower())
            if score > highest_score:
                highest_score = score
                best_match = candidate

        if highest_score >= threshold:
            return best_match, highest_score
        return None, 0

    def find_cell_by_keywords(self, worksheet, keywords, search_range=None):
        """
        Находит ячейку, содержащую одно из ключевых слов (регистронезависимо).
        `keywords` - список строк.
        `search_range` - кортеж (min_row, min_col, max_row, max_col) для ограничения поиска.
        Возвращает (row, col) найденной ячейки или None.
        """
        # Преобразуем ключевые слова в нижний регистр для регистронезависимого поиска
        lower_keywords = [k.lower() for k in keywords]

        if search_range:
            min_row, min_col, max_row, max_col = search_range
        else:
            min_row, min_col, max_row, max_col = 1, 1, worksheet.max_row, worksheet.max_column

        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                cell_value = worksheet.cell(row=row, column=col).value
                if cell_value is not None and isinstance(cell_value, str):
                    for keyword in lower_keywords:
                        if keyword in cell_value.lower():
                            return row, col # Возвращаем 1-based индексы
        return None

    def find_value_cell(self, worksheet, start_row, start_col, max_search_distance=5):
        """
        Пытается найти подходящую ячейку для записи значения,
        начиная от start_cell и двигаясь вправо, затем вниз.
        Учитывает объединенные ячейки.
        Возвращает (row, col) ячейки для записи или None.
        """
        # Проверяем ячейки справа
        for col_offset in range(1, max_search_distance + 1):
            target_col = start_col + col_offset
            if target_col > worksheet.max_column:
                break
            cell = worksheet.cell(row=start_row, column=target_col)
            if self._is_suitable_for_value(cell, worksheet):
                return cell.row, cell.column

        # Если справа не нашли, проверяем ячейки снизу
        for row_offset in range(1, max_search_distance + 1):
            target_row = start_row + row_offset
            if target_row > worksheet.max_row:
                break
            cell = worksheet.cell(row=target_row, column=start_col)
            if self._is_suitable_for_value(cell, worksheet):
                return cell.row, cell.column

        return None

    def _is_suitable_for_value(self, cell, worksheet):
        """
        Проверяет, подходит ли ячейка для записи значения.
        Ячейка подходит, если она пуста или содержит только пробелы,
        или если она является частью объединенной ячейки,
        и эта объединенная ячейка пуста или содержит только пробелы.
        """
        # Проверяем, является ли ячейка частью объединенной области
        for merged_cell_range in worksheet.merged_cells.ranges:
            if cell.coordinate in merged_cell_range:
                # Если это объединенная ячейка, проверяем ее верхнюю левую ячейку
                # т.к. только она хранит значение для объединенной области
                min_col_letter = get_column_letter(merged_cell_range.min_col)
                top_left_cell_coord = f"{min_col_letter}{merged_cell_range.min_row}"
                top_left_cell = worksheet[top_left_cell_coord]
                return top_left_cell.value is None or str(top_left_cell.value).strip() == ''

        # Если ячейка не объединена, проверяем ее значение
        return cell.value is None or str(cell.value).strip() == ''

    def get_cell_value(self, worksheet, row, col):
        """Безопасно получает значение ячейки, учитывая объединенные ячейки."""
        cell = worksheet.cell(row=row, column=col)
        for merged_range in worksheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                # Если ячейка находится в объединенном диапазоне, возвращаем значение из верхней левой ячейки
                return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
        return cell.value

    def get_boolean_from_text(self, text):
        """Преобразует текстовое значение в булево (для наличия газа)."""
        if text is None:
            return False
        text_lower = str(text).strip().lower()
        if text_lower in ('да', 'true', 'есть', 'yes', '+'):
            return True
        return False



3. commission_manager.py
Этот модуль будет управлять всеми данными, касающимися комиссий: их составом по районам/газу и сопоставлением адресов.

Python


import pandas as pd
import os
import re

class CommissionManager:
    """
    Управляет данными о комиссиях: их составом по районам и наличию газа,
    а также сопоставлением адресов с типами комиссий.
    """
    def __init__(self, config_manager=None, log_callback=None):
        self.config_manager = config_manager
        self.log_message = log_callback if log_callback else print

        # commission_types: {(район, has_gas): {role: name, role_pos: position, ...}}
        self.commission_types = {}
        # address_to_commission_map: {address: (район, has_gas)}
        self.address_to_commission_map = {}

        self._load_initial_data()

    def _load_initial_data(self):
        """Загружает данные комиссий и сопоставлений при инициализации."""
        if self.config_manager:
            types_file = self.config_manager.get('Paths', 'commission_types_file')
            if types_file and os.path.exists(types_file):
                self.load_commission_types(types_file)

            map_file = self.config_manager.get('Paths', 'address_map_file')
            if map_file and os.path.exists(map_file):
                self.load_address_commission_map(map_file)

    def load_commission_types(self, file_path):
        """
        Загружает типы комиссий из Excel/CSV файла.
        Ожидаемый формат: колонки "Район", "Газ", и далее колонки с ролями
        и должностями (например, "Председатель", "Должность Председателя", "Член 1", и т.д.).
        """
        if not os.path.exists(file_path):
            self.log_message(f"Файл типов комиссий не найден: {file_path}", level="error")
            return False

        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8')
            else:
                df = pd.read_excel(file_path)

            new_commission_types = {}
            for index, row in df.iterrows():
                try:
                    region = str(row['Район']).strip()
                    has_gas_str = str(row['Газ']).strip().lower()
                    has_gas = (has_gas_str == 'да' or has_gas_str == 'true' or has_gas_str == 'есть')

                    commission_key = (region, has_gas)
                    composition = {}
                    for col in df.columns:
                        if col not in ['Район', 'Газ'] and pd.notna(row[col]):
                            composition[col.strip()] = str(row[col]).strip()
                    new_commission_types[commission_key] = composition
                except KeyError as ke:
                    self.log_message(f"Ошибка в файле типов комиссий: отсутствует обязательная колонка {ke} в строке {index+2}. Пропустил строку.", level="warning")
                except Exception as e:
                    self.log_message(f"Ошибка обработки строки {index+2} в файле типов комиссий: {e}. Пропустил строку.", level="error")

            self.commission_types = new_commission_types
            self.log_message(f"Загружено {len(self.commission_types)} типов комиссий из {file_path}", level="success")
            return True
        except Exception as e:
            self.log_message(f"Ошибка при загрузке файла типов комиссий {file_path}: {e}", level="error")
            return False

    def load_address_commission_map(self, file_path):
        """
        Загружает сопоставление адресов с типами комиссий из Excel/CSV файла.
        Ожидаемый формат: колонки "Адрес", "Район", "Газ".
        """
        if not os.path.exists(file_path):
            self.log_message(f"Файл сопоставления адресов не найден: {file_path}", level="error")
            return False

        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8')
            else:
                df = pd.read_excel(file_path)

            new_address_map = {}
            for index, row in df.iterrows():
                try:
                    address = str(row['Адрес']).strip()
                    region = str(row['Район']).strip()
                    has_gas_str = str(row['Газ']).strip().lower()
                    has_gas = (has_gas_str == 'да' or has_gas_str == 'true' or has_gas_str == 'есть')

                    if address in new_address_map:
                        self.log_message(f"Дубликат адреса '{address}' в файле сопоставления адресов (строка {index+2}). Будет использована последняя запись.", level="warning")
                    new_address_map[address] = (region, has_gas)
                except KeyError as ke:
                    self.log_message(f"Ошибка в файле сопоставления адресов: отсутствует обязательная колонка {ke} в строке {index+2}. Пропустил строку.", level="warning")
                except Exception as e:
                    self.log_message(f"Ошибка обработки строки {index+2} в файле сопоставления адресов: {e}. Пропустил строку.", level="error")

            self.address_to_commission_map = new_address_map
            self.log_message(f"Загружено {len(self.address_to_commission_map)} сопоставлений адресов из {file_path}", level="success")
            return True
        except Exception as e:
            self.log_message(f"Ошибка при загрузке файла сопоставления адресов {file_path}: {e}", level="error")
            return False

    def get_commission_composition(self, address, has_gas):
        """
        Возвращает состав комиссии для данного адреса и наличия газа.
        Сначала пытается найти по адресу, затем по типу комиссии из map.
        """
        # Пытаемся найти район по адресу
        if address in self.address_to_commission_map:
            region, gas_status_from_map = self.address_to_commission_map[address]
            # Используем gas_status_from_map, если он надежнее, или переданный has_gas
            # В данном случае, используем переданный has_gas, так как он извлечен из самого отчета.
            commission_key = (region, has_gas)
            if commission_key in self.commission_types:
                return self.commission_types[commission_key]
            else:
                self.log_message(f"Не найден тип комиссии для {commission_key} по адресу {address}. Проверьте файл типов комиссий.", level="warning")
        else:
            self.log_message(f"Адрес '{address}' не найден в файле сопоставления адресов.", level="warning")
        return None

    def get_all_commission_types_for_display(self):
        """Возвращает список всех типов комиссий в формате для отображения в Treeview."""
        display_data = []
        for (region, has_gas), composition in self.commission_types.items():
            row_data = {
                "Район": region,
                "Газ": "Да" if has_gas else "Нет"
            }
            # Добавляем все роли и должности из состава
            for role, name in composition.items():
                row_data[role] = name
            display_data.append(row_data)
        return display_data

    def get_all_address_maps_for_display(self):
        """Возвращает список всех сопоставлений адресов в формате для отображения в Treeview."""
        display_data = []
        for address, (region, has_gas) in self.address_to_commission_map.items():
            display_data.append({
                "Адрес": address,
                "Район": region,
                "Газ": "Да" if has_gas else "Нет"
            })
        return display_data

    # --- Методы для CRUD операций (пока заглушки, будут расширены с UI) ---
    def add_commission_type(self, region, has_gas, composition):
        """Добавляет новый тип комиссии."""
        commission_key = (region, has_gas)
        if commission_key in self.commission_types:
            self.log_message(f"Тип комиссии для {commission_key} уже существует. Обновляю.", level="warning")
        self.commission_types[commission_key] = composition
        self.log_message(f"Добавлен/обновлен тип комиссии: {commission_key}", level="info")

    def delete_commission_type(self, region, has_gas):
        """Удаляет тип комиссии."""
        commission_key = (region, has_gas)
        if commission_key in self.commission_types:
            del self.commission_types[commission_key]
            self.log_message(f"Удален тип комиссии: {commission_key}", level="info")
            return True
        self.log_message(f"Тип комиссии для {commission_key} не найден для удаления.", level="warning")
        return False

    def add_address_map(self, address, region, has_gas):
        """Добавляет новое сопоставление адреса."""
        if address in self.address_to_commission_map:
            self.log_message(f"Сопоставление для адреса '{address}' уже существует. Обновляю.", level="warning")
        self.address_to_commission_map[address] = (region, has_gas)
        self.log_message(f"Добавлено/обновлено сопоставление для адреса: {address} -> ({region}, {has_gas})", level="info")

    def delete_address_map(self, address):
        """Удаляет сопоставление адреса."""
        if address in self.address_to_commission_map:
            del self.address_to_commission_map[address]
            self.log_message(f"Удалено сопоставление для адреса: {address}", level="info")
            return True
        self.log_message(f"Сопоставление для адреса '{address}' не найдено для удаления.", level="warning")
        return False

    def export_commission_types(self, file_path):
        """Экспортирует текущие типы комиссий в Excel/CSV."""
        try:
            data = []
            # Собираем все возможные роли/должности, чтобы создать столбцы DataFrame
            all_roles = set()
            for comp in self.commission_types.values():
                all_roles.update(comp.keys())

            for (region, has_gas), composition in self.commission_types.items():
                row = {"Район": region, "Газ": "Да" if has_gas else "Нет"}
                row.update(composition)
                data.append(row)

            df = pd.DataFrame(data)
            # Убедимся, что ключевые столбцы идут первыми
            cols = ["Район", "Газ"] + sorted([c for c in df.columns if c not in ["Район", "Газ"]])
            df = df[cols]

            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8')
            else:
                df.to_excel(file_path, index=False)
            self.log_message(f"Типы комиссий успешно экспортированы в {file_path}", level="success")
            return True
        except Exception as e:
            self.log_message(f"Ошибка при экспорте типов комиссий: {e}", level="error")
            return False

    def export_address_map(self, file_path):
        """Экспортирует текущие сопоставления адресов в Excel/CSV."""
        try:
            data = []
            for address, (region, has_gas) in self.address_to_commission_map.items():
                data.append({"Адрес": address, "Район": region, "Газ": "Да" if has_gas else "Нет"})
            df = pd.DataFrame(data)
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8')
            else:
                df.to_excel(file_path, index=False)
            self.log_message(f"Сопоставления адресов успешно экспортированы в {file_path}", level="success")
            return True
        except Exception as e:
            self.log_message(f"Ошибка при экспорте сопоставлений адресов: {e}", level="error")
            return False



4. report_processor.py
Этот модуль будет содержать основную логику обработки Excel-отчётов: чтение данных, поиск полей, запись значений, вставку строк и создание отчёта о выполнении.

Python


import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange
import re

class ReportProcessor:
    """
    Обрабатывает Excel-отчёты: сканирует, извлекает данные, заполняет поля,
    вставляет строки для членов комиссии (если требуется) и генерирует отчёты о выполнении.
    """
    def __init__(self, config_manager, commission_manager, utils, log_callback=None):
        self.config_manager = config_manager
        self.commission_manager = commission_manager
        self.utils = utils
        self.log_message = log_callback if log_callback else print
        self.report_files = [] # Список найденных файлов отчетов

        # Настраиваемые параметры для поиска газа в отчете
        self.gas_detection_keywords = [
            k.strip() for k in self.config_manager.get('Regex', 'gas_detection_keywords').split(',')
        ]
        self.gas_cell_offset_x = int(self.config_manager.get('Regex', 'gas_detection_cell_offset_x'))
        self.gas_cell_offset_y = int(self.config_manager.get('Regex', 'gas_detection_cell_offset_y'))

        # Для хранения ручных сопоставлений полей (пока не реализовано в UI)
        self.manual_field_mappings = self.config_manager.get('FieldMapping', 'manual_mappings', fallback={})


    def scan_reports(self, reports_folder):
        """Сканирует указанную папку на наличие файлов отчётов Excel."""
        self.report_files = []
        if not os.path.isdir(reports_folder):
            self.log_message(f"Папка с отчётами не найдена: {reports_folder}", level="error")
            return []

        for root, _, files in os.walk(reports_folder):
            for file in files:
                if file.lower().endswith(('.xlsx', '.xls')):
                    full_path = os.path.join(root, file)
                    self.report_files.append(full_path)
        self.log_message(f"Найдено {len(self.report_files)} файлов отчётов в {reports_folder}", level="info")
        return self.report_files

    def process_all_reports(self, reports_folder, data_folder, output_folder, update_progress_callback=None):
        """Обрабатывает все найденные отчёты."""
        if not self.report_files:
            self.log_message("Нет отчётов для обработки. Сначала просканируйте папку.", level="warning")
            return

        results = []
        total_reports = len(self.report_files)
        for i, report_path in enumerate(self.report_files):
            file_name = os.path.basename(report_path)
            self.log_message(f"Обработка отчёта {i+1}/{total_reports}: {file_name}", level="info")
            if update_progress_callback:
                update_progress_callback(i + 1, total_reports, file_name)

            report_result = self.process_single_report(report_path, data_folder, output_folder)
            results.append(report_result)

        self._generate_processing_report(output_folder, results)
        self.log_message("Обработка всех отчётов завершена.", level="success")
        return results

    def process_single_report(self, report_path, data_folder, output_folder):
        """
        Обрабатывает один файл отчёта: извлекает данные, заполняет поля,
        вставляет строки и сохраняет.
        Возвращает словарь с результатом обработки.
        """
        file_name = os.path.basename(report_path)
        address = self.utils.extract_address_from_filename(file_name)
        if not address:
            self.log_message(f"Не удалось извлечь адрес из имени файла '{file_name}'. Пропускаю.", level="error")
            return {
                "file": file_name,
                "address": "Неизвестен",
                "status": "Ошибка",
                "message": "Не удалось извлечь адрес из имени файла",
                "filled_fields": 0,
                "missing_data_fields": []
            }

        data_file_path = self._find_data_file_for_address(data_folder, address)
        if not data_file_path:
            self.log_message(f"Не найден файл данных для адреса '{address}'. Пропускаю '{file_name}'.", level="warning")
            return {
                "file": file_name,
                "address": address,
                "status": "Ошибка",
                "message": "Не найден файл данных для адреса",
                "filled_fields": 0,
                "missing_data_fields": []
            }

        try:
            data_from_file = self._read_data_file(data_file_path)
        except Exception as e:
            self.log_message(f"Ошибка чтения файла данных {data_file_path}: {e}. Пропускаю '{file_name}'.", level="error")
            return {
                "file": file_name,
                "address": address,
                "status": "Ошибка",
                "message": f"Ошибка чтения файла данных: {e}",
                "filled_fields": 0,
                "missing_data_fields": []
            }

        # Теперь объединяем данные из файла и данные комиссии
        full_data_for_report = data_from_file.copy()

        try:
            workbook = load_workbook(report_path)
            sheet = workbook.active # Или выбрать конкретный лист, если нужно
            self.log_message(f"Открыт отчёт: {file_name}", level="info")

            # 1. Определение наличия газа в отчёте
            has_gas_in_report = self._detect_gas_in_report(sheet)
            self.log_message(f"Для '{address}' обнаружено газоснабжение: {'Да' if has_gas_in_report else 'Нет'}", level="info")

            # 2. Получение состава комиссии
            commission_composition = self.commission_manager.get_commission_composition(address, has_gas_in_report)
            if commission_composition:
                self.log_message(f"Состав комиссии для '{address}' ({'Газ' if has_gas_in_report else 'Без газа'}) найден.", level="info")
                full_data_for_report.update(commission_composition) # Добавляем данные комиссии к общим данным
            else:
                self.log_message(f"Не удалось найти состав комиссии для '{address}' ({'Газ' if has_gas_in_report else 'Без газа'}).", level="warning")


            filled_count = 0
            missing_fields = []
            matched_cells = {} # Для избежания повторного заполнения одной ячейки

            # Если есть "Ресурсник" и газа не было, возможно нужно добавить строки
            # Эта логика должна быть более сложной и зависеть от шаблона отчета
            # Для примера, предположим, что ресурсник всегда добавляется последним
            # А также где в отчете искать место для вставки.
            # Пока оставим простой поиск и запись
            if has_gas_in_report and 'Ресурсник' in full_data_for_report:
                # Находим строку, после которой нужно вставить нового члена комиссии
                # Например, ищем "Член комиссии", и после него вставляем "Ресурсника"
                # Это очень сильно зависит от шаблона отчета!
                # Для примера, ищем "Член комиссии"
                member_row_col = self.utils.find_cell_by_keywords(sheet, ["Член комиссии", "Член"])
                if member_row_col:
                    insert_row = member_row_col[0] + 1 # Вставляем после члена комиссии
                    self._insert_rows(sheet, insert_row, 1) # Вставляем одну строку
                    self.log_message(f"Вставлена строка для ресурсника после строки {insert_row-1}.", level="info")
                else:
                    self.log_message("Не удалось найти 'Член комиссии' для вставки строки ресурсника. Заполнение будет произведено в существующие поля.", level="warning")


            # Ищем и заполняем поля в отчете
            for data_field, data_value in full_data_for_report.items():
                found_cell_coords = None
                # Сначала пытаемся найти по точному совпадению или ручному маппингу
                if data_field in self.manual_field_mappings:
                    target_coord = self.manual_field_mappings[data_field]
                    if re.match(r"^[A-Z]+\d+$", target_coord): # Проверка формата A1
                        found_cell_coords = (sheet[target_coord].row, sheet[target_coord].column)
                else:
                    # Ищем поле нечётко
                    # Предполагаем, что ключевые слова для поиска полей могут быть в заголовках столбцов или рядом
                    # Проходим по всем ячейкам листа в поисках совпадения
                    for row_idx in range(1, sheet.max_row + 1):
                        for col_idx in range(1, sheet.max_column + 1):
                            cell = sheet.cell(row=row_idx, column=col_idx)
                            cell_value = str(cell.value).strip() if cell.value is not None else ""

                            matched_keyword, score = self.utils.fuzzy_match(data_field, [cell_value])
                            if score >= 85: # Высокий порог для прямых совпадений
                                # Нашли потенциальное поле, теперь ищем ячейку для значения
                                value_row, value_col = self.utils.find_value_cell(sheet, cell.row, cell.column)
                                if value_row and value_col:
                                    # Проверим, что мы не заполняем ту же ячейку снова, если поле уже было найдено
                                    if (value_row, value_col) not in matched_cells.values():
                                        found_cell_coords = (value_row, value_col)
                                        break
                        if found_cell_coords:
                            break

                if found_cell_coords:
                    row, col = found_cell_coords
                    current_cell_value = self.utils.get_cell_value(sheet, row, col)
                    if current_cell_value is None or str(current_cell_value).strip() == '':
                        # Записываем значение
                        sheet.cell(row=row, column=col, value=data_value)
                        matched_cells[data_field] = (row, col)
                        filled_count += 1
                        # self.log_message(f"  Заполнено поле '{data_field}' в ячейке {get_column_letter(col)}{row} значением '{data_value}'", level="debug")
                    else:
                        self.log_message(f"  Ячейка {get_column_letter(col)}{row} для поля '{data_field}' уже содержит значение: '{current_cell_value}'. Пропускаю.", level="warning")
                else:
                    missing_fields.append(data_field)
                    self.log_message(f"  Не найдено подходящее место для заполнения поля '{data_field}'", level="warning")


            # Сохранение заполненного отчёта
            output_file_name = f"{os.path.splitext(file_name)[0]}_FILLED.xlsx"
            output_path = os.path.join(output_folder, output_file_name)
            workbook.save(output_path)
            self.log_message(f"Отчёт '{file_name}' успешно заполнен и сохранён как '{output_file_name}'", level="success")

            return {
                "file": file_name,
                "address": address,
                "status": "Успешно",
                "message": "Отчёт успешно заполнен",
                "filled_fields": filled_count,
                "missing_data_fields": missing_fields,
                "has_gas": has_gas_in_report
            }

        except Exception as e:
            self.log_message(f"Критическая ошибка при обработке отчёта '{file_name}': {e}", level="error")
            return {
                "file": file_name,
                "address": address,
                "status": "Ошибка",
                "message": f"Критическая ошибка: {e}",
                "filled_fields": 0,
                "missing_data_fields": list(full_data_for_report.keys()) # Все поля могли быть незаполнены
            }

    def _find_data_file_for_address(self, base_data_folder, address):
        """
        Ищет файл данных для данного адреса в подпапках base_data_folder.
        """
        # Сначала ищем папку, соответствующую адресу
        for root, dirs, files in os.walk(base_data_folder):
            if address in dirs:
                address_data_folder = os.path.join(root, address)
                # Теперь ищем Excel/CSV файл в этой папке
                for f in os.listdir(address_data_folder):
                    if f.lower().endswith(('.xlsx', '.xls', '.csv')):
                        return os.path.join(address_data_folder, f)
        return None

    def _read_data_file(self, file_path):
        """
        Читает данные из Excel или CSV файла и возвращает их в виде словаря.
        Предполагается, что данные находятся в первом листе, и формат:
        первый столбец - название поля, второй столбец - значение.
        """
        if file_path.lower().endswith('.csv'):
            df = pd.read_csv(file_path, header=None, encoding='utf-8')
        else:
            df = pd.read_excel(file_path, header=None, sheet_name=0) # Читаем первый лист

        data_dict = {}
        for index, row in df.iterrows():
            if len(row) >= 2 and pd.notna(row[0]) and pd.notna(row[1]):
                key = str(row[0]).strip()
                value = str(row[1]).strip()
                data_dict[key] = value
        return data_dict

    def _detect_gas_in_report(self, worksheet):
        """
        Пытается определить наличие газоснабжения в отчете по ключевым словам.
        Возвращает True, если найдено, False иначе.
        """
        for keyword in self.gas_detection_keywords:
            found_pos = self.utils.find_cell_by_keywords(worksheet, [keyword])
            if found_pos:
                # Если ключевое слово найдено, проверяем смещенную ячейку
                target_row = found_pos[0] + self.gas_cell_offset_y
                target_col = found_pos[1] + self.gas_cell_offset_x
                
                if target_row > 0 and target_col > 0: # Убедимся, что индексы валидны
                    cell_value = self.utils.get_cell_value(worksheet, target_row, target_col)
                    if self.utils.get_boolean_from_text(cell_value):
                        return True
        return False

    def _insert_rows(self, worksheet, start_row, num_rows):
        """
        Вставляет `num_rows` пустых строк, начиная со `start_row`.
        Корректирует объединенные ячейки, которые пересекаются с точкой вставки.
        """
        # Сдвигаем все строки ниже start_row
        worksheet.insert_rows(start_row, num_rows)

        # Корректируем объединенные ячейки
        # Важно: openpyxl.insert_rows обычно сам корректирует объединенные ячейки,
        # но сложные случаи могут потребовать ручной обработки.
        # Для базового случая insert_rows должен сработать.
        # Дополнительная проверка на всякий случай
        for merged_range_str in list(worksheet.merged_cells.ranges): # Создаем копию списка, т.к. будем изменять
            merged_range = CellRange(merged_range_str)
            if merged_range.min_row >= start_row:
                # Если объединенная ячейка находится полностью ниже точки вставки, просто сдвигаем
                new_min_row = merged_range.min_row + num_rows
                new_max_row = merged_range.max_row + num_rows
                new_range = CellRange(min_col=merged_range.min_col, max_col=merged_range.max_col,
                                      min_row=new_min_row, max_row=new_max_row)
                worksheet.merged_cells.remove(merged_range_str)
                worksheet.merged_cells.add(str(new_range))
            elif merged_range.min_row < start_row <= merged_range.max_row:
                # Если точка вставки внутри объединенной ячейки, расширяем ее
                new_max_row = merged_range.max_row + num_rows
                new_range = CellRange(min_col=merged_range.min_col, max_col=merged_range.max_col,
                                      min_row=merged_range.min_row, max_row=new_max_row)
                worksheet.merged_cells.remove(merged_range_str)
                worksheet.merged_cells.add(str(new_range))


    def _generate_processing_report(self, output_folder, results):
        """Генерирует Excel-отчёт о результатах обработки."""
        report_data = []
        for res in results:
            report_data.append({
                "Файл отчёта": res["file"],
                "Адрес": res["address"],
                "Статус": res["status"],
                "Сообщение": res["message"],
                "Заполнено полей": res["filled_fields"],
                "Незаполненные поля (из данных)": ", ".join(res["missing_data_fields"]),
                "Наличие газа в отчёте": "Да" if res.get("has_gas") else "Нет"
            })

        df = pd.DataFrame(report_data)
        report_file_name = f"Отчёт_об_обработке_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        report_path = os.path.join(output_folder, report_file_name)

        try:
            df.to_excel(report_path, index=False)
            self.log_message(f"Отчёт об обработке успешно создан: {report_path}", level="success")
        except Exception as e:
            self.log_message(f"Ошибка при создании отчёта об обработке: {e}", level="error")



5. Интеграция в main_app.py
Теперь, когда все модули созданы, нам нужно обновить main_app.py, чтобы он импортировал и использовал эти классы.

Python


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
    """
    Главный класс приложения для заполнения отчётов и управления комиссиями.
    """
    def __init__(self):
        super().__init__()
        self.title("Заполнение паспортов готовности МКД к ОЗП")
        self.geometry("1000x700")
        self.minsize(800, 600)

        # Инициализация менеджеров
        self.config_manager = ConfigManager()
        self.utils = Utils(self.config_manager) # Передаем ConfigManager в Utils
        self.commission_manager = CommissionManager(self.config_manager, self.log_message)
        self.report_processor = ReportProcessor(self.config_manager, self.commission_manager, self.utils, self.log_message)

        self._create_widgets()
        self._setup_logging()
        self._load_initial_settings() # Загрузка настроек при запуске
        self._populate_commission_trees() # Обновить таблицы комиссий после загрузки данных

        # Привязка функции сохранения настроек к закрытию окна
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
        # Этот метод будет открывать модальное окно для ввода данных
        # Пока просто заглушка
        self.log_message("Добавление нового типа комиссии (не реализовано в UI).", level="info")
        messagebox.showinfo("Информация", "Функция добавления типа комиссии будет реализована позже с модальным окном.")
        # Пример использования:
        # self.commission_manager.add_commission_type("Новый Район", True, {"Председатель": "Тест", "Должность Председателя": "Тестовая должность"})
        # self._populate_commission_types_tree()

    def _edit_commission_type(self):
        selected_item = self.commission_types_tree.focus()
        if not selected_item:
            messagebox.showwarning("Предупреждение", "Выберите тип комиссии для редактирования.")
            return
        
        # Получаем значения из выбранной строки
        values = self.commission_types_tree.item(selected_item, 'values')
        # values[0] - Район, values[1] - Газ
        region = values[0]
        has_gas = self.utils.get_boolean_from_text(values[1])

        self.log_message("Редактирование выбранного типа комиссии (не реализовано в UI).", level="info")
        messagebox.showinfo("Информация", f"Редактирование типа комиссии '{region}' (Газ: {has_gas}) будет реализовано позже.")
        # Пример использования:
        # self.commission_manager.add_commission_type(region, has_gas, new_composition_dict) # add_commission_type также обновляет
        # self._populate_commission_types_tree()


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
        self.log_message("Добавление нового сопоставления адреса (не реализовано в UI).", level="info")
        messagebox.showinfo("Информация", "Функция добавления сопоставления адреса будет реализована позже с модальным окном.")
        # Пример использования:
        # self.commission_manager.add_address_map("Новый Адрес", "Новый Район", False)
        # self._populate_address_map_tree()

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



Пояснения и следующие шаги:
Создайте файлы: Убедитесь, что вы создали все 5 файлов: config_manager.py, utils.py, commission_manager.py, report_processor.py и обновленный main_app.py, в папке excel_report_filler.
Установите библиотеки: Если вы еще не сделали это, убедитесь, что у вас установлены:
Bash
pip install pandas openpyxl fuzzywuzzy python-Levenshtein


Запуск: Запустите main_app.py.
При первом запуске будет создан config.ini с настройками по умолчанию.
Попробуйте выбрать папки для отчетов, данных и вывода. Пути должны сохраняться между запусками.
Попробуйте загрузить заглушечные файлы комиссий (вам нужно будет создать их вручную, например, Excel-файл commission_types.xlsx с колонками "Район", "Газ", "Председатель", "Должность Председателя", "Член 1", "Должность Члена 1", "Ресурсник", "Должность Ресурсника", и address_map.xlsx с колонками "Адрес", "Район", "Газ").
Запустите сканирование отчетов.
Что реализовано:
Модульная структура: Код разбит на логические модули.
Сохранение/загрузка настроек: ConfigManager обрабатывает config.ini.
Вспомогательные функции: Utils для извлечения адресов, нечеткого сравнения, поиска ячеек и обработки булевых значений.
Управление комиссиями: CommissionManager загружает/сохраняет типы комиссий (по районам и наличию газа) и сопоставления адресов. Реализованы заглушки для CRUD-операций и экспорт.
Обработка отчётов: ReportProcessor теперь:
Сканирует папки.
Ищет файлы данных по адресу.
Определяет наличие газа в самом отчёте по ключевым словам и смещению ячейки (настраивается в config.ini).
Получает состав комиссии на основе района и наличия газа.
Содержит логику вставки строк (_insert_rows) для дополнительных членов комиссии (например, Ресурсника), но её универсальность очень зависит от структуры шаблонов отчётов.
Использует fuzzy_match для поиска полей.
Генерирует отчёт об обработке в Excel.
Многопоточность: Операции заполнения отчётов выполняются в отдельных потоках, чтобы GUI не зависал.
Обновленный UI: Заглушки Treeview для комиссий, поля для файлов комиссий, индикаторы прогресса.
Следующие шаги для дальнейшей доработки:
Реализация полноценного UI для CRUD операций с комиссиями: Сейчас кнопки "Добавить/Редактировать" просто выдают сообщение. Нужно создать модальные окна для ввода и редактирования данных.
Уточнение логики вставки строк: Метод _insert_rows в ReportProcessor очень чувствителен к структуре отчётов. Возможно, потребуется более гибкий механизм (например, поиск конкретного блока текста и вставка строк внутри него).
Ручное сопоставление полей: Когда fuzzy_match не справляется, нужно реализовать UI для того, чтобы пользователь мог вручную указать соответствие поля в данных с ячейкой в отчёте.
Расширенная валидация данных: Улучшить проверку данных при загрузке комиссий.
Пожалуйста, попробуйте запустить это и дайте мне знать, как оно работает!
Источники
1. https://github.com/Dadudekc/MLRobotmaker
2. https://github.com/victors003/python_guide
3. https://www.blackhatworld.com/posts/16355277/
4. https://github.com/OmarElsiry/yt-bulk-downloader
5. https://github.com/respunn/re-render
