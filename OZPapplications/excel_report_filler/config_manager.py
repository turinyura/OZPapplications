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
        self.load_config()

    def load_config(self):
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