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