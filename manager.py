import os
from detail import Detail
from table_reader import TableReader
from saving import Saving
        
class Manager:
    def __init__(self, table_reader: TableReader, saving: Saving):
        self.table_reader = table_reader
        self.saving = saving
        self.detail = None
    
    def _create_dict(self, values: dict) -> dict:
        """
        Собирает словарь из собранных в интерфейсе данных
        для создания детали.
        """
        return {
            'blueprint_name': values['Название чертежа'],
            'metal_category': values['Категория металла'],
            'metal_type': values['Тип металла'],
            'metal_thickness': float(values["Толщина металла"]),
            'metal_area': float(values["Площадь металла"]),
            'cutting': float(values["Резка, м. п."]),
            'in_cutting_amount': int(values["Врезка, количество"]),
            'details_amount': int(values["Количество деталей"]),
            'complects_amount': int(values["Количество комплектов"])
        }

    def _create_prices_str(self, detail: Detail) -> str:
        """
        Создает строку для вывода стоимости работ.
        """
        return (f"Стоимость металла для детали: {detail.detail_price}\n"
                f"Цена резки и врезки: {detail.full_cutting_price}\n"
                f"Цена одного комплекта деталей: {detail.complect_price}\n"
                f"Полная цена: {detail.final_price}")
        
    def _validate_fields(self, values: dict) -> None:
        """
        Проверяет значения в полях на отсутствие пустых строк
        и наличие числовых значений там, где они требуются.
        """
        for index in values:
            value = values[index].strip()
            if value == "":
                raise ValueError(f"{index}: поле не может быть пустым.")
            if index not in ["Название чертежа", "Категория металла", "Тип металла"] and not value.isdigit():
                raise ValueError(f"{index}: введите числовое значение.")
    
    def _validate_table_path(self, table_path: str) -> None:
        """
        Проверяет, что в качестве адреса таблицы передан
        файл Excel.
        """
        if table_path is not None and table_path != "":
            if not table_path.endswith((".xls", ".xlsx")):
                raise Exception('Пожалуйста, убедитесь что выбраны таблицы формата excel (".xls", ".xlsx")')
    
    def _check_path_if_exists(self, file_path: str) -> None:
        """
        Проверяет, существует ли файл по заданному пути.
        """
        if os.path.exists(file_path):
            raise FileExistsError
            
    def _create_detail(self, values: dict) -> Detail:
        """
        Создает деталь, заполняет поля детали значениями
        из таблиц и расчетами, возвращает созданную деталь.
        """
        data = self._create_dict(values)
        self.detail = Detail(**data)
        self.detail.metal_price = self.table_reader.get_metal_price(self.detail.metal_type)
        self.detail.cutting_price = self.table_reader.get_cutting_price(self.detail.metal_category, self.detail.metal_thickness, self.detail.details_amount)
        self.detail.in_cutting_price = self.table_reader.get_in_cutting_price(self.detail.metal_thickness)
        self.detail.set_prices()
        return self.detail
    
    def get_metals_list(self) -> list:
        """
        Возвращает список металлов из таблицы с ценами
        металлов.
        """
        return self.table_reader.get_metals_list()
    
    def get_file_paths(self) -> tuple:
        """
        Возвращает пути к таблице с металлами и таблице с расчетами.
        """
        return self.table_reader.metal_prices_table_path, self.saving.accounting_table_path
    
    def save_settings(self, metals_table_path=None, accounting_table_path=None) -> None:
        """
        Сохраняет настройки таблиц с металлами и расчетами.
        """
        self._validate_table_path(metals_table_path)
        self._validate_table_path(accounting_table_path)
        self.table_reader.set_metal_prices_table(metals_table_path)
        self.saving.set_accounting_table_path(accounting_table_path)
        
    def count_prices(self, values: dict) -> str:
        """
        Запускает проверку значений полей и затем создание детали,
        возвращает строку с расчетом цен.
        """
        self._validate_fields(values)
        self._create_detail(values)
        return self._create_prices_str(self.detail)
    
    def save_accounts(self) -> None:
        """
        Запускает сохранение расчета в таблицу с расчетами.
        """
        self.saving.save_accounts(self.detail)
    
    def save_doc_to_print(self, values: dict) -> None: #TODO: separate validation?
        file_name, file_folder = values['Имя'], values['Папка']
        if file_name == '':
            raise Exception("Укажите имя файла.")
        if file_folder == '':
            raise Exception("Выберите папку для сохранения.")
        saving_path = f"{values['Папка']}/{values['Имя']}.xlsx"
        self._check_path_if_exists(saving_path)
        self.saving.create_doc_to_print(self.detail, saving_path)
        
    def overvrite_file(self, values: dict) -> None:
        """
        Запускает сохранение файла вместо существующего.
        """
        saving_path = f"{values['Папка']}/{values['Имя']}.xlsx"
        self.saving.create_doc_to_print(self.detail, saving_path)