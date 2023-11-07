import os
from calculation import Calculation
from table_reader import TableReader
from saving import Saving
from const import metal_density_dict
        
class Manager:
    def __init__(self, table_reader: TableReader, saving: Saving):
        self.table_reader = table_reader
        self.saving = saving
        self.calculation = None
    
    def _create_dict(self, values: dict) -> dict:
        """
        Собирает словарь для создания расчета.
        """
        return {
            'blueprint_name': values['Название чертежа'],
            'metal_category': values['Категория металла'],
            'metal_type': values['Тип металла'],
            'metal_thickness': values["Толщина металла"],
            'metal_area': values["Площадь металла"],
            'metal_price': self.table_reader.get_metal_price(values['Тип металла']),
            'cutting': values["Резка, м. п."],
            'inset_amount': values["Врезка, количество"],
            'cutting_price': self.table_reader.get_cutting_price(values['Категория металла'], values["Толщина металла"], values["Количество деталей"]),
            'inset_price': self.table_reader.get_inset_price(values['Категория металла'], values["Толщина металла"]),
            'details_amount': values["Количество деталей"],
            'complects_amount': values["Количество комплектов"],
            'density': metal_density_dict[values['Категория металла']]
        }
    
    def _validate_fields_not_empty(self, values: dict):
        """
        Проверяет, что переданные поля не пустые.
        """
        for index in values:
            value = values[index]
            if value == "":
                raise ValueError(f"{index}: поле не может быть пустым.")
        
    def _validate_calculation_fields(self, values: dict) -> dict:
        """
        Проверяет, что переданные поля не пустые, в нужных полях
        указаны числовые значения правильного числового типа, перезаполняет
        словарь.
        """
        for index in values:
            value = values[index]
            if value == "":
                raise ValueError(f"{index}: поле не может быть пустым.")
            try:
                if index in ("Площадь металла", "Резка, м. п."):
                        values[index] = float(value.replace(',', '.'))
                elif index in ("Врезка, количество", "Количество деталей", "Количество комплектов"):
                        values[index] = int(value)
            except Exception as e:
                raise ValueError(f"{index}: недопустимое значение.\n{e}")
        return values
    
    def _validate_table_path(self, table_path: str) -> None:
        """
        Проверяет, что в качестве адреса таблицы передан
        файл Excel.
        """
        if table_path is not None and table_path != "":
            if not table_path.endswith((".xls", ".xlsx")):
                raise Exception('Пожалуйста, убедитесь что выбраны таблицы формата excel (".xls", ".xlsx").')
        
    def _check_path_if_exists(self, file_path: str) -> None:
        """
        Проверяет, существует ли файл по заданному пути.
        """
        if os.path.exists(file_path):
            raise FileExistsError
    
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
    
    def save_settings(self, metals_table_path: str=None, calculationing_table_path: str=None) -> None:
        """
        Сохраняет заданные пути таблиц с металлами и расчетами.
        """
        self._validate_table_path(metals_table_path)
        self._validate_table_path(calculationing_table_path)
        self.table_reader.set_metal_prices_table(metals_table_path)
        self.saving.set_accounting_table_path(calculationing_table_path)
        
    def create_calculation(self, values: dict) -> str:
        """
        Запускает проверку значений полей и затем создание расчета.
        """
        values_validated = self._validate_calculation_fields(values)
        data = self._create_dict(values_validated)
        self.calculation = Calculation(**data)
    
    def create_prices_message(self) -> str:
        """
        Создает строку для вывода стоимости работ.
        """
        return (f"Стоимость металла для детали: {self.calculation.detail_price}\n"
                f"Цена резки и врезки: {self.calculation.full_cutting_price}\n"
                f"Цена одного комплекта деталей: {self.calculation.complect_price}\n"
                f"Полная цена: {self.calculation.final_price}\n"
                f"Масса детали: {self.calculation.mass}")
    
    def save_calculations(self) -> None:
        """
        Запускает сохранение расчета в таблицу с расчетами.
        """
        self.saving.save_calculations(self.calculation)
    
    def save_doc_to_print(self, values: dict) -> None:
        """
        Запускает проверку полей, если поля не пустые, создает
        путь для файла, запускает проверку, существует ли файл,
        если нет — запускает сохрание расчета в файл для печати.
        """
        self._validate_fields_not_empty(values)
        saving_path = f"{values['Папка']}/{values['Имя']}.xlsx"
        self._check_path_if_exists(saving_path)
        self.saving.create_doc_to_print(self.calculation, saving_path)
        
    def overvrite_file(self, values: dict) -> None:
        """
        Запускает сохранение файла вместо существующего.
        """
        saving_path = f"{values['Папка']}/{values['Имя']}.xlsx"
        self.saving.create_doc_to_print(self.calculation, saving_path)