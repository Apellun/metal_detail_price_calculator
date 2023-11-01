import os
from table_reader import TableReader
from saving import Saving
from detail import Detail

table_reader = TableReader()
saving = Saving()
        
class Manager:
    def __init__(self):
        self.detail = None
    
    def _create_dict(self, values: dict) -> dict:
        """
        Собирает словарь с введенными данными для передачи в деталь.
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

    def _create_prices_str(self, detail: object) -> str:
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
    
    def _validate_table(self, table_path):
        if table_path is not None and table_path != "":
            if not table_path.endswith((".xls", ".xlsx")):
                raise Exception('Пожалуйста, убедитесь что выбраны таблицы формата excel (".xls", ".xlsx")')
    
    def _validate_path_for_saving(self, file_path):
        if os.path.exists(file_path):
            raise FileExistsError
            
    def _create_detail(self, values: dict) -> None:
        """
        Создает деталь, заполняет поля значениями
        из таблиц и счиатет стоимость работы.
        """
        data = self._create_dict(values)
        self.detail = Detail(**data)
        self.detail.metal_price = table_reader.get_metal_price(self.detail.metal_type)
        self.detail.cutting_price = table_reader.get_cutting_price(self.detail.metal_category, self.detail.metal_thickness, self.detail.details_amount)
        self.detail.in_cutting_price = table_reader.get_in_cutting_price(self.detail.metal_thickness)
        self.detail.set_prices()
        return self.detail
    
    def get_metals_list(self) -> list:
        return table_reader.get_metals_list()
    
    def get_file_paths(self) -> tuple:
        return table_reader.metal_prices_table_path, saving.accounting_table_path
    
    def save_settings(self, metals_table_path=None, accounting_table_path=None) -> None:
        self._validate_table(metals_table_path)
        self._validate_table(accounting_table_path)
        table_reader.set_metal_prices_table(metals_table_path)
        saving.set_accounting_table_path(accounting_table_path)
        
    def count_prices(self, values: dict) -> None:
        self._validate_fields(values)
        self._create_detail(values)
        return self._create_prices_str(self.detail)
    
    def save_accounts(self) -> None:
        saving.save_accounts(self.detail)
    
    def save_doc_to_print(self, values: dict) -> None:
        file_name, file_folder = values['Имя'], values['Папка']
        if file_name == '':
            raise Exception("Укажите имя файла.")
        if file_folder == '':
            raise Exception("Выберите папку для сохранения.")
        saving_path = f"{values['Папка']}/{values['Имя']}.xlsx"
        self._validate_path_for_saving(saving_path)
        saving.create_doc_to_print(self.detail, saving_path)
        
    def overvrite_file(self, values: dict) -> None:
        saving_path = f"{values['Папка']}/{values['Имя']}.xlsx"
        saving.create_doc_to_print(self.detail, saving_path)