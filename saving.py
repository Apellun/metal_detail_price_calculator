import pandas as pd
from const import DEFAULT_ACCOUNTS_TABLE

class Saving:
    def __init__(self):
        self.accounting_table_path = DEFAULT_ACCOUNTS_TABLE
        self._detail_dict = None
        self.detail_df = None
        self.full_df = None
        
    def _create_detail_df(self, detail) -> None:
        """
        Создает из полей детали таблицу.
        """
        dict = {
        'Название чертежа': detail.blueprint_name,
        'Материал': detail.metal_category,
        'Толщина детали': detail.metal_thickness,
        'Площадь детали': detail.metal_area,
        'Цена детали': detail.detail_price,
        'Резка, м.п': detail.cutting,
        'Врезка, количество': detail.in_cutting_amount,
        'Цена врезки': detail.full_in_cutting_price,
        'Цена, рез+врезка': detail.full_cutting_price,
        'Количетво деталей': detail.details_amount,
        'Стоимость деталей': detail.complect_price,
        'Количетво комплектов': detail.complects_amount,
        'Стоимость комплектов': detail.final_price
        }
        self.detail_df = pd.DataFrame(dict, index=[0]).reset_index(drop=True)
        
    def _add_detail_to_accounts(self) -> None:
        """
        Создает таблицу с предыдущими деталями и новой деталью.
        """
        try:
            accounts_df = pd.read_excel(self.accounting_table_path, index_col=0).reset_index(drop=True)
            self.full_df = pd.concat([accounts_df, self.detail_df]).reset_index(drop=True)
            self.full_df.index = self.full_df.index + 1
        except FileNotFoundError:
            self.full_df = self.detail_df
    
    def set_accounting_table_path(self, accounting_table_path=None) -> None:
        """
        Задает путь к таблице для сохранения деталей.
        """
        if accounting_table_path is None:
            accounting_table_path = DEFAULT_ACCOUNTS_TABLE
        if accounting_table_path != "":
            self.accounting_table_path = accounting_table_path
        
    def save_accounts(self, detail: object) -> None:
        """
        Запускает создание таблицы и сохраняет ее в таблицу с деталями.
        """
        self._create_detail_df(detail)
        self._add_detail_to_accounts()
        try:
            self.full_df.to_excel(self.accounting_table_path)
        except FileNotFoundError:
            with open(self.accounting_table_path, 'w', encoding='utf-8'):
                self.full_df.to_excel(self.accounting_table_path)
        
    def create_doc_to_print(self, detail: object, file_path: str) -> None:
        """
        Создает файл с таблицей по текущей детали для печати.
        """
        if self.detail_df is None:
            self._create_detail_df(detail)
        with open(file_path, 'w', encoding='utf-8'):
            self.detail_df.to_excel(file_path)