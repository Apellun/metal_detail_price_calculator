import pandas as pd
from const import DEFAULT_ACCOUNTS_TABLE

class Saving:
    def __init__(self):
        self.accounting_table_path = DEFAULT_ACCOUNTS_TABLE
        self.detail = None
        self._detail_dict = None
        self.detail_df = None
        self.full_df = None
        
    def _create_detail_df(self) -> None:
        """
        Создает из полей детали таблицу.
        """
        dict = {
        'Название чертежа': self.detail.blueprint_name,
        'Материал': self.detail.metal_category,
        'Толщина детали': self.detail.metal_thickness,
        'Площадь детали': self.detail.metal_area,
        'Цена детали': self.detail.detail_price,
        'Резка, м.п': self.detail.cutting,
        'Врезка, количество': self.detail.in_cutting_amount,
        'Цена врезки': self.detail.full_in_cutting_price,
        'Цена, рез+врезка': self.detail.full_cutting_price,
        'Количетво деталей': self.detail.details_amount,
        'Стоимость деталей': self.detail.complect_price,
        'Количетво комплектов': self.detail.complects_amount,
        'Стоимость комплектов': self.detail.final_price
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
    
    def set_accounting_table_path(self, accounts_path: str=None) -> None:
        """
        Задает путь к таблице для сохранения деталей.
        """
        if accounts_path is None:
            self.accounting_table_path = DEFAULT_ACCOUNTS_TABLE
        elif accounts_path != "":
            self.accounting_table_path = accounts_path
        
    def save_detail(self) -> None:
        """
        Запускает создание таблицы и сохраняет ее в таблицу с деталями.
        """
        self._create_detail_df()
        self._add_detail_to_accounts()
        try:
            self.full_df.to_excel(self.accounting_table_path)
        except FileNotFoundError:
            with open(self.accounting_table_path, 'w', encoding='utf-8'):
                self.full_df.to_excel(self.accounting_table_path)
        
    def create_doc_to_print(self, folder, name) -> None:
        """
        Создает файл с таблицей по текущей детали для печати.
        """
        if self.detail_df is None:
            self._create_detail_df()
        file_path = f"{folder}/{name}.xlsx"
        with open(file_path, 'w', encoding='utf-8'):
            self.detail_df.to_excel(file_path)