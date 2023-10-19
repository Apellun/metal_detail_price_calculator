import pandas as pd
from detail import Detail
from const import BASE_DIR_STR

class TableCreator:
    def __init__(self, detail: Detail):
        self.detail = detail
        self.detail_df = None
        self.full_df = None
        self._accounts_table_path = None
        self._dict = None
        self._previous_df = None
        
    def _create_dict(self) -> None:
        """
        Создает словарь с названием столбца и значением для детали.
        """
        self._dict = {
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
    
    def _create_detail_df(self) -> None:
        """
        Создает из словаря таблицу.
        """
        self.detail_df = pd.DataFrame(self._dict, index=[0]).reset_index(drop=True)
        
    def _read_previous(self) -> None:
        """
        Читает таблицу с информацией о предыдущих деталях.
        """
        self._accounts_table_path = BASE_DIR_STR + '/accounts.xlsx'
        self.previous_df = pd.read_excel(self._accounts_table_path, index_col=0).reset_index(drop=True)
        
    def _create_full_df(self) -> None:
        """
        Соединяет старые значения из таблицы с новым.
        """
        self._create_dict()
        self._create_detail_df()
        self._read_previous()
        self.full_df = pd.concat([self.previous_df, self.detail_df])
        
    def save_detail(self) -> None:
        """
        Запускает создание таблицы и сохраняет ее.
        """
        self._create_full_df()
        self.full_df.to_excel(self._accounts_table_path)
        
    def create_doc_to_print(self) -> None:
        """
        Создает отдельную таблицу с текущей деталью для печати.
        """
        table_dir = BASE_DIR_STR + '/doc_to_print.xlsx'
        self.detail_df.to_excel(table_dir)