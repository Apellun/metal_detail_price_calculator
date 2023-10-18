import pandas as pd
from detail import Detail

class TableCreator:
    def __init__(self, detail: Detail):
        self.detail = detail
        self._dict = NotImplemented
        self.detail_df = NotImplemented
        self.previous_df = NotImplemented
        self.full_df = NotImplemented
        
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
        self.previous_df = pd.read_excel('accounts.xlsx', index_col=0).reset_index(drop=True)
        
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
        self.full_df.to_excel('accounts.xlsx')
        
    def create_doc_to_print(self) -> None:
        """
        Создает отдельную таблицу с текущей деталью для печати.
        """
        self.detail_df.to_excel('doc_to_print.xlsx')