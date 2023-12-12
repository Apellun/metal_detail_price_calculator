import pandas as pd
from const import DEFAULT_PRICES_TABLE_PATH

class TableReader:
    def __init__(self):
        self.metal_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        self.cutting_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        self.cutting_prices_df = None
        self.metal_prices_df = None
        
        
    def _set_metal_prices_table_path(self, metal_prices_table_path: str) -> None:
        self.metal_prices_table_path = metal_prices_table_path
            
    def _set_cutting_prices_table_path(self, cutting_prices_table_path: str) -> None:
        self.cutting_prices_table_path = cutting_prices_table_path

    def _set_metal_prices_table(self, metal_prices_table_path: str) -> None:
        """
        Создает датафрейм со стоимостями металлов.
        """
        try:
            self.metal_prices_df = pd.read_excel(metal_prices_table_path, sheet_name='metal_price', index_col=0)
        except Exception as e:
            raise Exception(f'Не удалось прочитать таблицу со стоимостями металлов.\n{e}')
        
    def _set_cutting_prices_table(self, metal_category: str) -> None:
        """
        Создает датафрейм со стоимостями резки и врезки.
        """
        try:
            self.cutting_prices_df = pd.read_excel(self.cutting_prices_table_path, sheet_name=f'cutting_inset {metal_category}', index_col=0)
        except Exception as e:
                raise Exception(f'Не удалось прочитать таблицу со стоимостями резки и врезки.\n{e}')
            
    def set_metal_prices_table(self, metal_prices_table_path: str = None, save_for_all: bool = False) -> None:
        """
        Запускает создание датафрейма со стоимостями металлов, изменение путей
        к файлам с ценами. Если аргумент с путем не передан, передает путь
        по умолчанию.
        """
        if metal_prices_table_path is None:
            metal_prices_table_path = DEFAULT_PRICES_TABLE_PATH
            
        if metal_prices_table_path != "":
            try:
                self._set_metal_prices_table(metal_prices_table_path)
                self._set_metal_prices_table_path(metal_prices_table_path)
                if save_for_all:
                    self._set_cutting_prices_table_path(metal_prices_table_path)
            except Exception as e:
                raise Exception(e)
        
    def get_metals_list(self) -> list:
        """
        Получает список доступных типов металлов из датафрейма
        со стоимостями металлов.
        """
        if self.metal_prices_df is None:
            self.set_metal_prices_table()
        return self.metal_prices_df.index.values.tolist()
        
    def get_metal_price(self, metal_type: str) -> float:
        """
        Возвращает из датафрейма стоимость металла.
        """
        try:
            return float(self.metal_prices_df.loc[metal_type].iloc[0])
        except Exception as e:
            raise Exception(f'Цена металла не обнаружена.\n{e}')
        
    def get_cutting_price(self, metal_category: str, metal_thickness: float, details_amount: int) -> float:
        """
        Возвращает из датафрейма стоимость резки для
        указанных категории и толщины металла, учитывая
        количество деталей.
        """
        if self.cutting_prices_df is None:
            self._set_cutting_prices_table(metal_category)
            
        try:
            if details_amount <= 100:
                return float(self.cutting_prices_df.loc[metal_thickness].iloc[0])
            else:
                return float(self.cutting_prices_df.loc[metal_thickness].iloc[1])
        except Exception as e:
            raise Exception(f'Цена резки для введенных толщины и категории металла не обнаружена.\n{e}')
        
    def get_inset_price(self, metal_category: str, metal_thickness: float) -> float:
        """
        Возвращает из таблицы стоимость врезки для
        указанных категории и толщины металла.
        """
        if self.cutting_prices_df is None:
            self._set_cutting_prices_table(metal_category)
            
        try:
            return float(self.cutting_prices_df.loc[metal_thickness].iloc[2])
        except Exception as e:
            raise Exception(f'Цена врезки для введенных толщины и категории металла не обнаружена.\n{e}')
    
