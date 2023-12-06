import pandas as pd
from const import DEFAULT_PRICES_TABLE_PATH

class TableReader:
    def __init__(self):
        self.metal_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        self._cutting_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        self.cutting_prices_df = None
        self.metal_prices_df = None
        
    def _set_metal_prices_table(self, metal_prices_table_path: str) -> None:
        """
        Создает таблицу со стоимостью металлов.
        """
        try:
            self.metal_prices_df = pd.read_excel(metal_prices_table_path, sheet_name='стоимость металлов', index_col=0)
            self.metal_prices_table_path = metal_prices_table_path
            self._cutting_prices_table_path = metal_prices_table_path
        except Exception as e:
            raise Exception(f'Не удалось прочитать таблицу со стоимостями металлов.\n{e}')
        
    def _set_cutting_prices_table(self, metal_category: str) -> None:
        """
        Создает таблицу со стоимостью резки и врезки.
        """
        try:
            self.cutting_prices_df = pd.read_excel(self._cutting_prices_table_path, sheet_name=f'резка и врезка {metal_category}', index_col=0)
        except Exception as e:
                raise Exception(f'Не удалось прочитать таблицу со стоимостями резки и врезки.\n{e}')
            
    def set_metal_prices_table(self, metal_prices_table_path: str=None) -> None:
        """
        Устанавливает путь к таблице со стоимостями металлов. Если
        аргумент с путем не передан, устанавливает путь по умолчанию.
        """
        if metal_prices_table_path is None:
            metal_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        if metal_prices_table_path != "":
            self._set_metal_prices_table(metal_prices_table_path)
        
    def get_metals_list(self) -> list:
        """
        Получает список доступных типов металла из таблицы.
        """
        if self.metal_prices_df is None:
            self.set_metal_prices_table()
        return self.metal_prices_df.index.values.tolist()
        
    def get_metal_price(self, metal_type: str) -> float:
        """
        Возвращает из таблицы стоимость металла.
        """
        try:
            return float(self.metal_prices_df.loc[metal_type].iloc[0])
        except Exception as e:
            raise Exception(f'Цена металла не обнаружена.\n{e}')
        
    def get_cutting_price(self, metal_category: str, metal_thickness: float, details_amount: int) -> float:
        """
        Возвращает из таблицы стоимость резки.
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
        Возвращает из таблицы стоимость врезки.
        """
        if self.cutting_prices_df is None:
            self._set_cutting_prices_table(metal_category)
        try:
            return float(self.cutting_prices_df.loc[metal_thickness].iloc[2])
        except Exception as e:
            raise Exception(f'Цена врезки для введенных толщины и категории металла не обнаружена.\n{e}')
    
