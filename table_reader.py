import pandas as pd
from const import DEFAULT_PRICES_TABLE_PATH

class TableReader:
    def __init__(self):
        self.cutting_prices_df = None
        self.metal_prices_df = None
        self._cutting_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        self._metal_prices_table_path = DEFAULT_PRICES_TABLE_PATH
    
    def _set_metal_prices_table(self, metal_prices_table_path=None):
        if metal_prices_table_path is None:
            metal_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        #add extension check
        if metal_prices_table_path != "":
            try:
                self.metal_prices_df = pd.read_excel(metal_prices_table_path, sheet_name='стоимость металлов', index_col=0)
            except:
                raise Exception('Не удалось прочитать таблицу.')
    
    def _set_cutting_prices_table_path(self, cutting_prices_table_path=None):
        if cutting_prices_table_path is None:
            cutting_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        #add extension check
        if cutting_prices_table_path != "":
            self._cutting_prices_table_path = cutting_prices_table_path
        
    def _set_cutting_prices_table(self, metal_category):
        try:
            self.cutting_prices_df = pd.read_excel(self._cutting_prices_table_path, sheet_name=f'резка и врезка {metal_category}', index_col=0)
        except:
                raise Exception('Таблица не найдена.')
            
    def set_tables(self, metal_prices_table_path=None, cutting_prices_table_path=None):
        self._set_metal_prices_table(metal_prices_table_path)
        # self._set_cutting_prices_table_path(cutting_prices_table_path)
        
    def get_metals_list(self)-> None:
        if self.metal_prices_df is None:
            self._set_metal_prices_table()
        return self.metal_prices_df.index.values.tolist()
        
    def get_metal_price(self, metal_type: str) -> None:
        """
        Возвращает из таблицы стоимость металла.
        """
        try:
            return self.metal_prices_df.loc[metal_type].iloc[0]
        except:
            raise Exception('Цена металла не обнаружена.')
        
    def get_cutting_price(self, metal_category: str, metal_thickness: float, details_amount: int) -> float: #TODO
        """
        Возвращает из таблицы стоимость резки.
        """
        self._set_cutting_prices_table(metal_category)
        
        try:
            if details_amount <= 100:
                return self.cutting_prices_df.loc[metal_thickness].iloc[0]
            else:
                return self.cutting_prices_df.loc[metal_thickness].iloc[1]
        except Exception as e:
            raise Exception(f'Цена резки не обнаружена: {e}')
        
    def get_in_cutting_price(self, metal_thickness: float) -> float:
        """
        Возвращает из таблицы стоимость врезки.
        """
        try:
            return self.cutting_prices_df.loc[metal_thickness].iloc[2]
        except:
            raise Exception('Цена врезки не обнаружена.')
    
