import pandas as pd
from const import DEFAULT_PRICES_TABLE_PATH

class TableReader:
    def __init__(self):
        self.metal_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        self.cutting_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        self.cutting_prices_df = None
        self.metal_prices_df = None
        
    def _set_metal_prices_table(self, metal_prices_table_path: str, save_for_all: bool = None) -> None:
        """
        Creates a dataframe with metal prices
        """
        try:
            self.metal_prices_table_path = metal_prices_table_path
            self.metal_prices_df = pd.read_excel(metal_prices_table_path, sheet_name='metal_price', index_col=0)
            if save_for_all:
                self.cutting_prices_table_path = metal_prices_table_path
        except Exception as e:
            raise Exception(f'Error reading the metal prices spreadsheet.\n{e}')
        
    def _set_cutting_prices_table(self, metal_category: str) -> None:
        """
        Creates a dataframe with cutting and inset prices.
        """
        try:
            self.cutting_prices_df = pd.read_excel(self.cutting_prices_table_path, sheet_name=f'cutting_inset {metal_category}', index_col=0)
        except Exception as e:
                raise Exception(f'Error reading the cutting and inset prices spreadsheet.\n{e}')
            
    def set_metal_prices_table(self, metal_prices_table_path: str = None, save_for_all: bool = None) -> None:
        """
        Sets a path to the spreadsheet with the metal prices. If
        the path argument is empty, sets a default path.
        """
        if metal_prices_table_path is None:
            metal_prices_table_path = DEFAULT_PRICES_TABLE_PATH
        if metal_prices_table_path != "":
            self._set_metal_prices_table(metal_prices_table_path, save_for_all)
        
    def get_metals_list(self) -> list:
        """
        Gets a list of the possible metal types from the
        matal prices dataframe.
        """
        if self.metal_prices_df is None:
            self.set_metal_prices_table()
        return self.metal_prices_df.index.values.tolist()
        
    def get_metal_price(self, metal_type: str) -> float:
        """
        Returns a price of the metal from the dataframe.
        """
        try:
            return float(self.metal_prices_df.loc[metal_type].iloc[0])
        except Exception as e:
            raise Exception(f'Metal price not found.\n{e}')
        
    def get_cutting_price(self, metal_category: str, metal_thickness: float, details_amount: int) -> float:
        """
        Returns a price of the cutting from the dataframe.
        """
        if self.cutting_prices_df is None:
            self._set_cutting_prices_table(metal_category)
            
        try:
            if details_amount <= 100:
                return float(self.cutting_prices_df.loc[metal_thickness].iloc[0])
            else:
                return float(self.cutting_prices_df.loc[metal_thickness].iloc[1])
        except Exception as e:
            raise Exception(f'Cutting price for set thickness and category of metal not found.\n{e}')
        
    def get_inset_price(self, metal_category: str, metal_thickness: float) -> float:
        """
        Returns a price of the inset from the dataframe.
        """
        if self.cutting_prices_df is None:
            self._set_cutting_prices_table(metal_category)
            
        try:
            return float(self.cutting_prices_df.loc[metal_thickness].iloc[2])
        except Exception as e:
            raise Exception(f'Inset price for set thickness and category of metal not found.\n{e}')
    
