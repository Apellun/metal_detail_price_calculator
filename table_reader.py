import pandas as pd

class TableReader:
    def __init__(self, metal_cost_file_path: str, cutting_cost_file_path=None):
        self._metal_cost_file_path = metal_cost_file_path
        self._cutting_cost_file_path = cutting_cost_file_path
        self.cutting_cost_df = None
        self.metal_cost_df = None
        
    def _read_metal_costs_table(self) -> None:
        """
        Инициализирует таблицу со стоимостью металлов.
        """
        self.metal_cost_df = pd.read_excel(self._metal_cost_file_path, sheet_name='стоимость металлов', index_col=0)
    
    def _read_cutting_costs_table(self, metal_category: str) -> None:
        """
        Инициализирует таблицу со стоимостью резки и врезки.
        """
        if self._cutting_cost_file_path == None:
            self._cutting_cost_file_path = self._metal_cost_file_path
        self.cutting_cost_df = pd.read_excel(self._cutting_cost_file_path, sheet_name=f'резка и врезка {metal_category}', index_col=0)
        
    def get_metal_cost(self, metal_type: str) -> None:
        """
        Возвращает из таблицы стоимость металла.
        """
        self._read_metal_costs_table()
        return self.metal_cost_df.loc[metal_type].iloc[0]
        
    def get_cutting_cost(self, metal_category: str, metal_thickness: str, details_amount: int) -> float:
        """
        Возвращает из таблицы стоимость резки.
        """
        self._read_cutting_costs_table(metal_category)
        if details_amount <= 100:
            return self.cutting_cost_df.loc[metal_thickness].iloc[0]
        else:
            return self.cutting_cost_df.loc[metal_thickness].iloc[1]
        
    def get_in_cutting_cost(self, metal_thickness: str) -> float:
        """
        Возвращает из таблицы стоимость врезки.
        """
        return self.cutting_cost_df.loc[metal_thickness].iloc[2]
    
