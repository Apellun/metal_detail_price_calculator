import pandas as pd
from const import DEFAULT_ACCOUNTS_TABLE

class Saving:
    def __init__(self):
        self.accounting_table_path = DEFAULT_ACCOUNTS_TABLE
        self._calculation_dict = None
        self.calculation_df = None
        self._accounts_df = None
        
    def _create_calculation_df(self, calculation) -> None:
        """
        Creates a dataframe from the calculation data.
        """
        dict = {
        'Detail name': calculation.detail_name,
        'Material': calculation.metal_category,
        'Detail thickness': calculation.metal_thickness,
        'Detail area': calculation.metal_area,
        'Detail mass': calculation.mass,
        'Detail price': calculation.detail_price,
        'Cutting, running m': calculation.cutting,
        'Inset amount': calculation.inset_amount,
        'Inset price': calculation.full_inset_price,
        'Price of cutting + inset': calculation.full_cutting_price,
        'Details amount': calculation.details_amount,
        'Details cost': calculation.complect_price,
        'Complects amount': calculation.complects_amount,
        'Complects cost': calculation.final_price
        }
        self.calculation_df = pd.DataFrame(dict, index=[0])
        
    def _add_calculation_to_accounts(self) -> None:
        """
        Unites the previous accounts with a new one into a
        single dataframe.
        """
        try:
            accounts_df = pd.read_excel(self.accounting_table_path, index_col=0)
            self._accounts_df = pd.concat([accounts_df, self.calculation_df]).reset_index(drop=True)
            self._accounts_df.index = self._accounts_df.index + 1
            
        except FileNotFoundError:
            self._accounts_df = self.calculation_df
            
        except Exception as e:
            raise Exception(f"Saving error.\n{e}")
    
    def _set_accounting_table_path(self, accounting_table_path: str) -> None:
        """
        Sets a path to the accounting spreadsheet.
        """
        self.accounting_table_path = accounting_table_path
    
    def set_accounting_table_path(self, accounting_table_path: str=None) -> None:
        """
        Sets a path to the accounting spreadsheet. If the path
        argument is empty, sets it to the default path.
        """
        if accounting_table_path is None:
            accounting_table_path = DEFAULT_ACCOUNTS_TABLE
        if accounting_table_path != "":
            self._set_accounting_table_path(accounting_table_path)
        
    def save_calculations(self, calculation: object) -> None:
        """
        Runs the creation of the calculation dataframe and its
        join with the previous calculations.
        """
        self._create_calculation_df(calculation)
        self._add_calculation_to_accounts()
        
        try:
            self._accounts_df.to_excel(self.accounting_table_path)
            
        except FileNotFoundError:
            with open(self.accounting_table_path, 'w', encoding='utf-8'):
                self._accounts_df.to_excel(self.accounting_table_path)
                
        except Exception as e:
            raise Exception(f"Saving error.\n{e}")        
        
    def create_doc_to_print(self, calculation: object, file_path: str) -> None:
        """
        Creates a spreadsheet with a current calculation
        for print.
        """
        self._create_calculation_df(calculation)
        
        try:
            with open(file_path, 'w', encoding='utf-8'):
                self.calculation_df.to_excel(file_path)
        except Exception as e:
            raise Exception(f"Saving error.\n{e}")