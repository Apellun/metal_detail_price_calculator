import os
from calculation import Calculation
from const import metal_density_dict
        
class Manager:
    def __init__(self, table_reader, saving):
        self.table_reader = table_reader
        self.saving = saving
        self.calculation = None
    
    def _create_dict(self, values: dict) -> dict:
        """
        Creates a dictionary for the calculation.
        """
        return {
            'detail_name': values['Detail name'],
            'metal_category': values['Metal category'],
            'metal_type': values['Metal type'],
            'metal_thickness': values["Metal thickness"],
            'metal_area': values["Metal area"],
            'metal_price': self.table_reader.get_metal_price(values['Metal type']),
            'cutting': values["Cutting"],
            'inset_amount': values["Inset amount"],
            'cutting_price': self.table_reader.get_cutting_price(values['Metal category'], values["Metal thickness"], values["Details amount"]),
            'inset_price': self.table_reader.get_inset_price(values['Metal category'], values["Metal thickness"]),
            'details_amount': values["Details amount"],
            'complects_amount': values["Complects amount"],
            'density': metal_density_dict[values['Metal category']]
        }
    
    def _validate_fields_not_empty(self, values: dict):
        """
        Checks if the fields are not empty.
        """
        for index in values:
            value = values[index]
            if value == "":
                raise ValueError(f"{index}: field can't be empty.")
        
    def _validate_calculation_fields(self, values: dict) -> dict:
        """
        Check if the fields are not empty, specific fields have
        correct numeric values, converts field values to
        correct data types.
        """
        for index in values:
            value = values[index]
            
            if value == "":
                raise ValueError(f"{index}: field can't be empty.")
            
            try:
                if index in ("Metal area", "Cutting"):
                        values[index] = float(value.replace(',', '.'))
                elif index in ("Inset amount", "Details amount", "Complects amount"):
                        values[index] = int(value)
            except Exception as e:
                raise ValueError(f"{index}: invalid value.\n{e}")
        
        return values
    
    def _validate_table_path(self, table_path: str) -> None:
        """
        Checks that the chosen path to the spreadsheet
        is the path to the Excel file.
        """
        if table_path is not None and table_path != "":
            if not table_path.endswith((".xls", ".xlsx")):
                raise Exception("Please make sure you chose excel files ('.xls', '.xlsx').")
        
    def _check_path_if_exists(self, file_path: str) -> None:
        """
        Checks if the file exists.
        """
        if os.path.exists(file_path):
            raise FileExistsError
    
    def get_metals_list(self) -> list:
        """
        Returns a list of metals from the metals
        prices table.
        """
        return self.table_reader.get_metals_list()
    
    def get_file_paths(self) -> tuple:
        """
        Returns paths for the prices spreadsheet and
        the accounting spreadsheet.
        """
        return self.table_reader.metal_prices_table_path, self.saving.accounting_table_path
    
    def save_settings(self, metals_table_path: str = None, calculationing_table_path: str = None, save_for_all: bool = False) -> None:
        """
        Sets the paths to the prices spreadsheet and
        the accounting spreadsheet.
        """
        self._validate_table_path(metals_table_path)
        self._validate_table_path(calculationing_table_path)
        self.table_reader.set_metal_prices_table(metals_table_path, save_for_all)
        self.saving.set_accounting_table_path(calculationing_table_path)
        
    def create_calculation(self, values: dict) -> str:
        """
        Runs the field value check and then creates a calculation.
        """
        values_validated = self._validate_calculation_fields(values)
        data = self._create_dict(values_validated)
        self.calculation = Calculation(**data)
    
    def create_prices_message(self) -> str:
        """
        Creates a string with prices and detail info.
        """
        return (f"The price of the metal: {self.calculation.detail_price}\n"
                f"The price of cutting and inset: {self.calculation.full_cutting_price}\n"
                f"One complect price: {self.calculation.complect_price}\n"
                f"Full price: {self.calculation.final_price}\n"
                f"Detail's weight: {self.calculation.mass} kg")
    
    def save_calculations(self) -> None:
        """
        Runs the saving of the calculation to the accounts
        table.
        """
        self.saving.save_calculations(self.calculation)
    
    def save_doc_to_print(self, values: dict) -> None:
        """
        Runs the field value check, if they are not empty, creates
        a path for the file and checks is the file by this path
        exists. If not — runs the creation of the spreadsheet
        with calculation details by the set path. 
        """
        self._validate_fields_not_empty(values)
        saving_path = f"{values['Folder']}/{values['Name']}.xlsx"
        self._check_path_if_exists(saving_path)
        self.saving.create_doc_to_print(self.calculation, saving_path)
        
    def overvrite_file(self, values: dict) -> None:
        """
        Runs the creation of the spreadsheet with calculation
        details with overwriting an existing file.
        """
        saving_path = f"{values['Folder']}/{values['Name']}.xlsx"
        self.saving.create_doc_to_print(self.calculation, saving_path)