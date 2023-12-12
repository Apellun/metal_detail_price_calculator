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
        Создает из полей расчета датафрейм.
        """
        dict = {
        'Название чертежа': calculation.blueprint_name,
        'Материал': calculation.metal_category,
        'Толщина детали': calculation.metal_thickness,
        'Площадь детали': calculation.metal_area,
        'Масса детали': calculation.mass,
        'Цена детали': calculation.detail_price,
        'Резка, м.п': calculation.cutting,
        'Врезка, количество': calculation.inset_amount,
        'Цена врезки': calculation.full_inset_price,
        'Цена, рез+врезка': calculation.full_cutting_price,
        'Количетво деталей': calculation.details_amount,
        'Стоимость деталей': calculation.complect_price,
        'Количетво комплектов': calculation.complects_amount,
        'Стоимость комплектов': calculation.final_price
        }
        self.calculation_df = pd.DataFrame(dict, index=[0])
        
    def _add_calculation_to_accounts(self) -> None:
        """
        Объединяет датафреймы из таблицы с учетом
        расчетов и с новым расчетом.
        """
        try:
            accounts_df = pd.read_excel(self.accounting_table_path, index_col=0)
            self._accounts_df = pd.concat([accounts_df, self.calculation_df]).reset_index(drop=True)
            self._accounts_df.index = self._accounts_df.index + 1
            
        except FileNotFoundError:
            self._accounts_df = self.calculation_df
            
        except Exception as e:
            raise Exception(f"Ошибка сохранения.\n{e}")
    
    def _set_accounting_table_path(self, accounting_table_path: str) -> None:
        """
        Задает новый путь к таблице для учета
        расчетов.
        """
        self.accounting_table_path = accounting_table_path
    
    def set_accounting_table_path(self, accounting_table_path: str=None) -> None:
        """
        Зпускает изменение путя к таблице с учетом расчетов.
        Если в функцию не передается аргумент с путем, передает
        путь по умолчанию.
        """
        if accounting_table_path is None:
            accounting_table_path = DEFAULT_ACCOUNTS_TABLE
        if accounting_table_path != "":
            self._set_accounting_table_path(accounting_table_path)
        
    def save_calculations(self, calculation: object) -> None:
        """
        Запускает создание датафрейма из полей расчета
        и его сохранение в таблице с учетом расчетов.
        """
        self._create_calculation_df(calculation)
        self._add_calculation_to_accounts()
        
        try:
            self._accounts_df.to_excel(self.accounting_table_path)
            
        except FileNotFoundError:
            with open(self.accounting_table_path, 'w', encoding='utf-8'):
                self._accounts_df.to_excel(self.accounting_table_path)
                
        except Exception as e:
            raise Exception(f"Ошибка сохранения.\n{e}")        
        
    def create_doc_to_print(self, calculation: object, file_path: str) -> None:
        """
        Запускает создание датафрейма из полей расчета
        и создание из него файла для печати.
        """
        self._create_calculation_df(calculation)
        
        try:
            with open(file_path, 'w', encoding='utf-8'):
                self.calculation_df.to_excel(file_path)
        except Exception as e:
            raise Exception(f"Ошибка сохранения.\n{e}")