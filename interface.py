import PySimpleGUI as sg
from table_reader import TableReader
from saving import Saving
from detail import Detail
from const import metal_categories_list

data={}

class Interface:
    def __init__(self):
        self.table_reader = TableReader()
        self.saving = Saving()
        self.metals_list = self.table_reader.get_metals_list()
        self.detail = None
        
        self.main_layout = [
            [sg.Button('Настройки')],
            [sg.Text('Название чертежа:'), sg.InputText(size=[50, 1], key="Название чертежа")],
            [sg.Text('Металл:'), sg.InputCombo(metal_categories_list, key="Категория металла"), sg.InputCombo(self.metals_list, key="Тип металла")],
            [sg.Text('Толщина металла:'), sg.InputText(key="Толщина металла")],
            [sg.Text('Площадь металла:'), sg.InputText(key="Площадь металла")],
            [sg.Text('Резка, м. п.:'), sg.InputText(key="Резка, м. п.")],
            [sg.Text('Врезка, количество:'), sg.InputText(key="Врезка, количество")],
            [sg.Text('Количество деталей:'), sg.InputText(key="Количество деталей")],
            [sg.Text('Количество комплектов:'), sg.InputText(key="Количество комплектов")],
            [sg.Button('Создать деталь и посчитать стоимость'), sg.Button('Сохранить деталь'), sg.Button('Сохранить файл для печати')],
        ]
        
        self.main_window = sg.Window(
            "Расчет стоимости резки и гибки",
            self.main_layout,
            default_element_size=[10],
            element_padding=10,
            font=(60),
            finalize=True
            )
        
    def _validate_fields(self, values):
        for index in values:
            value = values[index].strip()
            if value == "":
                raise ValueError(f"{index}: поле не может быть пустым.")
            if index not in ["Название чертежа", "Категория металла", "Тип металла"] and not value.isdigit():
                raise ValueError(f"{index}: введите числовое значение.")
        
    def _create_dict(self, values):
        return {
            'blueprint_name': values['Название чертежа'],
            'metal_category': values['Категория металла'],
            'metal_type': values['Тип металла'],
            'metal_thickness': float(values["Толщина металла"]),
            'metal_area': float(values["Площадь металла"]),
            'cutting': float(values["Резка, м. п."]),
            'in_cutting_amount': int(values["Врезка, количество"]),
            'details_amount': int(values["Количество деталей"]),
            'complects_amount': int(values["Количество комплектов"])
        }
        
    def _count_prices(self, values):
        data = self._create_dict(values)
        self.detail = Detail(**data)
        self.detail.metal_price = self.table_reader.get_metal_price(self.detail.metal_type)
        self.detail.cutting_price = self.table_reader.get_cutting_price(self.detail.metal_category, self.detail.metal_thickness, self.detail.details_amount)
        self.detail.in_cutting_price = self.table_reader.get_in_cutting_price(self.detail.metal_thickness)
        self.detail.set_prices()
    
    def _create_prices_str(self):
        return (f"Стоимость металла для детали: {self.detail.detail_price}\n"
                f"Цена резки и врезки: {self.detail.full_cutting_price}\n"
                f"Цена одного комплекта деталей: {self.detail.complect_price}\n"
                f"Полная цена: {self.detail.final_price}")
    
    def exception_popup(self, exception):
        sg.popup_no_buttons(
                        exception,
                        modal=True,
                        title="Ошибка!",
                        font=(90)
                    )
      
    def success_popup(self, text):
        sg.popup_no_buttons(
            text,
            modal=True,
            title="Успех",
            font=(90)
            ) 
          
    def settings(self):
        settings_layout =[
            [sg.Text('Файл со стоимостью металлов:'), sg.Input(readonly=True), sg.FileBrowse("Выбрать", file_types=(("Excel files", "*.xlsx"),))],
            # [sg.Text('Файл со стоимостью резки и врезки:'), sg.Input(readonly=True), sg.FileBrowse("Выбрать", file_types=(("Excel files", "*xlsx"),))],
            [sg.Text('Файл с таблицей для сохранения детали:'), sg.Input(readonly=True), sg.FileBrowse("Выбрать", file_types=(("Excel files", "*.xlsx"),))],
            [sg.Button('Сохранить'), sg.Button('По умолчанию')]           
            ]
        settings_window = sg.Window(
            "Настройки",
            settings_layout,
            modal=True,
            default_element_size=[10],
            element_padding=10,
            font=(60)
            )
        
        while True:
            event, values = settings_window.read()
            if event == sg.WIN_CLOSED:
                break
            elif event == "Сохранить":
                try:
                    self.table_reader.set_tables(values[0])
                    self.saving.set_accounting_table_path(values[1])
                except Exception as e:
                    self.exception_popup(e)
                self.main_window['Тип металла'].update(values=[value for value in self.table_reader.get_metals_list()])#TODO: refactor??
                break
            elif event == "По умолчанию":
                self.table_reader.set_tables()
                self.saving.set_accounting_table_path()
                self.main_window['Тип металла'].update(values=[value for value in self.table_reader.get_metals_list()])
                break 
        settings_window.close()
        
    def count_prices(self, values):
        try:
            self._validate_fields(values)
            self._count_prices(values)
            res = self._create_prices_str()
            sg.popup_no_buttons(
                res,
                title="Стоимость",
                font=(90)
                )
        except Exception as e:
            self.exception_popup(e)
        
    def save_detail(self):
        self.saving.detail = self.detail
        try:
            self.saving.save_detail()
            self.success_popup("Деталь сохранена")
        except Exception as e:
            self.exception_popup(e)  
        
    def save_to_print(self):
        save_doc_to_print_layout = [
        [sg.Text('Сохранить как:'), sg.Input(key='-PATH-', readonly=True), sg.SaveAs("Выбрать", file_types=(('*.xls', '*.xlsx'),))],
        [sg.Button('Сохранить')]
        ]
        save_doc_to_print_window = sg.Window(
        "Сохранить для печати",
            save_doc_to_print_layout,
            default_element_size=[10],
            element_padding=10,
            font=(60)
        )
        
        self.saving.detail = self.detail
        while True:
            event, values = save_doc_to_print_window.read()
            if event == sg.WIN_CLOSED:
                break
            elif event == "Сохранить":
                try:
                    file_path = values['-PATH-']
                    if file_path == '':
                        raise Exception("Выберите имя и папку для сохранения.")
                    if not file_path.lower().endswith((".xlsx", ".xls")):
                        file_path += ".xlsx"
                        with open(file_path, 'w', encoding='utf-8'):
                            self.saving.create_doc_to_print(file_path)
                        self.success_popup("Файл сохранен")
                except Exception as e:
                    self.exception_popup(e)
        save_doc_to_print_window.close()
        
    def run(self):
        try:
            while True:
                event, values = self.main_window.read()
                if event == sg.WIN_CLOSED:
                    break
                elif event == 'Создать деталь и посчитать стоимость':
                    self.count_prices(values)    
                elif event == 'Сохранить деталь':
                    self.save_detail()
                elif event == 'Сохранить файл для печати':
                    self.save_to_print()
                elif event == 'Настройки':
                    self.settings()
            self.main_window.close()
        except Exception as e:
            self.exception_popup(e)