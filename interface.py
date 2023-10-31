import PySimpleGUI as sg
from table_reader import TableReader
from saving import Saving
from detail import Detail
from const import metal_categories_list
import tools

class Interface:
    def __init__(self):
        self.table_reader = TableReader()
        self.saving = Saving()
        self.metals_list = self.table_reader.get_metals_list()
        self.detail = None
        
        self.main_layout = [
            [sg.Button('Настройки')],
            [sg.Text('Название чертежа:'), sg.InputText(size=[50, 1], key="Название чертежа")],
            [sg.Text('Металл:'), sg.InputCombo(metal_categories_list, key="Категория металла", readonly=True), sg.InputCombo(self.metals_list, key="Тип металла", readonly=True)],
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
        
    def _create_detail(self, values: dict) -> None:
        """
        Создает деталь, заполняет поля значениями
        из таблиц и счиатет стоимость работы.
        """
        data = tools.create_dict(values)
        self.detail = Detail(**data)
        self.detail.metal_price = self.table_reader.get_metal_price(self.detail.metal_type)
        self.detail.cutting_price = self.table_reader.get_cutting_price(self.detail.metal_category, self.detail.metal_thickness, self.detail.details_amount)
        self.detail.in_cutting_price = self.table_reader.get_in_cutting_price(self.detail.metal_thickness)
        self.detail.set_prices()
    
    def exception_popup(self, exception: str) -> None:
        """
        Попап ошибки.
        """
        sg.popup_no_buttons(
                        exception,
                        modal=True,
                        title="Ошибка!",
                        font=(90)
                    )
      
    def success_popup(self, text: str, title: str="Успех") -> None:
        """
        Попап успеха.
        """
        sg.popup_no_buttons(
            text,
            modal=True,
            title=title,
            font=(90)
            ) 
        
    def count_prices(self, values: dict) -> None:
        """
        Запускает создание детали и подсчет цен, показывает
        попап со стоимостью работ.
        """
        try:
            tools.validate_fields(values)
            self._create_detail(values)
            prices_str = tools.create_prices_str(self.detail)
            self.success_popup(prices_str, "Стоимость")
        except Exception as e:
            self.exception_popup(e)
        
    def save_detail(self) -> None:
        """
        Сохраняет информацию о детали в таблицу.
        """
        self.saving.detail = self.detail
        try:
            self.saving.save_detail()
            self.success_popup("Деталь сохранена")
        except Exception as e:
            self.exception_popup(e)  
          
    def settings(self) -> None:
        """
        Окно настроек, позволяет загрузить другую таблицу с ценами металлов,
        выбрать другую таблицу для созранения деталей, а также откатить все
        изменения к настройкам по умолчанию.
        """
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
                self.main_window['Тип металла'].update(values=self.table_reader.get_metals_list())#TODO: refactor??
                break
            
            elif event == "По умолчанию":
                self.table_reader.set_tables()
                self.saving.set_accounting_table_path()
                self.main_window['Тип металла'].update(values=self.table_reader.get_metals_list())
                break
            
        settings_window.close()
        
    def save_to_print(self) -> None:
        """
        Окно сохранения документа с данными только о текущей
        детали для печати. Позволяет выбрать имя и папку, в
        которой сохранить файл.
        """
        save_doc_to_print_layout = [
        [sg.Text('Введите имя:'), sg.InputText(size=[20, 1], key='Имя'), sg.InputText(key='Папка'), sg.FolderBrowse("Выберите папку для сохранения")],
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
            event, values = save_doc_to_print_window.read(timeout=1000)
            if event == sg.WIN_CLOSED:
                break
            elif event == "Сохранить":
                try:
                    file_name, file_folder = values['Имя'], values['Папка']
                    if file_folder == '':
                        raise Exception("Выберите папку для сохранения.")
                    if file_name == '':
                        raise Exception("Укажите имя файла.")
                    self.saving.create_doc_to_print(file_folder, file_name)
                    self.success_popup("Файл сохранен")
                except Exception as e:
                    self.exception_popup(e)   
        save_doc_to_print_window.close()
        
    def run(self) -> None:
        """
        Главное окно.
        """
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