import PySimpleGUI as sg
from const import metal_categories_list

class Interface:
    def __init__(self, manager: object):
        self.manager = manager
        self.metals_list = self.manager.get_metals_list()
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
            [sg.Button('Создать деталь и рассчитать стоимость'), sg.Button('Сохранить расчет'), sg.Button('Сохранить файл для печати')],
        ]
        self.main_window = sg.Window(
            "Расчет стоимости резки и гибки",
            self.main_layout,
            default_element_size=[10],
            element_padding=10,
            finalize=True
            )
    
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
            font=(60)
            ) 
          
    def settings(self) -> None:
        """
        Окно настроек, позволяет загрузить другую таблицу с ценами металлов,
        выбрать другую таблицу для созранения деталей, а также откатить эти
        изменения к настройкам по умолчанию.
        """
        metal_prices_table_path, accounting_table_path = self.manager.get_file_paths()
        settings_layout =[
            [sg.Text('Таблица со стоимостью металлов:'), sg.Text(metal_prices_table_path, size=(30,2), key='Путь к таблице металлов'), sg.FileBrowse("Выбрать другой", file_types=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")), key="Таблица металлов")],
            [sg.Text('Таблица для сохранения рассчетов:'), sg.Text(accounting_table_path, size=(30,2), key='Путь к таблице с рассчетами'), sg.FileBrowse("Выбрать другой", file_types=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")), key="Таблица рассчетов")],
            [sg.Button('Сохранить'), sg.Button('По умолчанию')]           
            ]
        settings_window = sg.Window(
            "Настройки",
            settings_layout,
            modal=True,
            default_element_size=[10],
            element_padding=10
            )
        while True:
            event, values = settings_window.read()
            if event == sg.WIN_CLOSED:
                break
            elif event == "Сохранить":
                try:
                    self.manager.save_settings(values['Таблица металлов'], values['Таблица рассчетов'])
                    self.main_window['Тип металла'].update(values=self.manager.get_metals_list())
                    self.success_popup("Настройки обновлены.")
                except Exception as e:
                    self.exception_popup(e)
            elif event == "По умолчанию":
                self.manager.save_settings()
                self.main_window['Тип металла'].update(values=self.manager.get_metals_list())
                metal_prices_table_path, accounting_table_path = self.manager.get_file_paths()
                settings_window['Путь к таблице металлов'].update(metal_prices_table_path)
                settings_window['Путь к таблице с рассчетами'].update(accounting_table_path)
                self.success_popup("Установлены настройки по умолчанию.")
        settings_window.close()
        
    def count_prices(self, values: dict) -> None:
        """
        Запускает создание детали и подсчет цен, показывает
        попап со стоимостью работ.
        """
        try:
            prices = self.manager.count_prices(values)
            self.success_popup(prices, "Стоимость")
        except Exception as e:
            self.exception_popup(e)
        
    def save_accounts(self) -> None:
        """
        Сохраняет информацию о расчете в таблицу
        и выводит попа успеха.
        """
        try:
            self.manager.save_accounts()
            self.success_popup("Расчет сохранен")
        except Exception as e:
            self.exception_popup(e)  
    
    def overwrite_file_window(self, values: dict) -> None:
        """
        Окно, которое возникает если пользователь перезаписывает существующий файл.
        Позволяет продолжить или отменить действие.
        """
        layout = [
        [sg.Text("Файл с таким именем уже существует. Перезаписать?")],
        [sg.Button("OK"), sg.Button("Отмена")]
        ]
        window = sg.Window(
            "Подтвердите действие",
            layout
        )
        while True:
            event, _ = window.read()
            if event == sg.WINDOW_CLOSED or event == "Отмена":
                break
            elif event == "OK":
                self.manager.overvrite_file(values)
                break
        window.close()
        
    def save_to_print(self) -> None:
        """
        Окно сохранения документа с данными только о текущем
        расчете для печати. Позволяет выбрать имя файла и папку,
        в которой сохранить файл.
        """
        save_doc_to_print_layout = [
        [sg.Text('Введите имя:'), sg.InputText(size=[10, 1], key='Имя', enable_events=True), sg.InputText("Папка для сохранения", key='Папка', size=[30, 1]), sg.FolderBrowse("Выбрать папку")],
        [sg.Button('Сохранить')]
        ]
        save_doc_to_print_window = sg.Window(
        "Сохранить для печати",
            save_doc_to_print_layout,
            default_element_size=[10],
            element_padding=10
        )
        while True:
            event, values = save_doc_to_print_window.read(timeout=1000)
            if event == sg.WIN_CLOSED:
                break
            elif event == 'Имя':
                if values['Имя'] == 'Введите имя':
                    save_doc_to_print_window['Имя'].update(default_text='')
            elif event == "Сохранить":
                try:
                    self.manager.save_doc_to_print(values)
                    self.success_popup("Файл сохранен")
                except FileExistsError:
                    self.overwrite_file_window(values)
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
                elif event == 'Создать деталь и рассчитать стоимость':
                    self.count_prices(values)    
                elif event == 'Сохранить расчет':
                    self.save_accounts()
                elif event == 'Сохранить файл для печати':
                    self.save_to_print()
                elif event == 'Настройки':
                    self.settings()
            self.main_window.close()
        except Exception as e:
            self.exception_popup(e)