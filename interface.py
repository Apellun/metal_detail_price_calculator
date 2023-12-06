import PySimpleGUI as sg
from const import metal_categories_list, metal_thickness_list

class Interface:
    def __init__(self, manager):
        self.manager = manager
        self.metals_list = None
        self.main_window = None
    
    def exception_popup(self, exception: str) -> None:
        """
        Попап ошибки.
        """
        sg.popup_no_buttons(
                        exception,
                        modal=True,
                        title="Ошибка!"
                    )
      
    def success_popup(self, text: str, title: str="Успех") -> None:
        """
        Попап успеха.
        """
        sg.popup_no_buttons(
            text,
            modal=True,
            title=title
            ) 
                
    def settings(self) -> None:
        """
        Окно настроек, позволяет загрузить другую таблицу с ценами металлов,
        выбрать другую таблицу для сохранения расчетов, а также откатить эти
        изменения к настройкам по умолчанию.
        """
        metal_prices_table_path, accounting_table_path = self.manager.get_file_paths()
        settings_layout = [
            [sg.Text('Таблица со стоимостью металлов:'), sg.InputText(metal_prices_table_path, size=(40,3), key='Путь к таблице металлов'), sg.FileBrowse("Выбрать другой", file_types=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")), key="Таблица металлов")],
            [sg.Text('Таблица для сохранения рассчетов:'), sg.InputText(accounting_table_path, size=(40,3), key='Путь к таблице с рассчетами'), sg.FileBrowse("Выбрать другой", file_types=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")), key="Таблица рассчетов")],
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
        
    def show_prices(self) -> None:
        """
        Показывает попап со стоимостью работ.
        """
        try:
            prices_message = self.manager.create_prices_message()
            self.success_popup(prices_message, "Стоимость")
        except Exception as e:
            self.exception_popup(e)
        
    def save_calculations(self) -> None:
        """
        Сохраняет информацию о расчете в таблицу
        с расчетами и выводит попап успеха.
        """
        try:
            self.manager.save_calculations()
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
        расчете для печати. Сохраняет расчет в выбранной
        попке с заданным именем файла.
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
        
    def main(self) -> None:
        """
        Главное окно интерфейса. Создает расчет при
        нажатии на кнопки 'Показать стоимость',
        'Сохранить расчет', 'Сохранить файл для печати'.
        """
        main_layout = [
            [sg.Button('Настройки')],
            [sg.Text('Название чертежа:'), sg.InputText(size=[50, 1], key="Название чертежа")],
            [sg.Text('Металл:'), sg.InputCombo(metal_categories_list, key="Категория металла", readonly=True), sg.InputCombo(self.metals_list, key="Тип металла", readonly=True)],
            [sg.Text('Толщина металла:'), sg.InputCombo(metal_thickness_list, key="Толщина металла", readonly=True)],
            [sg.Text('Площадь металла:'), sg.InputText(key="Площадь металла", default_text="1")],
            [sg.Text('Резка, м. п.:'), sg.InputText(key="Резка, м. п.", default_text="1")],
            [sg.Text('Врезка, количество:'), sg.InputText(key="Врезка, количество", default_text="1")],
            [sg.Text('Количество деталей:'), sg.InputText(key="Количество деталей", default_text="1")],
            [sg.Text('Количество комплектов:'), sg.InputText(key="Количество комплектов", default_text="1")],
            [sg.Column([[sg.Button('Показать стоимость'), sg.Button('Сохранить расчет'), sg.Button('Сохранить файл для печати')]], justification='center')],
        ]
        window = sg.Window(
            "Расчет стоимости резки и гибки",
            main_layout,
            default_element_size=[10],
            element_padding=10
            )
        self.main_window = window
        while True:
            try:
                event, values = self.main_window.read()
                if event == sg.WIN_CLOSED:
                    break
                elif event == 'Настройки':
                    self.settings()
                else:
                    self.manager.create_calculation(values)
                    if event == 'Показать стоимость':
                        self.show_prices()    
                    if event == 'Сохранить расчет':
                        self.save_calculations()
                    if event == 'Сохранить файл для печати':
                        self.save_to_print()
            except Exception as e:
                self.exception_popup(e)
        self.main_window.close()
        
    def start(self) -> None:
        choose_spreadsheet_layout = [
            [sg.Text('Таблица со стоимостями металлов по умолчанию не найдена, выберите другую.'), sg.FileBrowse("Выбрать", file_types=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")), key="Таблица металлов")],
            [sg.Button('Сохранить')]    
        ]
        choose_spreadsheet_window = sg.Window(
            "Выбор таблицы",
            choose_spreadsheet_layout,
            modal=True,
            default_element_size=[10],
            element_padding=10
            )
        while True:
            try:
                self.metals_list = self.manager.get_metals_list()
            except:
                event, values = choose_spreadsheet_window.read()
                if event == sg.WIN_CLOSED:
                    break
                elif event == "Сохранить":
                    self.manager.save_settings(metals_table_path=values['Таблица металлов'])
                    self.metals_list = self.manager.get_metals_list()
                    self.success_popup("Таблица установлена")
                    self.main()
                    break