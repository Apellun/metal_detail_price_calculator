import PySimpleGUI as sg
from const import metal_categories_list, metal_thickness_list

class Interface:
    def __init__(self, manager):
        self.manager = manager
        self.metals_list = None
        self.main_window = None
    
    def exception_popup(self, exception: str) -> None:
        """
        Error popup.
        """
        sg.popup_no_buttons(
                        exception,
                        modal=True,
                        title="Error!",
                    )
      
    def success_popup(self, text: str, title: str="Success") -> None:
        """
        Success popup.
        """
        sg.popup_no_buttons(
                    text,
                    modal=True,
                    title=title,
                    ) 
                
    def settings(self) -> None:
        """
        Settings window, allows to set a different metal prices spreadsheet,
        set it also for cutting and inset, choose a different spreadsheet to
        save calculations. Also allows to set default paths for spreadsheets.
        """
        metal_prices_table_path, accounting_table_path = self.manager.get_file_paths()
        
        settings_layout = [
            [sg.Text('Metal prices spreadsheet:'), sg.InputText(metal_prices_table_path, size=(40,3), key='Metal prices spreadsheet path'), sg.FileBrowse("Choose other", file_types=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")), key="Metals spreadsheet")],
            [sg.Checkbox('Also set for cutting and inset', key='Save for all')],
            [sg.Text('Spreadsheet for saving accounts:'), sg.InputText(accounting_table_path, size=(40,3), key='Spreadsheet for saving accounts path'), sg.FileBrowse("Choose other", file_types=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")), key="Accounts spreadsheets")],
            [sg.Button('Save'), sg.Button('Default')]           
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
            
            elif event == "Save":
                try:
                    self.manager.save_settings(values['Metals spreadsheet'], values['Accounts spreadsheets'], values['Save for all'])
                    self.main_window['Metal type'].update(values=self.manager.get_metals_list())
                    self.success_popup("Settings updated.")
                except Exception as e:
                    self.exception_popup(e)
                    
            elif event == "Default":
                self.manager.save_settings(save_for_all=True)
                self.main_window['Metal type'].update(values=self.manager.get_metals_list())
                metal_prices_table_path, accounting_table_path = self.manager.get_file_paths()
                settings_window['Metal prices spreadsheet path'].update(metal_prices_table_path)
                settings_window['Spreadsheet for saving accounts path'].update(accounting_table_path)
                self.success_popup("Default settings are set.")
                
        settings_window.close()
        
    def show_prices(self) -> None:
        """
        Shows a popup with prices.
        """
        prices_message = self.manager.create_prices_message()
        self.success_popup(prices_message, "Price")
        
    def save_calculations(self) -> None:
        """
        Saves the calculation data into a
        spreadsheet and shows a success popup.
        """
        self.manager.save_calculations()
        self.success_popup("Calculation saved")
    
    def overwrite_file_window(self, values: dict) -> None:
        """
        The window that appears if the user tries to rewrite an
        existing file. Allows to cancel or to proceed.
        """
        layout = [
        [sg.Text("File with this name already exists. Rewrite?")],
        [sg.Button("OK"), sg.Button("Cancel")]
        ]
        
        window = sg.Window(
            "Confirm action",
            layout
        )
        
        while True:
            event, _ = window.read()
            if event == sg.WINDOW_CLOSED or event == "Cancel":
                break
            elif event == "OK":
                self.manager.overvrite_file(values)
                break
            
        window.close()
        
    def save_to_print(self) -> None:
        """
        The window of saving a spreadsheet for a current
        calculation only. Saves it with a set name in a
        selected folder.
        """
        save_doc_to_print_layout = [
        [sg.Text('Choose name:'), sg.InputText(size=[10, 1], key='Name', enable_events=True), sg.InputText("Folder", key='Folder', size=[30, 1]), sg.FolderBrowse("Choose folder")],
        [sg.Button('Save')]
        ]
        
        save_doc_to_print_window = sg.Window(
        "Save file for print",
            save_doc_to_print_layout,
            default_element_size=[10],
            element_padding=10
        )
        
        while True:
            event, values = save_doc_to_print_window.read(timeout=1000)
            
            if event == sg.WIN_CLOSED:
                break
            
            elif event == "Save":
                try:
                    self.manager.save_doc_to_print(values)
                    self.success_popup("Saved")
                    break
                except FileExistsError:
                    self.overwrite_file_window(values)
                except Exception as e:
                    self.exception_popup(e)   
                    
        save_doc_to_print_window.close()
        
    def main(self) -> None:
        """
        The main interface window. Contains fields to select
        and input data for the calculation, creates a calculation
        when any of the buttons besides 'Settings' is pushed.
        Opens windows for all other operations upon pressing 
        buttons.
        """
        self.metals_list = self.manager.get_metals_list()
        
        main_layout = [
            [sg.Button('Settings')],
            [sg.Text('Detail name:'), sg.InputText(size=[50, 1], key="Detail name")],
            [sg.Text('Metal:'), sg.InputCombo(metal_categories_list, key="Metal category", readonly=True), sg.InputCombo(self.metals_list, key="Metal type", readonly=True)],
            [sg.Text('Metal thickness, mm:'), sg.InputCombo(metal_thickness_list, key="Metal thickness", readonly=True)],
            [sg.Text('Metal area sq. m:'), sg.InputText(key="Metal area", default_text="1")],
            [sg.Text('Cutting, running m:'), sg.InputText(key="Cutting", default_text="1")],
            [sg.Text('Inset amount:'), sg.InputText(key="Inset amount", default_text="1")],
            [sg.Text('Details amount:'), sg.InputText(key="Details amount", default_text="1")],
            [sg.Text('Complects amount:'), sg.InputText(key="Complects amount", default_text="1")],
            [sg.Column([[sg.Button('Show prices'), sg.Button('Save calculation'), sg.Button('Save file for print')]], justification='center')],
        ]
        
        main_window = sg.Window(
            "РMetal detail price calculator",
            main_layout,
            default_element_size=[10],
            element_padding=10
            )
        
        self.main_window = main_window
        
        while True:
            try:
                event, values = self.main_window.read()
                
                if event == sg.WIN_CLOSED:
                    break
                
                elif event == 'Settings':
                    self.settings()
                    
                else:
                    self.manager.create_calculation(values)
                    if event == 'Show prices':
                        self.show_prices()    
                    if event == 'Save calculation':
                        self.save_calculations()
                    if event == 'Save file for print':
                        self.save_to_print()
                        
            except Exception as e:
                self.exception_popup(e)
                
        self.main_window.close()
        
    def start(self) -> None:
        """
        Checks if the prices spreadsheet is in the
        default folder. If not, suggests the user to select
        a different spreadsheet and then proceeds to open
        the main window.
        """
        choose_spreadsheet_layout = [
            [sg.Text('Default spreadsheet with metal prices is not found, please select another one.', size=(40, 2)), sg.FileBrowse("Choose spreadsheet", file_types=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")), key="Metals spreadsheet")],
            [sg.Button('Save')]    
        ]
        
        choose_spreadsheet_window = sg.Window(
            "Choose spreadsheet",
            choose_spreadsheet_layout,
            modal=True,
            default_element_size=[10],
            element_padding=10
            )
        
        is_ready = True
        
        while True:
            try:
                self.metals_list = self.manager.get_metals_list()
                break
                
            except:
                event, values = choose_spreadsheet_window.read()
                
                if event == sg.WIN_CLOSED:
                    is_ready = False
                    break
                
                if event == "Save":
                    try:
                        self.manager.save_settings(metals_table_path=values['Metals spreadsheet'], save_for_all=True)
                        self.metals_list = self.manager.get_metals_list()
                        self.success_popup("Spreadsheet is set")
                        choose_spreadsheet_window.close()
                        break
                    except Exception as e:
                        self.exception_popup(f"Can't read spreadsheet\n{e}")
        
        if is_ready:       
            self.main()
        
        choose_spreadsheet_window.close()