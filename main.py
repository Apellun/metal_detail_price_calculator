from interface import Interface
from manager import Manager
from table_reader import TableReader
from saving import Saving

def create_app():
    """
    Инициализирует объекты классов и собирает из них приложение.
    """
    table_reader = TableReader()
    saving = Saving()
    manager = Manager(table_reader, saving)
    app = Interface(manager)
    return app
    
if __name__ == "__main__":
    app = create_app()
    app.run()