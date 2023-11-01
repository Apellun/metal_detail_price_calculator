from interface import Interface
from manager import Manager
from table_reader import TableReader
from saving import Saving

def create_app():
    table_reader = TableReader()
    saving = Saving()
    manager = Manager(table_reader, saving)
    interface = Interface(manager)
    return interface
    
if __name__ == "__main__":
    interface = create_app()
    interface.run()