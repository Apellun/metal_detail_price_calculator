from manager import Manager
from saving import Saving
from table_reader import TableReader

saving = Saving()
table_reader = TableReader()
manager = Manager(table_reader, saving)