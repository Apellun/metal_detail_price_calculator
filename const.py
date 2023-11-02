import sys
from pathlib import Path

BASE_DIR_STR = Path(__file__).resolve().parent # для запуска из скрипта
BASE_DIR_STR = Path(sys.executable).parent # для запуска из исполняемого файла
DEFAULT_PRICES_TABLE_PATH = BASE_DIR_STR / 'detail_making_costs.xlsx'
DEFAULT_ACCOUNTS_TABLE = BASE_DIR_STR / 'accounts.xlsx'
metal_categories_list = ["ст3", "нерж", "алюминий"]