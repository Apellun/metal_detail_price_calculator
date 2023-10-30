from pathlib import Path

BASE_DIR_STR = str(Path(__file__).resolve().parent)
DEFAULT_PRICES_TABLE_PATH = BASE_DIR_STR + '/detail_making_costs.xlsx'
DEFAULT_ACCOUNTS_TABLE = BASE_DIR_STR + '/accounts.xlsx'
metal_categories_list = ["ст3", "нерж", "алюминий"]