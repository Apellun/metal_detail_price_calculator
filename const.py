import sys
from pathlib import Path

BASE_DIR_STR = Path(__file__).resolve().parent # для запуска из скрипта
# BASE_DIR_STR = Path(sys.executable).parent # для запуска из исполняемого файла
DEFAULT_PRICES_TABLE_PATH = BASE_DIR_STR / 'detail_making_costs_mock.xlsx'
DEFAULT_ACCOUNTS_TABLE = BASE_DIR_STR / 'accounts.xlsx'

metal_thickness_list = [0.5, 1, 1.5, 2, 3, 4, 5, 6, 8, 10, 12, 14, 16, 20]
metal_density_dict = {"кат1": 7.800, "кат2": 2.700, "кат3": 7.840}
metal_categories_list = list(metal_density_dict.keys())