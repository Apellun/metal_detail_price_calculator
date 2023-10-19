from table_reader import TableReader
from detail import Detail
from table_creator import TableCreator
from const import data, BASE_DIR_STR

class Main:
    
    def start(self, data):
        
        # создает деталь из переданных данных

        detail = Detail(**data)
        
        # находит путь к таблице
        
        costs_table_dir = BASE_DIR_STR + '/detail_making_costs.xlsx'
        
        # читает таблицы и добавляет в деталь стоимость металла, резки, врезки

        table_reader = TableReader(costs_table_dir)
        detail.metal_price = table_reader.get_metal_cost(detail.metal_type)
        detail.cutting_price = table_reader.get_cutting_cost(detail.metal_category, detail.metal_thickness, detail.details_amount)
        detail.in_cutting_price = table_reader.get_in_cutting_cost(detail.metal_thickness)
        
        # считает стоимость
        
        detail.set_prices()
        
        # выводит результат на экран

        print(f"Стоимость металла для детали: {detail.detail_price}\n"
                f"Цена резки и врезки: {detail.full_cutting_price}\n"
                f"Цена одного комплекта деталей: {detail.complect_price}\n"
                f"Полная цена: {detail.final_price}")
        
        create_table = input('Сохранить деталь? y/n\n')
        
        # создание таблиц (для этого нужно ввести игрек, но если раскладка русская, "н" тоже пойдет, регистр неважен)
        # / прерывание программы (в принципе любым другим сомволом)
        
        if create_table.lower() == 'y' or 'н':
            table_creator = TableCreator(detail)
            table_creator.save_detail()
            table_creator.create_doc_to_print()
        else:
            print('Прервано')

if __name__ == "__main__":
    main = Main()
    main.start(data)