# Файл, в который надо вбивать данные. Имхо, вбивать так проще —
# в форме диалога с консолью надо было бы дополнительно дописывать
# логику, чтобы можно было поправить какие-то неверно ввденные значения.

data = {
    'blueprint_name': 'Тестовый чертеж', # сюда вбивать имя детали
    'metal_category': 'ст3', # категория металла
    'metal_type': 'Лист г/к 2мм*1250*2500', # тип металла из списка в таблице на листе 'стоимость металлов'
    'metal_thickness': '2', # в ячейках таблиц — строка, поэтому тут тоже строка (всегда в кавычках)
    'metal_area': 0.01, # площадь металла
    'cutting': 1, # резка м. п.
    'in_cutting_amount': 2, # количество врезок
    'details_amount': 2, # количество деталей
    'complects_amount': 1 # количество комплектов
}