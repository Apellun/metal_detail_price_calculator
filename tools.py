def create_dict(values: dict) -> dict:
    """
    Собирает словарь с введенными данными для передачи в деталь.
    """
    return {
        'blueprint_name': values['Название чертежа'],
        'metal_category': values['Категория металла'],
        'metal_type': values['Тип металла'],
        'metal_thickness': float(values["Толщина металла"]),
        'metal_area': float(values["Площадь металла"]),
        'cutting': float(values["Резка, м. п."]),
        'in_cutting_amount': int(values["Врезка, количество"]),
        'details_amount': int(values["Количество деталей"]),
        'complects_amount': int(values["Количество комплектов"])
    }

def create_prices_str(detail: object) -> str:
    """
    Создает строку для вывода стоимости детали.
    """
    return (f"Стоимость металла для детали: {detail.detail_price}\n"
            f"Цена резки и врезки: {detail.full_cutting_price}\n"
            f"Цена одного комплекта деталей: {detail.complect_price}\n"
            f"Полная цена: {detail.final_price}")
    
def validate_fields(values: dict) -> None:
    """
    Проверяет значения в полях на отсутствие пустых строк
    и числовые значения там, где они требуются.
    """
    for index in values:
        value = values[index].strip()
        if value == "":
            raise ValueError(f"{index}: поле не может быть пустым.")
        if index not in ["Название чертежа", "Категория металла", "Тип металла"] and not value.isdigit():
            raise ValueError(f"{index}: введите числовое значение.")