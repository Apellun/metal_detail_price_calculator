from dataclasses import dataclass

@dataclass
class Detail:
    blueprint_name: str
    metal_category: str
    metal_type: str
    metal_thickness: str
    metal_area: float
    cutting: int
    in_cutting_amount: int
    details_amount: int
    complects_amount: int
    metal_price: float = NotImplemented
    cutting_price: float = NotImplemented
    in_cutting_price: float = NotImplemented
    full_in_cutting_price: float = NotImplemented
    full_cutting_price: float = NotImplemented
    detail_price: float = NotImplemented
    complect_price: float = NotImplemented
    final_price: float = NotImplemented
    
    def _set_detail_price(self) -> None:
        """
        Считает и сохраняет стоимость металла для детали.
        """
        self.detail_price = self.metal_price * self.metal_area
        
    def _set_full_in_cutting_price(self) -> None:
        """
        Считает и сохраняет стоимость врезки для детали.
        """
        self.full_in_cutting_price = self.in_cutting_price * self.in_cutting_amount
    
    def _set_full_cutting_price(self) -> None:
        """
        Считает и сохраняет полную стоимость резки для детали.
        """
        self.full_cutting_price = self.cutting_price * self.cutting + self.full_in_cutting_price
    
    def _set_complect_price(self) -> None:
        """
        Считает и сохраняет стоимость комплекта деталей.
        """
        self.complect_price = (self.detail_price + self.full_cutting_price) * self.details_amount
        
    def _set_final_price(self) -> None:
        """
        Считает и сохраняет полную стоимость.
        """
        self.final_price = self.complect_price * self.complects_amount

    def set_prices(self):
        """
        Запускате подсчеты.
        """
        self._set_detail_price()
        self._set_full_in_cutting_price()
        self._set_full_cutting_price()
        self._set_complect_price()
        self._set_final_price()