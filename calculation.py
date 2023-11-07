from dataclasses import dataclass

@dataclass
class Calculation:
    blueprint_name: str
    metal_category: str
    metal_type: str
    metal_thickness: str
    metal_area: float
    metal_price: float
    cutting: int
    inset_amount: int
    cutting_price: float
    inset_price: float
    details_amount: int
    complects_amount: int
    density: int
    mass: float = None
    full_inset_price: float = None
    full_cutting_price: float = None
    detail_price: float = None
    complect_price: float = None
    final_price: float = None
    
    def __post_init__(self):
        self.mass = round((self.metal_thickness * self.metal_area * self.density), 4)
        self.detail_price = round((self.metal_price * self.metal_area), 2)
        self.full_inset_price = round((self.inset_price * self.inset_amount), 2)
        self.full_cutting_price = round((self.cutting_price * self.cutting + self.full_inset_price), 2)
        self.complect_price = round(((self.detail_price + self.full_cutting_price) * self.details_amount), 2)
        self.final_price = round((self.complect_price * self.complects_amount), 2)